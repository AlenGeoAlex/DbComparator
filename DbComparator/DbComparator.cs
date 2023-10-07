using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DbComparator.Entity;
using DocumentFormat.OpenXml.Office2013.Word;
using Microsoft.Data.SqlClient;

using Column = DbComparator.Entity.Column;
using Index = DbComparator.Entity.Index;
using Table = DbComparator.Entity.Table;

namespace DbComparator;

public class DbComparator
{
    //Only 2 queries are getting executed that too synchronously, we can reuse the connection.
    private readonly string _sourceDbConnectionString;
    private readonly string _targetDbConnectionString;
    private readonly ComparisonSettings _settings;

    private SqlConnection _sourceDbConnection;
    private SqlConnection _targetDbConnection;

    private Database _source, _target;
    

    public DbComparator(string sourceDbConnectionString, string targetDbConnectionString, ComparisonSettings settings)
    {
        _sourceDbConnectionString = sourceDbConnectionString ??
                                   throw new ArgumentNullException(nameof(sourceDbConnectionString));
        _targetDbConnectionString = targetDbConnectionString ??
                                   throw new ArgumentNullException(nameof(targetDbConnectionString));
        _settings = settings;
    }
    

    public bool TryConnectSource()
    {
        var connection = createConnection(_sourceDbConnectionString);
        if (connection == null)
        {
            Console.WriteLine("Failed to connect to Source Database");
            return false;
        }
        
        _sourceDbConnection = connection;
        _sourceDbConnection.Open();
        return _sourceDbConnection.State == ConnectionState.Open;
    }

    public bool TryConnectTarget()
    {
        var connection = createConnection(_targetDbConnectionString);
        if (connection == null)
        {
            Console.WriteLine("Failed to connect to Target Database");
            return false;
        }
        
        _targetDbConnection = connection;
        _targetDbConnection.Open();
        return _targetDbConnection.State == ConnectionState.Open;
    }

    public void GetDataFromSource()
    {
        _source = new Database(this._sourceDbConnection);
        _source.GetTables();
        foreach (Table sourceTable in _source.Tables)
        {
            sourceTable.ReadColumns(this._sourceDbConnection);
        }

        if (_settings.CheckIndex)
        {
            _source.GetIndexes();
        }
    }

    public void GetDataFromTarget()
    {
        _target = new Database(this._targetDbConnection);
        _target.GetTables();
        foreach (Table table in _target.Tables)
        {
            table.ReadColumns(this._targetDbConnection);
        }
        if (_settings.CheckIndex)
        {
            _target.GetIndexes();
        }
    }
    private SqlConnection? createConnection(string connectionString)
    {
        try
        {
            SqlConnection connection = new SqlConnection(connectionString);
            return connection;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return null;
        }
    }

    public void Disconnect()
    {
        DisconnectSource();
        DisconnectTarget();
    }

    public void DisconnectSource()
    {
        try
        {
            if (_sourceDbConnection.State != ConnectionState.Closed)
            {
                _sourceDbConnection.Close();
                _sourceDbConnection.Dispose();
            }

            Console.WriteLine("Source Db Connection is closed");
        }
        catch (Exception ignored)
        {
            Console.WriteLine(ignored);
        }
    }

    public void DisconnectTarget()
    {
        try
        {
            if (_targetDbConnection.State != ConnectionState.Closed)
            {
                _targetDbConnection.Close();
                _sourceDbConnection.Dispose();
            }
            Console.WriteLine("Target Db Connection is closed");
        }
        catch (Exception ignored)
        {
            Console.WriteLine(ignored);
        }
    }

    public void Analyze()
    {
        List<string> missingTables = new List<string>();
        List<string> columnDifferences = new List<string>();
        Console.WriteLine("Statistics");
        Console.WriteLine("-------------------------");
        Console.WriteLine($"  - No of tables in Source  --- No of tables in Target");
        Console.WriteLine($"  -  {_source.Tables.Count}                    ---  {_target.Tables.Count}");
        Console.WriteLine($"  - No of schemas in Source  --- No of schemas in Target");
        Console.WriteLine($"  -  {_source.Tables.GroupBy(x => x.TableSchema).Count()}                    ---  {_target.Tables.GroupBy(x => x.TableSchema).Count()}");
        
        Console.WriteLine("Diff");
        Console.WriteLine("-------------------------");
        foreach (Table sourceTable in _source.Tables)
        {
            Console.WriteLine($"Analyzing {sourceTable.GetTableTypeFormatted()} "+sourceTable.GetFormattedTableName()+"...");
            bool hasChanges = false;
            var targetTable = _target.FindTable(sourceTable.TableName, sourceTable.TableSchema);
            if (targetTable == null)
            {
                Console.WriteLine($" - {sourceTable.GetFormattedTableName()} is missing on Target Database");
                missingTables.Add(sourceTable.TableName);
                hasChanges = true;
                continue;
            }
            
            foreach (Column sourceTableColumn in sourceTable.Columns)
            {
                Column? targetTableColumn = targetTable.FindColumn(sourceTableColumn.ColumnName);
                if (targetTableColumn == null)
                {
                    string message =
                        $"A column is missing on on Target table {targetTable.GetFormattedTableName()}. The missing column is {sourceTableColumn.ColumnName}";       
                    Console.WriteLine($"  -  - {message}");
                    columnDifferences.Add(message);
                    hasChanges = true;
                    continue;
                }

                if (sourceTableColumn.DataType != targetTableColumn.DataType)
                {
                    string message =
                        $"The data type has been changed for column {sourceTableColumn.ColumnName} on table {targetTable.GetFormattedTableName()} from {sourceTableColumn.DataType} [Source] to {targetTableColumn.DataType} [Target]";
                    Console.WriteLine($"  -  - {message}");
                    hasChanges = true;
                    columnDifferences.Add(message);
                }
                
                if (sourceTableColumn.MaxLength != targetTableColumn.MaxLength)
                {
                    string message =
                        $"The max length has been changed for column {sourceTableColumn.ColumnName} on table {targetTable.GetFormattedTableName()} from {sourceTableColumn.MaxLength} [Source] to {targetTableColumn.MaxLength} [Target]";
                    Console.WriteLine($"  -  - {message}");
                    hasChanges = true;
                    columnDifferences.Add(message);
                }
            }

            if (!hasChanges)
            {
                Console.WriteLine("No difference spotted on source and target");
            }
        }
    }
    
    public void CreateReport()
    {
        string workingDirectory = Environment.CurrentDirectory;
        string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss", CultureInfo.CurrentCulture);
        string reportFileName = $"report_{currentDateTime}.xlsx";
        string reportFilePath = Path.Combine(workingDirectory, reportFileName);

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Report");
            worksheet.SheetView.FreezeRows(1);

            // Add headers
            worksheet.Cell(1, 1).Value = "SourceSchema";
            worksheet.Cell(1, 2).Value = "TargetSchema";
            worksheet.Cell(1, 3).Value = "SourceTableName";
            worksheet.Cell(1, 4).Value = "TargetTableName";
            worksheet.Cell(1, 5).Value = "SourceColumnName";
            worksheet.Cell(1, 6).Value = "TargetColumnName";
            worksheet.Cell(1, 7).Value = "SourceDataType";
            worksheet.Cell(1, 8).Value = "TargetDataType";
            worksheet.Cell(1, 9).Value = "SourceLength";
            worksheet.Cell(1, 10).Value = "TargetLength";
            worksheet.Cell(1, 11).Value = "SourceIndex";
            worksheet.Cell(1, 12).Value = "TargetIndex";
            worksheet.Cell(1, 13).Value = "Change";
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.Beige;
            worksheet.Row(1).Style.Font.Bold = true;
            for (int i = 1; i <= 12; i++)
            {
                worksheet.Column(i).Width = 35;
            }
            int row = 3;

            foreach (Table sourceTable in _source.Tables.OrderBy(x => x.TableSchema).ThenBy(x => x.TableName))
            {
                Table? targetTable = _target.FindTable(sourceTable.TableName, sourceTable.TableSchema);
                foreach (Column sourceTableColumn in sourceTable.Columns)
                {
                    Column? targetColumn = targetTable?.FindColumn(sourceTableColumn.ColumnName);
                    worksheet.Cell(row, 1).Value = sourceTable.TableSchema;
                    worksheet.Cell(row, 2).Value = targetTable?.TableSchema ?? string.Empty;
                    worksheet.Cell(row, 3).Value = sourceTable.TableName;
                    worksheet.Cell(row, 4).Value = targetTable?.TableName ?? string.Empty;
                    worksheet.Cell(row, 5).Value = sourceTableColumn.ColumnName;
                    worksheet.Cell(row, 6).Value = targetColumn?.ColumnName ?? string.Empty;
                    worksheet.Cell(row, 7).Value = sourceTableColumn.DataType.ToUpper();
                    worksheet.Cell(row, 8).Value = targetColumn?.DataType.ToUpper() ?? string.Empty;
                    worksheet.Cell(row, 9).Value = sourceTableColumn.MaxLength.HasValue ? Convert.ToString(sourceTableColumn.MaxLength.Value) : string.Empty;
                    worksheet.Cell(row, 10).Value = (targetColumn is { MaxLength: not null } ? Convert.ToString(targetColumn?.MaxLength.Value) : string.Empty) ?? string.Empty;
                    worksheet.Cell(row, 11).Value = _settings.CheckIndex ? sourceTableColumn.GetIndexAsFormattedString() : string.Empty;
                    worksheet.Cell(row, 12).Value = _settings.CheckIndex ? targetColumn?.GetIndexAsFormattedString() ?? string.Empty : string.Empty;
                    worksheet.Cell(row, 13).Value = "No";
                    if (!AreColumnsEqual(sourceTableColumn, targetColumn))
                    {
                        worksheet.Cell(row, 13).Value = "Yes";
                        worksheet.Row(row).Style.Fill.BackgroundColor = XLColor.Aqua;
                        worksheet.Row(row).Style.Font.Bold = true;
                    }

                    row++;
                }

                row++;
            }

            var statistics = workbook.Worksheets.Add("Statistics");
            statistics.SheetView.FreezeRows(1);
            statistics.Cell(1, 1).Value = "SourceSchema";
            statistics.Cell(1, 2).Value = "TargetSchema";
            statistics.Cell(1, 3).Value = "SourceTableName";
            statistics.Cell(1, 4).Value = "TargetTableName";
            statistics.Cell(1, 5).Value = "SourceColumnCount";
            statistics.Cell(1, 6).Value = "TargetColumnCount";
            statistics.Cell(1, 7).Value = "Change";
            statistics.Row(1).Style.Fill.BackgroundColor = XLColor.Beige;
            statistics.Row(1).Style.Font.Bold = true;
            for (int i = 1; i <= 7; i++)
            {
                statistics.Column(i).Width = 35;
            }
            int statRow = 3;
            foreach (Table sourceTable in _source.Tables.OrderBy(x => x.TableSchema).ThenBy(x => x.TableName))
            {
                Table? targetTable = _target.FindTable(sourceTable.TableName, sourceTable.TableSchema);
                statistics.Cell(statRow, 1).Value = sourceTable.TableSchema;
                statistics.Cell(statRow, 2).Value = targetTable?.TableSchema ?? string.Empty;
                statistics.Cell(statRow, 3).Value = sourceTable.TableName;
                statistics.Cell(statRow, 4).Value = targetTable?.TableName ?? string.Empty;
                statistics.Cell(statRow, 5).Value = sourceTable.Columns.Count();
                statistics.Cell(statRow, 6).Value = targetTable?.Columns.Count();
                bool isChanged = sourceTable.Columns.Count() != targetTable?.Columns.Count();
                statistics.Cell(statRow, 7).Value = isChanged ? "Yes" : "No";
                if(isChanged)
                    statistics.Row(statRow).Style.Fill.BackgroundColor = XLColor.Ruby;
                statRow++;
            }

            if (!_settings.CheckIndex)
            {
                workbook.SaveAs(reportFilePath);
                return;
            }

            var indexSheet = workbook.Worksheets.Add("Index");
            int indexSheetRow = 2;
            int maxColCount = 0;
            foreach (Table sourceTable in _source.Tables.OrderBy(x => x.TableSchema).ThenBy(x => x.TableName))
            {
                indexSheet.Row(indexSheetRow).Style.Fill.BackgroundColor = XLColor.Black;
                indexSheet.Row(indexSheetRow).Style.Font.Bold = true;
                indexSheet.Cell(indexSheetRow, 1).Value = sourceTable.GetFormattedTableName();
                indexSheet.Cell(indexSheetRow, 1).Style.Font.FontColor = XLColor.Awesome;
                indexSheetRow++;
                Table? targetTable = _target.FindTable(sourceTable.TableName, sourceTable.TableSchema);
                if (targetTable == null)
                {
                    indexSheetRow++;
                    continue;
                }

                bool isFirstColumn = true;
                foreach (Column sourceTableColumn in sourceTable.Columns)
                {
                    var targetColumn = targetTable.FindColumn(sourceTableColumn.ColumnName);
                    int sourceRow = indexSheetRow;
                    int targetRow = sourceRow + 1;
                    if (!isFirstColumn){
                        indexSheet.Row(indexSheetRow).Style.Fill.BackgroundColor = XLColor.FromArgb(166, 166, 166);
                        indexSheetRow++;
                        sourceRow = indexSheetRow;
                        targetRow = sourceRow + 1;
                    }
                    
                    indexSheet.Cell(sourceRow, 1).Value = sourceTableColumn.ColumnName;
                    indexSheet.Cell(targetRow, 1).Value = targetColumn?.ColumnName ?? string.Empty;

                    int colCount = 3;
                    var sourceChangeRow = indexSheet.Cell(sourceRow, 2);
                    sourceChangeRow.Value = "No";
                    var targetChangeRow = indexSheet.Cell(targetRow, 2);
                    targetChangeRow.Value = "No";
                    foreach (var sourceIndex in sourceTableColumn.Indexes)
                    {
                        Index? targetIndex = targetColumn?.FindIndexOfType(sourceIndex);
                        var sourceCell = indexSheet.Cell(sourceRow, colCount);
                        var targetCell = indexSheet.Cell(targetRow, colCount);
                        colCount++;
                        XLColor color = sourceIndex.IsInclude ? XLColor.Khaki : XLColor.LightGreen;
                        XLColor textColor = XLColor.Black;
                        if (sourceIndex.IndexType == "NONCLUSTERED")
                        {
                            textColor = XLColor.Red;
                        }else if (sourceIndex.IndexType == "CLUSTERED")
                        {
                            textColor = XLColor.Purple;
                        }
                        if (targetIndex == null)
                        {
                            sourceCell.Value = sourceIndex.IndexName;
                            targetCell.Value = "MISSING/INVALID IN TARGET";
                            color = XLColor.IndianRed;
                            sourceCell.Style.Fill.BackgroundColor = color;
                            targetCell.Style.Fill.BackgroundColor = color;
                            sourceCell.Style.Font.FontColor = XLColor.Black;
                            targetCell.Style.Font.FontColor = XLColor.Black;
                            sourceCell.Style.Font.Bold = true;
                            targetCell.Style.Font.Bold = true;
                            if (sourceIndex.IsInclude)
                                sourceCell.CreateComment()
                                    .AddText("This column is added as an include in source");

                            if (sourceChangeRow.Value.ToString() == "No" || targetChangeRow.Value.ToString() == "No")
                            {
                                sourceChangeRow.Value = "Yes";
                                targetChangeRow.Value = "Yes";
                                sourceChangeRow.Style.Font.Bold = targetChangeRow.Style.Font.Bold = true;
                                sourceChangeRow.Style.Fill.BackgroundColor = targetChangeRow.Style.Fill.BackgroundColor = XLColor.Aqua;
                            }
                            
                            continue;
                        }

                        if (sourceIndex.IsInclude)
                        {
                            color = XLColor.Khaki;
                            sourceCell.CreateComment().AddText("This column is added as an include in source");
                            targetCell.CreateComment().AddText("This column is added as an include in source");
                        }
                    

                        sourceCell.Value = sourceIndex.IndexName;
                        targetCell.Value = targetIndex.IndexName;
                        sourceCell.Style.Fill.BackgroundColor = color;
                        targetCell.Style.Fill.BackgroundColor = color;
                        sourceCell.Style.Font.FontColor = textColor;
                        targetCell.Style.Font.FontColor = textColor;
                    }
                    
                    isFirstColumn = false;
                    indexSheetRow = targetRow + 1;
                    if (sourceTableColumn.Indexes.Count > maxColCount)
                        maxColCount = sourceTableColumn.Indexes.Count;
                }

                indexSheetRow++;
            }

            indexSheet.Columns(3, maxColCount + 1).Width = 75;

            var indexHelpSheet = workbook.Worksheets.Add("IndexHelpSheet");

            indexHelpSheet.Cell(1, 1).Value =
                "If the text is red like this, it means the index is of NONCLUSTERED Type";
            indexHelpSheet.Cell(1, 1).Style.Font.FontColor = XLColor.Red;

            var purpleText = indexHelpSheet.Cell(2, 1);
            purpleText.Value = "If the text is purple like this, It means the index is of CLUSTERED Type";
            purpleText.Style.Font.FontColor = XLColor.Purple;

            var includeIndexHelpCell = indexHelpSheet.Cell(3, 1);
            includeIndexHelpCell.Value =
                "If the background of the cell is this color (A pale yellow), it means this column is added as an include in the index named here";
            includeIndexHelpCell.Style.Fill.BackgroundColor = XLColor.Khaki;

            var noChange = indexHelpSheet.Cell(4, 1);
            noChange.Value =
                "If the background of the cell is this color (A Pale/Light Green), it means, this column has no change in both source and target";
            noChange.Style.Fill.BackgroundColor = XLColor.LightGreen;

            var change = indexHelpSheet.Cell(5, 1);
            change.Value =
                "If the background of the cell is this color (red), It means this column has changes in the specified index, Either the index is missing, or something. If the index in the source is an include, check whether the source has a comment or not";
            change.Style.Fill.BackgroundColor = XLColor.IndianRed;

            indexHelpSheet.Column(1).Width = 250;

            workbook.SaveAs(reportFilePath);
        }

        Console.WriteLine($"Excel report file created at: {reportFilePath}");
    }

    private bool AreColumnsEqual(Column? sourceColumn, Column? targetColumn)
    {
        return sourceColumn?.ColumnName == targetColumn?.ColumnName &&
               sourceColumn?.DataType == targetColumn?.DataType &&
               sourceColumn?.MaxLength == targetColumn?.MaxLength &&
               (_settings.CheckIndex ? sourceColumn.IsIndexSame(targetColumn) : true);
    }
}