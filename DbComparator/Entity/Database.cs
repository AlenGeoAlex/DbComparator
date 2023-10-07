using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace DbComparator.Entity;

public class Database
{
    public readonly SqlConnection Connection;
    public HashSet<Table> Tables { get; private set; } = new HashSet<Table>(); 

    public Database(SqlConnection connection)
    {
        Connection = connection ?? throw new ArgumentNullException(nameof(connection));
    }

    public void GetTables()
    {
        if (Connection.State != ConnectionState.Open)
            throw new Exception("Database connection is closed");

        using (SqlCommand command = new SqlCommand("SELECT * FROM INFORMATION_SCHEMA.TABLES", Connection))
        {
            using (SqlDataReader sqlDataReader = command.ExecuteReader())
            {
                while (sqlDataReader.Read())
                {
                    string tableSchema = sqlDataReader["TABLE_SCHEMA"].ToString();
                    string tableName = sqlDataReader["TABLE_NAME"].ToString();
                    string tableType = sqlDataReader["TABLE_TYPE"].ToString();

                    if (tableName == null || tableType == null)
                    {
                        Console.WriteLine($"Skipping a possible empty table {tableSchema}.{tableName}.{tableType}");
                        continue;
                    }
                        
                    
                    Table table = new Table(tableName, tableSchema, tableType);
                    Tables.Add(table);
                }
            }
        }
    }

    public void GetIndexes()
    {
        if (Connection.State != ConnectionState.Open)
            throw new Exception("Database connection is closed");
        
        using (SqlCommand command = new SqlCommand(@"SELECT DISTINCT
                                                            TableName = t.Name,
                                                            IndexName = i.Name, 
                                                            IndexType = i.type_desc,
                                                            ColumnOrdinal = ic.key_ordinal,
                                                            ColumnName = c.name,
                                                            SchemaName = SCHEMA_NAME(t.schema_id)
                                                        FROM sys.indexes i
                                                        INNER JOIN sys.tables t ON t.object_id = i.object_id
                                                        INNER JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
                                                        INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
                                                        INNER JOIN sys.types ty ON c.system_type_id = ty.system_type_id
                                                        ORDER BY t.Name, i.name, ic.key_ordinal",
                   Connection))
        {
            using (SqlDataReader sqlDataReader = command.ExecuteReader())
            {
                while (sqlDataReader.Read())
                {
                    try
                    {
                        string tableName = sqlDataReader["TableName"].ToString();
                        string indexName = sqlDataReader["IndexName"].ToString();
                        string indexType = sqlDataReader["IndexType"].ToString();
                        int? columnOrdinal = Convert.ToInt32(sqlDataReader["ColumnOrdinal"].ToString());
                        string columnName = sqlDataReader["ColumnName"].ToString();
                        string tableSchema = sqlDataReader["SchemaName"].ToString();

                        if(tableName == null || indexName == null || indexType == null || !columnOrdinal.HasValue || columnName == null || tableSchema == null)
                            continue;

                        Table? table = FindTable(tableName, tableSchema);
                        if (table == null)
                        {
                            continue;
                        }

                        Column? column = table.FindColumn(columnName);
                        if (column == null)
                        {
                            continue;
                        }
                        column.CreateIndex(new Index(indexName, indexType, columnOrdinal.Value));
                    }
                    catch (Exception ignored)
                    {
                        // ignored
                    }
                }
            }
        }
    }
            
    public Table? FindTable(string name, string schema)
    {
        return Tables.FirstOrDefault(x => x.TableName == name && x.TableSchema == schema);
    }
}