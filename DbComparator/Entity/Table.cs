using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace DbComparator.Entity;

public class Table
{
    public readonly string TableName;
    public readonly string TableSchema;
    public readonly string TableType;
    public readonly HashSet<Column> Columns = new HashSet<Column>();
    public readonly HashSet<Index> Indexes = new HashSet<Index>();
    public Table(string tableName, string tableSchema, string tableType)
    {
        TableName = tableName;
        TableSchema = tableSchema;
        TableType = tableType;
    }

    public void ReadColumns(SqlConnection connection)
    {
        using (SqlCommand command = new SqlCommand($"SELECT TABLE_SCHEMA, table_name as [Table], column_name as 'Column Name', data_type as 'Data Type', character_maximum_length as 'Max Length' FROM information_schema.columns c WHERE TABLE_SCHEMA = '{TableSchema}' AND TABLE_NAME = '{TableName}'", connection))
        {
            using (SqlDataReader sqlDataReader = command.ExecuteReader())
            {
                while (sqlDataReader.Read())
                {
                    string columnName = sqlDataReader["Column Name"].ToString();
                    string dataType = sqlDataReader["Data Type"].ToString();
                    long? maxLength = sqlDataReader["Max Length"] is DBNull ? null : (long?)Convert.ToInt64(sqlDataReader["Max Length"]);

                    Column column = new Column(columnName, dataType, maxLength);
                    Columns.Add(column);
                }
            }
        }
    }

    public bool IsIndexingSame(Table? other)
    {
        if (other == null)
            return false;

        foreach (Column columnInSource in this.Columns)
        {
            Column? columnInTarget = other.FindColumn(columnInSource.ColumnName);
            if (columnInTarget == null)
                return false;

            if (!columnInSource.IsIndexSame(columnInTarget))
                return false;
        }

        return true;
    }
    
    public string GetFormattedTableName()
    {
        return $"[{TableSchema}].[{TableName}]";
    }

    public string GetTableTypeFormatted()
    {
        return TableType == "BASE TABLE" ? "Table" : TableType;
    }

    public Column? FindColumn(string name)
    {
        return Columns.FirstOrDefault(x => x.ColumnName == name);
    }
    
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + (TableName != null ? TableName.GetHashCode() : 0);
            hash = hash * 23 + (TableSchema != null ? TableSchema.GetHashCode() : 0);
            hash = hash * 23 + (TableType != null ? TableType.GetHashCode() : 0);
            foreach (var column in Columns)
            {
                hash = hash * 23 + (column != null ? column.GetHashCode() : 0);
            }
            return hash;
        }
    }

    public override bool Equals(object obj)
    {
        if (obj == null || GetType() != obj.GetType())
        {
            return false;
        }

        Table other = (Table)obj;

        return (string.Equals(TableName, other.TableName) &&
                string.Equals(TableSchema, other.TableSchema) &&
                string.Equals(TableType, other.TableType) &&
                Columns.SetEquals(other.Columns));
    }
}