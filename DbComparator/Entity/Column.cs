namespace DbComparator.Entity;

public class Column
{
    public readonly string ColumnName;
    public readonly string DataType;
    public long? MaxLength;
    public HashSet<Index> Indexes { get;  }

    public Column(string columnName, string dataType, long? maxLength)
    {
        ColumnName = columnName;
        DataType = dataType;
        MaxLength = maxLength;
        Indexes = new HashSet<Index>();
    }
    
    public void CreateIndex(Index index)
    {
        this.Indexes.Add(index);
    }

    public bool IsIndexSame(Column? other)
    {
        if (other == null)
            return false;

        foreach (Index indexInThisColumn in this.Indexes)
        {
            Index? indexInTargetColumn = other.FindIndexOfType(indexInThisColumn);
            if (indexInTargetColumn == null)
                return false;
        }

        return true;
    }

    public string GetIndexAsFormattedString()
    {
        return string.Join(", ", Indexes.Select(x => x.IndexName).ToList());
    }
    
    public Index? FindIndexOfType(Index index)
    {
        return FindIndexOfType(index.IndexName, index.Ordinal, index.IndexType);
    }

    public Index? FindIndexOfType(string indexName, int ordinal, string indexType)
    {
        return Indexes.FirstOrDefault(x => x.IndexName == indexName && x.Ordinal == ordinal && x.IndexType == indexType);
    }
    
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + (ColumnName != null ? ColumnName.GetHashCode() : 0);
            hash = hash * 23 + (DataType != null ? DataType.GetHashCode() : 0);
            hash = hash * 23 + (MaxLength != null ? MaxLength.GetHashCode() : 0);
            return hash;
        }
    }

    public override bool Equals(object obj)
    {
        if (obj == null || GetType() != obj.GetType())
        {
            return false;
        }

        Column other = (Column)obj;

        return (string.Equals(ColumnName, other.ColumnName) &&
                string.Equals(DataType, other.DataType) &&
                MaxLength == other.MaxLength);
    }
}