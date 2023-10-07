namespace DbComparator.Entity;

public class Index
{
    public readonly string IndexName;
    public readonly string IndexType;
    public readonly int Ordinal;
    public bool IsInclude => Ordinal == 0;

    public Index(string indexName, string indexType, int ordinal)
    {
        IndexName = indexName;
        IndexType = indexType;
        Ordinal = ordinal;
    }

    public string GetIndexAsFormatted()
    {
        return $"{IndexName} - {IndexType}";
    }
    
}