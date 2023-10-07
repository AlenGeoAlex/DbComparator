namespace DbComparator;

public class ComparisonSettings
{
    public static ComparisonSettings Default { get; } = new ComparisonSettings();

    public static ComparisonSettings Create => new();
    public bool CreateReport { get; set; }
    public bool CheckIndex { get; set; }

    private ComparisonSettings()
    {
        CreateReport = true;
        CheckIndex = false;
    }
}