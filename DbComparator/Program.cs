using System;

namespace DbComparator;

static class Program
{
    private static global::DbComparator.DbComparator? _dbComparator;

    static void Main()
    {
        string? source = null;
        while (source == null)
        {
            Console.WriteLine("Enter the source database connection string");
            source = Console.ReadLine();
        }

        string? target = null;
        while (target == null)
        {
            Console.WriteLine("Enter the target database connection string");
            target = Console.ReadLine();
        }

        if (source == target)
        {
            Console.WriteLine("Exiting, Provided same connection string for source and target!");
            return;
        }

        var settings = ComparisonSettings.Create;

        Console.WriteLine("Do you want to compare the indexes too?");
        var checkIndexResponse = Console.ReadLine();
        if (checkIndexResponse != null &&
            (checkIndexResponse.ToUpper() == "Y" || checkIndexResponse.ToUpper() == "Yes"))
        {
            settings.CheckIndex = true;
            Console.WriteLine("Checking index is set to true..");
        }
        
        bool createExcel = false;
        Console.WriteLine("Create Excel file for the entire structural changes? Press Y/N");
        var excelConfirmation = Console.ReadLine();
        if (excelConfirmation != null && (excelConfirmation.ToUpper() == "Y" || excelConfirmation.ToUpper() == "YES"))
        {
            settings.CreateReport = true;
            Console.WriteLine("CSV file generation is set to true..");
        }

        _dbComparator = new global::DbComparator.DbComparator(source, target, settings);

        AppDomain.CurrentDomain.ProcessExit += CurrentDomainOnProcessExit;

        Console.WriteLine("Connect to Source VPN and Press Enter");
        Console.ReadLine();
        Console.WriteLine("Trying connection...");
        if (!_dbComparator.TryConnectSource())
        {
            Console.WriteLine("Failed to connect to source database");
            return;
        }
        Console.WriteLine("Connection Success, Starting to fetch structures...");
        _dbComparator.GetDataFromSource();
        _dbComparator.DisconnectSource();

        Console.WriteLine("Connect to Target VPN and Press Enter");
        Console.ReadLine();
        Console.WriteLine("Trying connection...");
        if (!_dbComparator.TryConnectTarget())
        {
            Console.WriteLine("Failed to connect to target database");
            return;
        }
        Console.WriteLine("Connection Success, Starting to fetch structures...");
        _dbComparator.GetDataFromTarget();
        _dbComparator.DisconnectTarget();
        
        _dbComparator.Analyze();

        if (settings.CreateReport)
        {
            _dbComparator.CreateReport();
        }
    }

    static void CurrentDomainOnProcessExit(object? sender, EventArgs e)
    {
        _dbComparator?.Disconnect();
        Console.WriteLine("Cleanup completed.");
    }
}