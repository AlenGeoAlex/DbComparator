# DbComparator
This is a small utility application which can compare two different SQL Server databases and create an excel sheet based on that.
The created report will highlight the differences. Currently the differences it can identify are
 - Missing tables from first database to second database
 - Differences in existing table structures
   - This includes data type differences
   - Length differences etc
 - Index difference (Missing & Changes)
   - Both CLUSTERED and NONCLUSTERED INDEXES

#### NOTE : The design or architecture is just a simple console application. Nothing modular, nothing configurable. This was something which I designed for a quick need under 2 hours. It does its job pretty good as a small utility application to compare 2 databases rather than manually checking each out or buying something premium


## Working

1) Firstly it will ask you to enter the Source Server's Connection String
2) Secondly it will ask you to enter the Target Server's Connection String
3) After which, it will ask you to connect to source vpn (For some connections such as Azure SQL Server, private networks VPN's are required. If no VPN is there, you can just press enter)
4) After which, it will ask you to connect to target vpn (For some connections such as Azure SQL Server, private networks VPN's are required. If no VPN is there, you can just press enter)
5) It will ask whether report needs to be generated, index needs to checked etc.
6) Wait for sometime and it will generate reports


