# PerformanceEye
A proc-based toolkit for solving both basic and advanced SQL Server performance and stability issues

Performance Eye ("for the database guy") is a set of tables, procedures, and SQL Agent jobs that polls the SQL Server catalog, DMVs, and other core diagnostic info to provide detailed information during troubleshooting sessions. A strong emphasis has been placed on performance; PerformanceEye code is likely to use < 2% CPU during a given minute, while still providing a wide set of metrics that will meet the needs of most DBAs for most performance and stability issues.
Additionally, a set of procedures provide an interface for displaying the most important data points from the collected info. Much effort has been made to ensure the output of these procedures returns the most relevant info in a digestible fashion. 

