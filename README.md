Database Setup and Excel Integration Documentation (Lab7)

Overview:
This assignment aims to guide you through setting up a MSSQL Docker container, deploying the Adventure2022.bak database, establishing a connection to Excel, and executing SQL queries to retrieve data and perform tasks like sorting Employees and Departments.

Tools Used:
Docker
MSSQL Server
Excel

Instructions:
Install Docker: If you haven't already, install Docker on your machine.
Run Docker Container:
    docker run -e "ACCEPT_EULA=Y" -e "MSSQL_SA_PASSWORD=your_pass" -e "MSSQL_PID=Evaluation" -p 1433:1433 --name CSA103 --hostname CSA103 -d mcr.microsoft.com/mssql/server:2022-preview-ubuntu-22.04

Deploying Adventure2022 Database:
Ensure you have the adventure2022.bak file ready.
Open MySQL server, connect to localhost .,1433
login - sa
password - your_password
Restore Your Database
Add AdventureBak2022 file to your server

Excel Integration:
Open Excel and navigate to the Data tab.
Get External Data:
Select "From Other Sources" -> "From SQL Server".

Enter server details:
Server: localhost,1433
Authentication: SQL Server Authentication
Username: sa
Password: your_password
Database: Adventure2022
Follow the prompts to import data from the desired tables. (Dont take first views, scroll down to second HumanResources)

Write SQL queries using functions like VLOOKUP to create tables with sorted Employees and find out their Departments.

You will certainly need those functions:
=IF(ISNUMBER(MATCH([@BusinessEntityID], Employee_TABLE_2[BusinessEntityID], 0)), "Employee", "Non-Employee")

=VLOOKUP(VLOOKUP([@BusinessEntityID],HumanResources_EmployeeDepartmentHistory,2),HumanResources_Department,2)

Additional Notes
1) If encountering issues connecting to the Docker container or executing queries, ensure Docker is running and the container is up.
2) Replace /path/to/adventure2022.bak with the actual path to your adventure2022.bak file.
3) For Excel queries, refer to Excel documentation for specific functions and query writing.