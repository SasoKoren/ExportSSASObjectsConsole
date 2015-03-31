# ExportSSASObjectsConsole
##A windows console application that exports SSAS objects to a SQL Server table.

Syntax:
*ExportSSASObjectsConsole* *[SSAS Server]* *[SSAS Database Name]* *[SQL Server]* *[SQL Server Database Name]*

*[SSAS Server]* -> Instance name of the SSAS server (e.g. SERVER\INSTANCE)
*[SSAS Database Name]* -> Name of the SSAS database
*[SQL Server]* -> Target SQL Server instance 
*[SQL Server Database Name]* -> Target SQL Server database 

Application opens the SSAS database and reads the following structures of the SSAS database:

- Dimensions
-- Dimension attributes
--Dimension hierarchies and levels
- Cubes
-- Cube dimensions
-- Cube Measure groups
--- Cube measures + measure display folders
-- Calculated measures

The structure is then exported to the table *dbo.SSASDatabaseImport* in the target SQL database. In case the table does not exist, it is created.

---

##History

v1.0 (31.3.2015): Initial commit

