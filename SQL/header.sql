USE master;
GO
IF DB_ID (N'TestDb') IS NOT NULL
BEGIN
DECLARE @DatabaseName nvarchar(50)
SET @DatabaseName = N'TestDb'
DECLARE @SQL varchar(max)
SELECT @SQL = COALESCE(@SQL,'') + 'Kill ' + Convert(varchar, SPId) + ';'
FROM MASTER..SysProcesses
WHERE DBId = DB_ID(@DatabaseName) AND SPId <> @@SPId
EXEC(@SQL)
END
GO
IF DB_ID (N'TestDb') IS NOT NULL
DROP DATABASE TestDb;

GO
CREATE DATABASE TestDb;
GO
-- Verify the database files and sizes
SELECT name, size, size*1.0/128 AS [Size in MBs]
FROM sys.master_files
WHERE name = N'TestDb';
GO

USE TestDb;
GO
