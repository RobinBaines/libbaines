-----------------------------------------------------------------------------
--1. 2 meta procs 
-----------------------------------------------------------------------------
IF EXISTS ( SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID(N'p_dropFK')
                    AND type IN ( N'P', N'PC' ) ) 
DROP PROC p_dropFK
GO
-- =============================================
-- Author:  RPB
-- Create date: 20160320
-- Description: Proc to drop an FK. 
-- Used By 
-- =============================================
CREATE PROC p_dropFK
@table_name VARCHAR(1048),
@column_name VARCHAR(1048),
@ref_table_name VARCHAR(1048)
AS
DECLARE @fkname AS VARCHAR(3144)
SET @fkname = ISNULL((select fks.name --t.name as TableWithForeignKey, fk.constraint_column_id as FK_PartNo , c.name as ForeignKeyColumn 
					from sys.foreign_key_columns as fk
					inner join sys.tables as t on fk.parent_object_id = t.object_id
					inner join sys.columns as c on fk.parent_object_id = c.object_id and fk.parent_column_id = c.column_id
					JOIN sys.foreign_keys fks ON fks.OBJECT_ID = fk.constraint_object_id
					where fk.referenced_object_id = (select object_id from sys.tables where name = @ref_table_name)
					AND c.name = @column_name AND t.name = @table_name
					), '')
IF len(@fkname) > 0
		BEGIN
		DECLARE @ex AS VARCHAR(3144)
			SET @ex = 'ALTER TABLE ' + @table_name + ' DROP CONSTRAINT ' + @fkname 
			EXEC(@ex)
		END
GO


IF EXISTS ( SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID(N'p_createFK')
                    AND type IN ( N'P', N'PC' ) ) 
DROP PROC p_createFK
GO
-- =============================================
-- Author:  RPB
-- Create date: 20160320
-- Description: Proc to make an FK. 
-- Used By 
-- =============================================
CREATE PROC p_createFK
@table_name VARCHAR(1048),
@column_name VARCHAR(1048),
@ref_table_name VARCHAR(1048),
@cascade VARCHAR(1048)
AS
DECLARE @exists AS BIT
SET @exists = ISNULL((select 1 --t.name as TableWithForeignKey, fk.constraint_column_id as FK_PartNo , c.name as ForeignKeyColumn 
					from sys.foreign_key_columns as fk
					inner join sys.tables as t on fk.parent_object_id = t.object_id
					inner join sys.columns as c on fk.parent_object_id = c.object_id and fk.parent_column_id = c.column_id
					where fk.referenced_object_id = (select object_id from sys.tables where name = @ref_table_name)
					AND c.name = @column_name AND t.name = @table_name
					), 0)
IF @exists = 0
		BEGIN
			DECLARE @ex AS VARCHAR(3144)
			SET @ex = 'ALTER TABLE ' + @table_name + ' WITH CHECK ADD FOREIGN KEY(' + @column_name + ')' +
				' REFERENCES ' + @ref_table_name + '(' + @column_name + ') ' + @cascade --ON UPDATE CASCADE'
			EXEC(@ex)
		END
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_createcolumn]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].f_createcolumn
GO
-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Function to make a string to create a column if the column does not already exist. 
-- A function cannot run EXEC so this is done in p_createcolumn.
-- Used By p_createcolumn
-- =============================================
CREATE FUNCTION dbo.f_createcolumn
(
@table_name VARCHAR(1048),
@column_name VARCHAR(1048),
@type_etc VARCHAR(1048)
)
RETURNS VARCHAR(3144) AS 
	BEGIN
		IF NOT EXISTS(SELECT * FROM sys.columns
		WHERE Name = @column_name AND OBJECT_ID = OBJECT_ID(@table_name))
		BEGIN
			DECLARE @ex AS VARCHAR(3144)
			SET @ex = 'ALTER TABLE ' + @table_name + ' ADD ' + @column_name + ' ' + @type_etc
			RETURN @ex
		END
	RETURN '';
END
GO

IF EXISTS ( SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID(N'p_createcolumn')
                    AND type IN ( N'P', N'PC' ) ) 
DROP PROC p_createcolumn
GO
-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Proc to make to create a column if the column does not already exist. 
-- Used By p_createcolumn
-- =============================================
CREATE PROC p_createcolumn
@table_name VARCHAR(1048),
@column_name VARCHAR(1048),
@type_etc VARCHAR(1048)
AS
DECLARE @ex AS VARCHAR(3144)
SET @ex =  dbo.f_createcolumn(@table_name, @column_name, @type_etc)
IF len(@ex) > 0 
	EXEC(@ex)
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_dropview]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].f_dropview
GO

-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Function to make a string to drop a view. 
-- works only in schema dbo.
-- Used By p_dropview
-- =============================================
CREATE FUNCTION dbo.f_dropview
(
@view_name VARCHAR(1048)
)
RETURNS VARCHAR(1048) AS 
	BEGIN
		IF EXISTS(select * FROM sys.views where object_id = OBJECT_ID('dbo.' + @view_name))
		BEGIN
			DECLARE @ex AS VARCHAR(3144)
			SET @ex = 'DROP VIEW ' + @view_name 
			RETURN @ex
		END
	RETURN '';
END
GO

IF EXISTS ( SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID(N'p_dropview')
                    AND type IN ( N'P', N'PC' ) ) 
DROP PROC p_dropview
GO

-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Proc to make to drop a view. 
-- Used manually
-- =============================================
CREATE PROC p_dropview
@view_name VARCHAR(1048)
AS
DECLARE @ex AS VARCHAR(3144)
SET @ex =  dbo.f_dropview(@view_name)
IF len(@ex) > 0 
	EXEC(@ex)
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_dropproc]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].f_dropproc
GO
-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Function to make a string to drop a proc. 
-- Used By p_dropproc
-- =============================================
CREATE FUNCTION dbo.f_dropproc
(
@proc_name VARCHAR(1048)
)
RETURNS VARCHAR(1048) AS 
	BEGIN
		IF EXISTS(SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID('dbo.' + @proc_name)
                    AND type IN ( N'P', N'PC' ) )
		BEGIN
			DECLARE @ex AS VARCHAR(3144)
			SET @ex = 'DROP PROC ' + @proc_name 
			RETURN @ex
		END
	RETURN '';
END
GO
IF EXISTS ( SELECT  * FROM   sys.objects WHERE   object_id = OBJECT_ID(N'p_dropproc')
                    AND type IN ( N'P', N'PC' ) ) 
DROP PROC p_dropproc
GO

-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Proc to make to drop a proc. 
-- Used manually
-- =============================================
CREATE PROC p_dropproc
@proc_name VARCHAR(1048)
AS
DECLARE @ex AS VARCHAR(3144)
SET @ex =  dbo.f_dropproc(@proc_name)
IF len(@ex) > 0 
	EXEC(@ex)
GO

EXEC p_dropproc 'p_dropDefaultConstraint'
GO
CREATE PROC [dbo].[p_dropDefaultConstraint]
@table_name VARCHAR(1048),
@column_name VARCHAR(1048)
AS
	DECLARE @table_id AS INT
	DECLARE @name_column_id AS INT
	DECLARE @sql nvarchar(255) 

	-- Find table id
	SET @table_id = OBJECT_ID(@table_name)

	-- Find name column id
	SELECT @name_column_id = column_id
	FROM sys.columns
	WHERE object_id = @table_id
	AND name = @column_name

	-- Remove default constraint from name column
	SELECT @sql = 'ALTER TABLE ' + @table_name + '  DROP CONSTRAINT ' + D.name
	FROM sys.default_constraints AS D
	WHERE D.parent_object_id = @table_id
	AND D.parent_column_id = @name_column_id
	EXECUTE sp_executesql @sql
GO