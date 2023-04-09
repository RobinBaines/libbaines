EXEC p_dropview 'v_adhoc_views'
GO
---- =============================================
---- Author:  	RPB
---- Create date: 20141231
---- Description: Get list of views in the ADHOC schema.
---- Used By frmAdhocViews
---- =============================================
CREATE VIEW [dbo].[v_adhoc_views] 
AS
SELECT 
TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME
,(SELECT OBJECT_DEFINITION(object_id) FROM sys.views WHERE name = TABLE_NAME
AND [schema_id] = (SELECT [schema_id] FROM [sys].[schemas] WHERE NAME = TABLE_SCHEMA)) AS VIEW_DEFINITION
, CHECK_OPTION, IS_UPDATABLE
FROM information_schema.views 
WHERE TABLE_SCHEMA= 'ADHOC'
GO


EXEC p_dropview 'v_all_views'
GO
---- =============================================
---- Author:  	RPB
---- Create date: 20141231
---- Description: Get list of all views in all schemas.
---- Used By frmAdhocViews
---- =============================================
CREATE VIEW [dbo].[v_all_views] 
AS
SELECT 
TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME
,(SELECT OBJECT_DEFINITION(object_id) FROM sys.views WHERE name = TABLE_NAME
AND [schema_id] = (SELECT [schema_id] FROM [sys].[schemas] WHERE NAME = TABLE_SCHEMA)) AS VIEW_DEFINITION
, CHECK_OPTION, IS_UPDATABLE
FROM information_schema.views 
GO

EXEC p_dropview 'v_all_views_column'
GO
-- =============================================
-- Author:  	RPB
-- Create date: 20141231
-- Description: Get list of views and columns in all schema.
-- Used By 
-- =============================================
CREATE VIEW dbo.v_all_views_column
AS
select table_name, column_name, table_schema, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH 
FROM information_schema.columns 
GO


--DO THIS TO GENERATE THE VIEW CODE AND THEN DROP
EXEC p_dropview v_referenced_objects
GO
-- =============================================
-- Author:  	RPB
-- Create date: 
-- Description: Get list of objects (procs and views) refrenced
-- Used By frmAdhocViews
-- =============================================
CREATE VIEW dbo.v_referenced_objects
AS
	select referenced_server_name, referenced_schema_name, referenced_entity_name
	,referenced_id, referenced_class, referenced_class_desc,
	is_caller_dependent, is_ambiguous, is_selected, is_updated, is_select_all, is_all_columns_found
FROM sys.dm_sql_referenced_entities('dbo.p_order_substitute', 'OBJECT')
WHERE referenced_minor_id = 0
GO
--DROP VIEW dbo.v_referenced_objects

--DO THIS TO GENERATE THE VIEW CODE AND THEN DROP
EXEC p_dropview v_referencing_objects
GO
CREATE VIEW dbo.v_referencing_objects
AS
select referencing_schema_name, referencing_entity_name, referencing_id, referencing_class_desc, is_caller_dependent
FROM sys.dm_sql_referencing_entities ('v_substitute_lookup', 'OBJECT');
GO


--DO THIS TO GENERATE THE VIEW CODE AND THEN DROP
EXEC p_dropview v_INFORMATION_SCHEMA_ROUTINES
GO
CREATE VIEW dbo.v_INFORMATION_SCHEMA_ROUTINES
AS
select [SPECIFIC_CATALOG]
      ,[SPECIFIC_SCHEMA]
      ,[SPECIFIC_NAME]
      ,[ROUTINE_CATALOG]
      ,[ROUTINE_SCHEMA]
      ,[ROUTINE_NAME]
      ,[ROUTINE_TYPE]
      ,[MODULE_CATALOG]
      ,[MODULE_SCHEMA]
      ,[MODULE_NAME]
      ,[UDT_CATALOG]
      ,[UDT_SCHEMA]
      ,[UDT_NAME]
      ,[DATA_TYPE]
      ,[CHARACTER_MAXIMUM_LENGTH]
      ,[CHARACTER_OCTET_LENGTH]
      ,[COLLATION_CATALOG]
      ,[COLLATION_SCHEMA]
      ,[COLLATION_NAME]
      ,[CHARACTER_SET_CATALOG]
      ,[CHARACTER_SET_SCHEMA]
      ,[CHARACTER_SET_NAME]
      ,[NUMERIC_PRECISION]
      ,[NUMERIC_PRECISION_RADIX]
      ,[NUMERIC_SCALE]
      ,[DATETIME_PRECISION]
      ,[INTERVAL_TYPE]
      ,[INTERVAL_PRECISION]
      ,[TYPE_UDT_CATALOG]
      ,[TYPE_UDT_SCHEMA]
      ,[TYPE_UDT_NAME]
      ,[SCOPE_CATALOG]
      ,[SCOPE_SCHEMA]
      ,[SCOPE_NAME]
      ,[MAXIMUM_CARDINALITY]
      ,[DTD_IDENTIFIER]
      ,[ROUTINE_BODY]
      ,[ROUTINE_DEFINITION]
      ,[EXTERNAL_NAME]
      ,[EXTERNAL_LANGUAGE]
      ,[PARAMETER_STYLE]
      ,[IS_DETERMINISTIC]
      ,[SQL_DATA_ACCESS]
      ,[IS_NULL_CALL]
      ,[SQL_PATH]
      ,[SCHEMA_LEVEL_ROUTINE]
      ,[MAX_DYNAMIC_RESULT_SETS]
      ,[IS_USER_DEFINED_CAST]
      ,[IS_IMPLICITLY_INVOCABLE]
      ,[CREATED]
      ,[LAST_ALTERED]
  FROM [INFORMATION_SCHEMA].[ROUTINES]
GO

IF NOT EXISTS(SELECT 1 FROM sys.columns 
          WHERE Name = N'characters'
          AND Object_ID = Object_ID(N'[m_sql_characters]'))
BEGIN
CREATE TABLE m_sql_characters
(
characters sysname NOT NULL PRIMARY KEY
)
END
GO
DELETE FROM m_sql_characters
GO
INSERT INTO m_sql_characters VALUES(N',')
INSERT INTO m_sql_characters VALUES(N'(')
INSERT INTO m_sql_characters VALUES(N')')
INSERT INTO m_sql_characters VALUES(N'NULL')
INSERT INTO m_sql_characters VALUES(N'IS')
INSERT INTO m_sql_characters VALUES(N'*')
INSERT INTO m_sql_characters VALUES(N'+')
INSERT INTO m_sql_characters VALUES(N'-')
INSERT INTO m_sql_characters VALUES(N'=')
INSERT INTO m_sql_characters VALUES(N'>')
INSERT INTO m_sql_characters VALUES(N'<')
INSERT INTO m_sql_characters VALUES(N'!')
INSERT INTO m_sql_characters VALUES(N'AND')
INSERT INTO m_sql_characters VALUES(N'OR')
INSERT INTO m_sql_characters VALUES(N'EXISTS')
INSERT INTO m_sql_characters VALUES(N'NOT')
INSERT INTO m_sql_characters VALUES(N'%')
INSERT INTO m_sql_characters VALUES(N'<>')
--INSERT INTO m_sql_characters VALUES(N'0')
--INSERT INTO m_sql_characters VALUES(N'1')
--INSERT INTO m_sql_characters VALUES(N'2')
--INSERT INTO m_sql_characters VALUES(N'3')
--INSERT INTO m_sql_characters VALUES(N'4')
--INSERT INTO m_sql_characters VALUES(N'5')
--INSERT INTO m_sql_characters VALUES(N'6')
--INSERT INTO m_sql_characters VALUES(N'7')
--INSERT INTO m_sql_characters VALUES(N'8')
--INSERT INTO m_sql_characters VALUES(N'9')
GO

IF NOT EXISTS(SELECT 1 FROM sys.columns 
          WHERE Name = N'keyword'
          AND Object_ID = Object_ID(N'[m_keywords]'))
BEGIN
CREATE TABLE m_keywords
(
keyword sysname NOT NULL PRIMARY KEY
)
END
GO
DELETE FROM m_keywords
GO
SET NOCOUNT ON

INSERT INTO m_keywords VALUES(N'ADD')
INSERT INTO m_keywords VALUES(N'ALL')
INSERT INTO m_keywords VALUES(N'ALTER')
INSERT INTO m_keywords VALUES(N'AND')
INSERT INTO m_keywords VALUES(N'ANY')
INSERT INTO m_keywords VALUES(N'AS')
INSERT INTO m_keywords VALUES(N'ASC')
INSERT INTO m_keywords VALUES(N'AUTHORIZATION')
INSERT INTO m_keywords VALUES(N'BACKUP')
INSERT INTO m_keywords VALUES(N'BEGIN')
INSERT INTO m_keywords VALUES(N'BETWEEN')
INSERT INTO m_keywords VALUES(N'BREAK')
INSERT INTO m_keywords VALUES(N'BROWSE')
INSERT INTO m_keywords VALUES(N'BULK')
INSERT INTO m_keywords VALUES(N'BY')
INSERT INTO m_keywords VALUES(N'CASCADE')
INSERT INTO m_keywords VALUES(N'CASE')
INSERT INTO m_keywords VALUES(N'CHECK')
INSERT INTO m_keywords VALUES(N'CHECKPOINT')
INSERT INTO m_keywords VALUES(N'CLOSE')
INSERT INTO m_keywords VALUES(N'CLUSTERED')
INSERT INTO m_keywords VALUES(N'COALESCE') 
INSERT INTO m_keywords VALUES(N'COLLATE')
INSERT INTO m_keywords VALUES(N'COLUMN')
INSERT INTO m_keywords VALUES(N'COMMIT')
INSERT INTO m_keywords VALUES(N'COMPUTE')
INSERT INTO m_keywords VALUES(N'CONSTRAINT')
INSERT INTO m_keywords VALUES(N'CONTAINS')
INSERT INTO m_keywords VALUES(N'CONTAINSTABLE')
INSERT INTO m_keywords VALUES(N'CONTINUE')
INSERT INTO m_keywords VALUES(N'CONVERT')
INSERT INTO m_keywords VALUES(N'CREATE')
INSERT INTO m_keywords VALUES(N'CROSS')
INSERT INTO m_keywords VALUES(N'CURRENT')
INSERT INTO m_keywords VALUES(N'CURRENT_DATE')
INSERT INTO m_keywords VALUES(N'CURRENT_TIME')
INSERT INTO m_keywords VALUES(N'CURRENT_TIMESTAMP')
INSERT INTO m_keywords VALUES(N'CURRENT_USER')
INSERT INTO m_keywords VALUES(N'CURSOR')
INSERT INTO m_keywords VALUES(N'DATABASE')
INSERT INTO m_keywords VALUES(N'DBCC')
INSERT INTO m_keywords VALUES(N'DEALLOCATE')
INSERT INTO m_keywords VALUES(N'DECLARE')
INSERT INTO m_keywords VALUES(N'DEFAULT')
INSERT INTO m_keywords VALUES(N'DELETE')
INSERT INTO m_keywords VALUES(N'DENY')
INSERT INTO m_keywords VALUES(N'DESC')
INSERT INTO m_keywords VALUES(N'DISK')
INSERT INTO m_keywords VALUES(N'DISTINCT')
INSERT INTO m_keywords VALUES(N'DISTRIBUTED')
INSERT INTO m_keywords VALUES(N'DOUBLE')
INSERT INTO m_keywords VALUES(N'DROP')
INSERT INTO m_keywords VALUES(N'DUMMY')
INSERT INTO m_keywords VALUES(N'DUMP')
INSERT INTO m_keywords VALUES(N'ELSE')
INSERT INTO m_keywords VALUES(N'END')
INSERT INTO m_keywords VALUES(N'ERRLVL')
INSERT INTO m_keywords VALUES(N'ESCAPE')
INSERT INTO m_keywords VALUES(N'EXCEPT')
INSERT INTO m_keywords VALUES(N'EXEC')
INSERT INTO m_keywords VALUES(N'EXECUTE')
INSERT INTO m_keywords VALUES(N'EXISTS')
INSERT INTO m_keywords VALUES(N'EXIT')
INSERT INTO m_keywords VALUES(N'FETCH')
INSERT INTO m_keywords VALUES(N'FILE')
INSERT INTO m_keywords VALUES(N'FILLFACTOR')
INSERT INTO m_keywords VALUES(N'FOR')
INSERT INTO m_keywords VALUES(N'FOREIGN')
INSERT INTO m_keywords VALUES(N'FREETEXT')
INSERT INTO m_keywords VALUES(N'FREETEXTTABLE')
INSERT INTO m_keywords VALUES(N'FROM')
INSERT INTO m_keywords VALUES(N'FULL')
INSERT INTO m_keywords VALUES(N'FUNCTION')
INSERT INTO m_keywords VALUES(N'GOTO')
INSERT INTO m_keywords VALUES(N'GRANT')
INSERT INTO m_keywords VALUES(N'GROUP')
INSERT INTO m_keywords VALUES(N'HAVING')
INSERT INTO m_keywords VALUES(N'HOLDLOCK')
INSERT INTO m_keywords VALUES(N'IDENTITY')
INSERT INTO m_keywords VALUES(N'IDENTITY_INSERT')
INSERT INTO m_keywords VALUES(N'IDENTITYCOL')
INSERT INTO m_keywords VALUES(N'IF')
INSERT INTO m_keywords VALUES(N'IN')
INSERT INTO m_keywords VALUES(N'INDEX')
INSERT INTO m_keywords VALUES(N'INNER')
INSERT INTO m_keywords VALUES(N'INSERT')
INSERT INTO m_keywords VALUES(N'INTERSECT')
INSERT INTO m_keywords VALUES(N'INTO')
INSERT INTO m_keywords VALUES(N'IS')
INSERT INTO m_keywords VALUES(N'JOIN')
INSERT INTO m_keywords VALUES(N'KEY')
INSERT INTO m_keywords VALUES(N'KILL')
INSERT INTO m_keywords VALUES(N'LEFT')
INSERT INTO m_keywords VALUES(N'LIKE')
INSERT INTO m_keywords VALUES(N'LINENO')
INSERT INTO m_keywords VALUES(N'LOAD')
INSERT INTO m_keywords VALUES(N'NATIONAL')
INSERT INTO m_keywords VALUES(N'NOCHECK')
INSERT INTO m_keywords VALUES(N'NONCLUSTERED')
INSERT INTO m_keywords VALUES(N'NOT')
INSERT INTO m_keywords VALUES(N'NULL')
INSERT INTO m_keywords VALUES(N'NULLIF')
INSERT INTO m_keywords VALUES(N'OF')
INSERT INTO m_keywords VALUES(N'OFF')
INSERT INTO m_keywords VALUES(N'OFFSETS')
INSERT INTO m_keywords VALUES(N'ON')
INSERT INTO m_keywords VALUES(N'OPEN')
INSERT INTO m_keywords VALUES(N'OPENDATASOURCE')
INSERT INTO m_keywords VALUES(N'OPENQUERY')
INSERT INTO m_keywords VALUES(N'OPENROWSET')
INSERT INTO m_keywords VALUES(N'OPENXML')
INSERT INTO m_keywords VALUES(N'OPTION')
INSERT INTO m_keywords VALUES(N'OR')
INSERT INTO m_keywords VALUES(N'ORDER')
INSERT INTO m_keywords VALUES(N'OUTER')
INSERT INTO m_keywords VALUES(N'OVER')
INSERT INTO m_keywords VALUES(N'PERCENT')
INSERT INTO m_keywords VALUES(N'PLAN')
INSERT INTO m_keywords VALUES(N'PRECISION')
INSERT INTO m_keywords VALUES(N'PRIMARY')
INSERT INTO m_keywords VALUES(N'PRINT')
INSERT INTO m_keywords VALUES(N'PROC')
INSERT INTO m_keywords VALUES(N'PROCEDURE')
INSERT INTO m_keywords VALUES(N'PUBLIC')
INSERT INTO m_keywords VALUES(N'RAISERROR')
INSERT INTO m_keywords VALUES(N'READ')
INSERT INTO m_keywords VALUES(N'READTEXT')
INSERT INTO m_keywords VALUES(N'RECONFIGURE')
INSERT INTO m_keywords VALUES(N'REFERENCES')
INSERT INTO m_keywords VALUES(N'REPLICATION')
INSERT INTO m_keywords VALUES(N'RESTORE')
INSERT INTO m_keywords VALUES(N'RESTRICT')
INSERT INTO m_keywords VALUES(N'RETURN')
--added by RPB
INSERT INTO m_keywords VALUES(N'RETURNS')
INSERT INTO m_keywords VALUES(N'REVOKE')
INSERT INTO m_keywords VALUES(N'RIGHT')
INSERT INTO m_keywords VALUES(N'ROLLBACK')
INSERT INTO m_keywords VALUES(N'ROWCOUNT')
INSERT INTO m_keywords VALUES(N'ROWGUIDCOL')
INSERT INTO m_keywords VALUES(N'RULE')
INSERT INTO m_keywords VALUES(N'SAVE')
INSERT INTO m_keywords VALUES(N'SCHEMA')
INSERT INTO m_keywords VALUES(N'union select')
INSERT INTO m_keywords VALUES(N'SESSION_USER')
INSERT INTO m_keywords VALUES(N'SET')
INSERT INTO m_keywords VALUES(N'SETUSER')
INSERT INTO m_keywords VALUES(N'SHUTDOWN')
INSERT INTO m_keywords VALUES(N'SOME')
INSERT INTO m_keywords VALUES(N'STATISTICS')
INSERT INTO m_keywords VALUES(N'SYSTEM_USER')
INSERT INTO m_keywords VALUES(N'TABLE')
INSERT INTO m_keywords VALUES(N'TEXTSIZE')
INSERT INTO m_keywords VALUES(N'THEN')
INSERT INTO m_keywords VALUES(N'TO')
INSERT INTO m_keywords VALUES(N'TOP')
INSERT INTO m_keywords VALUES(N'TRAN')
INSERT INTO m_keywords VALUES(N'TRANSACTION')
INSERT INTO m_keywords VALUES(N'TRIGGER')
INSERT INTO m_keywords VALUES(N'TRUNCATE')
INSERT INTO m_keywords VALUES(N'TSEQUAL')
INSERT INTO m_keywords VALUES(N'UNION')
INSERT INTO m_keywords VALUES(N'UNIQUE')
INSERT INTO m_keywords VALUES(N'UPDATE')
INSERT INTO m_keywords VALUES(N'UPDATETEXT')
INSERT INTO m_keywords VALUES(N'USE')
INSERT INTO m_keywords VALUES(N'USER')
INSERT INTO m_keywords VALUES(N'VALUES')
INSERT INTO m_keywords VALUES(N'VARYING')
INSERT INTO m_keywords VALUES(N'VIEW')
INSERT INTO m_keywords VALUES(N'WAITFOR')
INSERT INTO m_keywords VALUES(N'WHEN')
INSERT INTO m_keywords VALUES(N'WHERE')
INSERT INTO m_keywords VALUES(N'WHILE')
INSERT INTO m_keywords VALUES(N'WITH')
INSERT INTO m_keywords VALUES(N'WRITETEXT')

--RPB
INSERT INTO m_keywords VALUES(N'NOFORMAT')
INSERT INTO m_keywords VALUES(N'NOINIT')
INSERT INTO m_keywords VALUES(N'SKIP')
INSERT INTO m_keywords VALUES(N'REWIND')
INSERT INTO m_keywords VALUES(N'NOUNLOAD')
INSERT INTO m_keywords VALUES(N'CATCH')
INSERT INTO m_keywords VALUES(N'TRY')
INSERT INTO m_keywords VALUES(N'OUT')

GO

--INSERT INTO [dbo].[m_form_tble_column__visibility]
--           ([form]
--           ,[tble]
--           ,[colmn]
--           ,[visible]
--           ,[prnt]
--           ,[sequence]
--           ,[bold]
--           ,[default_filter])

--union select [form]
--      ,[tble]
--      ,[colmn]
--      ,[visible]
--      ,[prnt]
--      ,[sequence]
--      ,[bold]
--      ,[default_filter]
--  FROM [dbo].[m_form_tble_column__visibility]
--GO





