------------------------------------------------------------------
--view_table_fields.sql
--@Robin Baines 2008
--uses sys to get field in a table
------------------------------------------------------------------

DROP VIEW v_table_fields 
GO
CREATE VIEW v_table_fields AS
	SELECT 
	ta.name 'TableName',
	c.object_id 'TableId',
    c.name 'ColumnName',
    t.Name 'DataType',
    c.max_length 'MaxLength',
    c.precision ,
    c.scale ,
    c.is_nullable,
    ISNULL(i.is_primary_key, 0) 'Primary Key'
FROM
sys.tables ta
INNER JOIN    
    sys.columns c ON c.object_id = ta.object_id 
INNER JOIN 
    sys.types t ON c.system_type_id = t.system_type_id
LEFT OUTER JOIN 
    sys.index_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
LEFT OUTER JOIN 
    sys.indexes i ON ic.object_id = i.object_id AND ic.index_id = i.index_id
WHERE
    --c.object_id = OBJECT_ID('b_gpk') AND
	ISNULL(i.is_primary_key, 0) = 0	AND 
	t.name <> 'sysname'
GO
--Example of how to use.
--This creates the records needed for a trigger which logs changes.
	SELECT 'INSERT INTO b_gpk_log(gpk, field, remark)
    SELECT i.gpk, '''  
	+ ColumnName + ''', '' was '''''' + '	+ 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(d.' + ColumnName + ','''')' 
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + '''''' and is  '''''' + ' + 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(i.' + ColumnName + ','''')'
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + ''''''.'''
	+ ' FROM inserted i inner join deleted d on i.gpk = d.gpk where ISNULL(i.' 
	+ ColumnName + ', '''') <> ISNULL(d.' + ColumnName  + ', '''')'   
	FROM	v_table_fields f where f.tablename='b_gpk'
GO

	SELECT 'INSERT INTO b_hpk_log(hpk, field, remark)
    SELECT i.hpk, '''  
	+ ColumnName + ''', '' was '''''' + '	+ 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(d.' + ColumnName + ','''')' 
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + '''''' and is  '''''' + ' + 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(i.' + ColumnName + ','''')'
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + ''''''.'''
	+ ' FROM inserted i inner join deleted d on i.hpk = d.hpk where ISNULL(i.' 
	+ ColumnName + ', '''') <> ISNULL(d.' + ColumnName  + ', '''')'   
	FROM	v_table_fields f where f.tablename='b_hpk'
GO

	SELECT 'INSERT INTO b_zinr_log(zinr, field, remark)
    SELECT i.zinr, '''  
	+ ColumnName + ''', '' was '''''' + '	+ 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(d.' + ColumnName + ','''')' 
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(d.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + '''''' and is  '''''' + ' + 
	CASE WHEN f.[DataType] = 'nvarchar' THEN  ' ISNULL(i.' + ColumnName + ','''')'
	ELSE 
		CASE WHEN f.[DataType] = 'float' THEN  ' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 10, 4)) ' 
		ELSE 
		' LTRIM(STR( ISNULL(i.' + ColumnName + ', '''') , 2, 0)) '
		END
	END
	+ ' + ''''''.'''
	+ ' FROM inserted i inner join deleted d on i.zinr = d.zinr where ISNULL(i.' 
	+ ColumnName + ', '''') <> ISNULL(d.' + ColumnName  + ', '''')'   
	FROM	v_table_fields f where f.tablename='b_zinr'
GO

