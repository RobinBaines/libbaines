--@RP Baines 200120624.
--Procedures to create SQL for preserving the format of tables and of tables in forms.
--This is useful for preserving the definitions before a database is dropped.
--The SQL output is run on the new version of the database to get the UI formatting back again.

--20120607 RPB ran these procs on gis_20120607 to copy the formats from the demo to a new version of the database.
--The resulting SQL was run on gis2. Result is that the definitions were copied from gis_20120607 to gis2.

--TODO Add to and then test delete statement [pCopyTxtDefinitions]

--use gis_20120607
--go
--exec dbo.[pCopyTableDefinitions] 'v_machine_medication_list'
--exec dbo.[pCopyTableDefinitions] 'v_order_bag_hpk'
--exec dbo.[pCopyTableDefinitions] 'v_order_expected'
--exec dbo.[pCopyTableDefinitions] 'v_order_reference'
--exec dbo.[pCopyTableDefinitions] 'v_reference_order_hpk_match_percent'
--exec dbo.[pCopyTableDefinitions] 'v_schema_dow_order_reference'
--exec dbo.[pCopyTableDefinitions] 'v_schema_machine_article_count_cassette_pivot'
--exec dbo.[pCopyTableDefinitions] 'v_schema_machine_capacity_bagcount_pivot'
--exec dbo.[pCopyTableDefinitions] 'v_schema_machine_order_pivot'
--exec dbo.[pCopyTableDefinitions] 'v_scoringpercentage'
--exec dbo.[pCopyTableDefinitions] 'v_gen'
--exec dbo.[pCopyTableDefinitions] 'v_reference_order_hpk_match_percent'
--exec dbo.pCopyFormDefinitions 'Medication Lists', 'v_machine_medication_list'
--exec dbo.pCopyFormDefinitions 'Planning',  'v_order_bag_hpk'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_order_expected'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_order_reference'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_reference_order_hpk_match_percent'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_schema_dow_order_reference'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_schema_machine_article_count_cassette_pivot'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_schema_machine_capacity_bagcount_pivot'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_scoringpercentage'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_gen'
--exec dbo.pCopyFormDefinitions  'Planning', 'b_schema'
--exec dbo.pCopyFormDefinitions  'Planning', 'b_dow'
--exec dbo.pCopyFormDefinitions  'Planning', 'v_reference_order_hpk_match_percent'

DROP PROCEDURE [dbo].[pCopyTableDefsFrom] 
go
CREATE PROCEDURE [dbo].[pCopyTableDefsFrom] 
	@ToTble nvarchar(100), 
	@FromTble nvarchar(100) 
	AS
	BEGIN
	DELETE FROM  m_tble_column WHERE tble = @ToTble
	
	INSERT INTO m_tble(tble)
	SELECT @ToTble where not exists(select 1 from m_tble where tble = @ToTble)
	
	INSERT INTO [m_tble_column]
		([tble]
		,[colmn]
		,[format]
		,[width])
	SELECT @ToTble
		,[colmn]
		,[format]
		,[width]
	FROM [m_tble_column]
	where tble = @FromTble
	end
GO
--example from Longshort2
--exec [dbo].[pCopyTableDefsFrom] 'rComponentStockSums', 'vComponentStockSums'

DROP PROCEDURE [dbo].[pCopyFormTableDefsFrom] 
go
CREATE PROCEDURE [dbo].[pCopyFormTableDefsFrom] 
	@ToForm nvarchar(100),
	@ToTble nvarchar(100), 
	@FromForm nvarchar(100),
	@FromTble nvarchar(100) 
	AS
	BEGIN
	DELETE FROM m_form_tble_column__visibility WHERE tble = @ToTble and form = @ToForm
	
	INSERT INTO [m_form_tble]([form],[tble])
    SELECT @ToForm, @ToTble where not exists(select 1 from [m_form_tble] where form = @ToForm and tble = @ToTble) 
	INSERT INTO [m_form_tble_column__visibility]
		([form]
		,[tble]
		,[colmn]
		,[visible]
		,[prnt]
		,[sequence]
		,[bold]
		,[default_filter])
	SELECT @ToForm
		,@ToTble
		,[colmn]
		,[visible]
		,[prnt]
		,[sequence]
		,[bold]
		,[default_filter]
	FROM [m_form_tble_column__visibility]
	where form = @FromForm and tble = @FromTble
	END
GO

DROP PROCEDURE [dbo].[pCopyFormDefsFrom] 
go
CREATE PROCEDURE [dbo].[pCopyFormDefsFrom] 
	@ToForm nvarchar(100),
	@FromForm nvarchar(100)

	AS
	BEGIN
	DELETE FROM m_form_tble_column__visibility WHERE form = @ToForm
	
	
	INSERT INTO [m_form_tble_column__visibility]
		([form]
		,[tble]
		,[colmn]
		,[visible]
		,[prnt]
		,[sequence]
		,[bold]
		,[default_filter])
	SELECT @ToForm
		,tble
		,[colmn]
		,[visible]
		,[prnt]
		,[sequence]
		,[bold]
		,[default_filter]
	FROM [m_form_tble_column__visibility]
	where form = @FromForm 
	END
GO
--example from Longshort2
--exec [dbo].[pCopyFormTableDefsFrom] 'MRPComp', 'rComponentStockSums', 'MRPComp', 'vComponentStockSums'

--Procs to make a script to copy from one database to another, for example from a test system to live.
DROP PROCEDURE [dbo].[pCopyTableDefinitions] 
go
CREATE PROCEDURE [dbo].[pCopyTableDefinitions] 
	@Tble nvarchar(100)
	AS
	BEGIN 
	SET NOCOUNT ON
	SET XACT_ABORT ON
	
	SELECT 'delete from m_tble_column_header where [tble] = ''' + 
	@Tble + ''''
	SELECT 'delete from m_tble_column where [tble] = ''' + 
	@Tble + ''''
	SELECT 'delete from m_tble where [tble] = ''' + 
	@Tble + ''''
	
	SELECT 'insert into m_tble ([tble]) values ( ''' + 
	[tble]
       		+ ''')'
	FROM m_tble
	where tble = @Tble
	
	SELECT 'insert into m_tble_column ([tble],[colmn],[format],[width]) values ( ''' + 
	[tble]
       + ''',''' + [colmn]
       + ''',''' + [format]
       + ''',' +  STR([width])
		+ ')'
	FROM m_tble_column
	where tble = @Tble
  
	SELECT 'insert into m_tble_column_header (tble, colmn, lang, header) values ( ''' +
		tble
       + ''',''' + colmn
       + ''',''' + lang
       + ''',''' + header
       + ''')'
    from m_tble_column_header
    where tble = @Tble
  
	END
	GO
	
DROP PROCEDURE [dbo].[pCopyFormDefinitions] 
go
CREATE PROCEDURE [dbo].[pCopyFormDefinitions] 
	@Form nvarchar(100),
	@Tble nvarchar(100)
	AS
	BEGIN 
	SET NOCOUNT ON
	SET XACT_ABORT ON
	
	
	SELECT 'delete from m_form_tble_column__visibility where form = ''' +
	@Form + ''' and [tble] =''' + @Tble + ''''
	SELECT 'delete from m_form_tble where form = ''' +
	@Form + ''' and [tble] =''' + @Tble + ''''
	
	SELECT 'insert into m_form ([form]) SELECT ''' + 
		[form]
		+ ''' WHERE NOT EXISTS (SELECT 1 FROM m_form WHERE  form = ''' + @Form + ''')'
	FROM m_form
	where form = @Form

	SELECT 'insert into m_form_tble ([form], tble) values ( ''' + 
		[form]
		+ ''',''' + [tble]
		+ ''')'
	FROM m_form_tble
	where form = @Form and tble = @Tble

	SELECT 'insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( ''' + [form]
		+ ''',''' + [tble]
		+ ''',''' + [colmn]
		+ ''', ' +  STR(visible)
		+ ', ' +  STR(prnt)
		+ ', ' +  STR(sequence)
		+', ' +  STR(bold)
		+ ',''' + default_filter
		+ ''')'
	FROM m_form_tble_column__visibility
	where form = @Form and tble = @Tble

	END
GO
DROP PROCEDURE [dbo].[pCopyTxtDefinitions] 
go
CREATE PROCEDURE [dbo].[pCopyTxtDefinitions] 
	AS
	BEGIN 
    SELECT 'insert into m_txt (txt,[typ],[descr]) values ( ''' + 
		[txt]
       + ''',''' + [typ]
       + ''',''' + [descr]
       + ''')'
    FROM [m_txt]

	SELECT 'insert into m_txt_header (txt,lang,header) values ( ''' + 
		[txt]
       + ''',''' + lang
       + ''',''' + header
       + ''')'
    FROM m_txt_header
	END
GO	

