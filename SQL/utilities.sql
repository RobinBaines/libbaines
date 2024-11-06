------------------------------------------------------------------
--Utilities.sql
--@Robin Baines 2008
--
--20120716 Remove old bUsers and bLevels tables.
--Remove TYPE definitions TYP_M_STRING and TYP_M_LANG.
--20241031 added [m_form_helptext]
--See readme.sql for instructions.
------------------------------------------------------------------
go
--Retain these definitions and do not change length of fields for compatability.
--CREATE TYPE TYP_M_STRING	FROM NVARCHAR (100)
--CREATE TYPE TYP_M_LANG		FROM VARCHAR(2)
go
--drop view v_txt_header
go

-------------------------------------------------
--utilities
-------------------------------------------------
IF NOT EXISTS(SELECT 1 FROM sys.columns 
          WHERE Name = N'helptext'
          AND Object_ID = Object_ID(N'[m_form_helptext]'))
BEGIN

CREATE TABLE [m_form_helptext](
	[form] [nvarchar](100) NOT NULL,
	[helptext] [nvarchar](max) NULL,
	PRIMARY KEY CLUSTERED ([form] ASC),
	FOREIGN KEY([form]) REFERENCES [dbo].[m_form] ([form]) ON UPDATE CASCADE ON DELETE CASCADE
)
END
GO


IF EXISTS (SELECT * FROM sys.tables where object_id = OBJECT_ID('dbo.m_sql_characters'))
BEGIN
	DROP TABLE [dbo].[m_sql_characters]
END
GO
CREATE TABLE [dbo].[m_sql_characters](
	[characters] [sysname] NOT NULL,
PRIMARY KEY([characters] ASC)
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[m_keywords](
	[keyword] [sysname] NOT NULL,
PRIMARY KEY([keyword] ASC)
) ON [PRIMARY]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr') 
AND type in (N'V'))
drop view v_usr
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr_grp') 
AND type in (N'V'))
drop view v_usr_grp
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_tble_column_header') 
AND type in (N'V'))
drop view v_tble_column_header
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_txt_header') 
AND type in (N'V'))
drop view v_txt_header
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr_grp_groupbox') 
AND type in (N'V'))
drop view v_usr_grp_groupbox
go
EXEC p_dropproc p_insert_m_form_grp_temp
GO
EXEC p_dropproc p_update_m_form_grp
GO
--drop view v_tble_column_header
-------------------------------------------------------
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_usr_change_log') 
AND type in (N'U'))
DROP TABLE m_usr_change_log
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_grp_log') 
AND type in (N'U'))
DROP TABLE [m_form_grp_log]
GO
--table to hold RO setting for a groupbox in a form.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_grp_groupbox') 
AND type in (N'U'))
DROP TABLE m_form_grp_groupbox
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_grp_temp') AND type in (N'U'))
DROP TABLE m_form_grp_temp 
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_version') AND type in (N'U'))
DROP table m_version
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_tble_column__visibility') AND type in (N'U'))
DROP table m_form_tble_column__visibility
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_tble') AND type in (N'U'))
DROP table m_form_tble
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_tble_column_header') AND type in (N'U'))
DROP table m_tble_column_header
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_tble_column') AND type in (N'U'))
DROP table m_tble_column
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_tble_header') AND type in (N'U'))
DROP table m_tble_header
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_tble') AND type in (N'U'))
DROP table m_tble
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_usr') 
AND type in (N'U'))
DROP table m_usr
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_grp') 
AND type in (N'U'))
DROP table m_form_grp
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_grp') 
AND type in (N'U'))
DROP table m_grp
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form_header') 
AND type in (N'U'))
DROP table m_form_header
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_form') 
AND type in (N'U'))
DROP table m_form
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_txt_header') 
AND type in (N'U'))
DROP table m_txt_header
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_txt') 
AND type in (N'U'))
DROP table m_txt
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_lang') 
AND type in (N'U'))
DROP table m_lang
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_format') 
AND type in (N'U'))
DROP table m_format
go






CREATE TABLE m_format	
			(
			format NVARCHAR (100),
			descr nvarchar(max), 
			PRIMARY KEY (format)
			)
	go		
go
CREATE TABLE m_version 	(
		id int not null default 1,
		version int not null,
		multi_lang bit not null default 0,
		PRIMARY KEY (version),
		check(id = 1)
		)
GO
Create TABLE m_lang
	(
		lang VARCHAR(2) NOT NULL,
		CONSTRAINT PK_m_lang PRIMARY KEY (lang)
	)
go
--insert into m_version(version, multi_lang) values (1, 1)
--the forms
CREATE TABLE m_form
	(
		form NVARCHAR (100), 
		menu bit not null default 0,
		menu_entry bit not null default 0,
		CONSTRAINT PK_m_form PRIMARY KEY (form)
	)
go
--grp of usr.
CREATE TABLE m_grp 	
	(
		grp NVARCHAR (100) Primary Key
		,createtime DATETIME DEFAULT GETDATE()
	)
GO

--20141230
CREATE TABLE m_usr_change_log 	(
		Id INT IDENTITY (1, 1) NOT NULL,
		usr NVARCHAR (100),
		grp NVARCHAR (100),
		lang VARCHAR(2) NOT NULL,
		blocking bit not null,
		name NVARCHAR (100),
		email NVARCHAR (100),
		telephone NVARCHAR (100), 
		createtime DATETIME DEFAULT GETDATE(),
		usr_change NVARCHAR(100) NULL DEFAULT (SUSER_NAME()),
		remark NVARCHAR(50),
		PRIMARY KEY(id)
	)
GO
	ALTER TABLE m_usr_change_log ADD CONSTRAINT FK_m_usr_change_log_grp
	FOREIGN KEY (grp) REFERENCES m_grp (grp)  ON DELETE CASCADE ON UPDATE CASCADE
GO

--20141230
CREATE TABLE [dbo].[m_form_grp_log](
	Id INT IDENTITY (1, 1) NOT NULL,
	grp [nvarchar](100) NOT NULL,
	form [nvarchar](100) NOT NULL,
	RO [bit] NOT NULL,
	createtime DATETIME DEFAULT GETDATE(),
	usr NVARCHAR(100) NULL DEFAULT (SUSER_NAME()),
	remark [nvarchar](50),
	PRIMARY KEY(id)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[m_form_grp_log]  WITH CHECK ADD  CONSTRAINT [FK_m_form_grp_log_form] FOREIGN KEY([form])
REFERENCES [dbo].[m_form] ([form])
ON DELETE CASCADE ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[m_form_grp_log] CHECK CONSTRAINT [FK_m_form_grp_log_form]
GO
ALTER TABLE [dbo].[m_form_grp_log]  WITH CHECK ADD  CONSTRAINT [FK_m_form_grp_log_grp] FOREIGN KEY([grp])
REFERENCES [dbo].[m_grp] ([grp])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[m_form_grp_log] CHECK CONSTRAINT [FK_m_form_grp_log_grp]
GO

---------------------------------------------
--20100817 Added blocking. set to 1 to activate blocking.
CREATE TABLE m_usr 	(
		usr NVARCHAR (100) PRIMARY KEY,
		grp NVARCHAR (100) default 'Commercial',
		lang VARCHAR(2) NOT NULL,
		blocking bit not null default 1,
		name NVARCHAR (100),
		email NVARCHAR (100),
		telephone NVARCHAR (100),
		--20141230
		createtime DATETIME DEFAULT GETDATE() 
	)
GO
	ALTER TABLE m_usr ADD CONSTRAINT FK_m_usr_grp
	FOREIGN KEY (grp) REFERENCES m_grp (grp) ON UPDATE CASCADE
GO
	ALTER TABLE m_usr ADD CONSTRAINT FK_m_usr_lang
	FOREIGN KEY (lang) REFERENCES m_lang (lang) --ON DELETE CASCADE ON UPDATE CASCADE
go

--20141230
CREATE TRIGGER m_usr_Trig_Log_insert ON m_usr AFTER INSERT AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO m_usr_change_log(usr, grp, lang, blocking, name, email, telephone, remark)
SELECT usr, grp, lang, blocking, name, email, telephone, 'INSERT' FROM inserted i 
GO

--20141230
CREATE TRIGGER m_usr_Trig_Log_update ON m_usr AFTER UPDATE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO m_usr_change_log(usr, grp, lang, blocking, name, email, telephone, remark)
SELECT i.usr, i.grp, i.lang, i.blocking, i.name, i.email, i.telephone, 'UPDATE' FROM inserted i 
GO

--20141230
CREATE TRIGGER m_usr_Trig_Log_delete ON m_usr AFTER DELETE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO m_usr_change_log(usr, grp, lang, blocking, name, email, telephone, remark)
SELECT i.usr, i.grp, i.lang, i.blocking, i.name, i.email, i.telephone, 'DELETE' FROM deleted i 
GO
---------------------------------------------
CREATE TABLE m_form_grp 	(
		grp NVARCHAR (100),
		form NVARCHAR (100),
		RO bit not null default 1,
		--20141230
		createtime DATETIME DEFAULT GETDATE(),
		PRIMARY KEY (grp, form),
		)
GO
	ALTER TABLE m_form_grp ADD CONSTRAINT FK_m_form_grp_grp
	FOREIGN KEY (grp) REFERENCES m_grp (grp) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_form_grp ADD CONSTRAINT FK_m_form_grp_form
	FOREIGN KEY (form) REFERENCES m_form(form) ON UPDATE CASCADE
GO

--20141230
CREATE TRIGGER m_form_grp_Trig_Log_insert ON m_form_grp AFTER INSERT AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO [m_form_grp_log](grp, form, RO, remark)
SELECT grp, form, RO, 'INSERT' FROM inserted i 
GO

--20141230
CREATE TRIGGER m_form_grp_Trig_Log_update ON m_form_grp AFTER UPDATE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO [m_form_grp_log](grp, form, RO, remark)
SELECT i.grp, i.form, i.RO, 'UPDATE' FROM inserted i 
--INNER JOIN deleted d 
GO

--20141230
CREATE TRIGGER m_form_grp_Trig_Log_delete ON m_form_grp AFTER DELETE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
INSERT INTO [m_form_grp_log](grp, form, RO, remark)
SELECT i.grp, i.form, i.RO, 'DELETE' FROM deleted i 
GO


------------------------------------------------------------
CREATE TABLE m_tble
	(
		tble NVARCHAR (100), 
		CONSTRAINT PK_m_tble PRIMARY KEY (tble)
	)
GO
-------------------------------------
--the headers of the forms
-------------------------------------
CREATE TABLE m_form_header
	(
		form NVARCHAR (100), 
		lang VARCHAR(2) NOT NULL,
		header NVARCHAR (100), 
		CONSTRAINT PK_m_form_header PRIMARY KEY (form, lang)
	)
go
ALTER TABLE m_form_header ADD CONSTRAINT FK_m_form_header_form
	FOREIGN KEY (form) REFERENCES m_form (form) ON DELETE CASCADE ON UPDATE CASCADE
go
ALTER TABLE m_form_header ADD CONSTRAINT FK_m_form_header_lang
	FOREIGN KEY (lang) REFERENCES m_lang (lang) --ON DELETE CASCADE ON UPDATE CASCADE
go
--important that a header is not used for the same form because form is lookup in the application using the header.
CREATE UNIQUE NONCLUSTERED INDEX m_form_header_lang_header ON m_form_header
(
		lang ASC, 
		header ASC
)WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]
go

-------------------------------------
--the headers of the forms
-------------------------------------
CREATE TABLE m_tble_header
	(
		tble NVARCHAR (100), 
		lang VARCHAR(2) NOT NULL,
		header NVARCHAR (100), 
		CONSTRAINT PK_m_tble_header PRIMARY KEY (tble, lang)
	)
go
	ALTER TABLE m_tble_header ADD CONSTRAINT FK_m_tble_header_tble
	FOREIGN KEY (tble) REFERENCES m_tble (tble) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_tble_header ADD CONSTRAINT FK_m_tble_header_lang
	FOREIGN KEY (lang) REFERENCES m_lang (lang) --ON DELETE CASCADE ON UPDATE CASCADE
go

--the grid column headers
--DROP TABLE m_tble_column
--20110309 RPB. format was not null. Changed to null but left on some databases (ABC, CJPT etc).
--The update to this table in Utilities uses isnull(format, '') in the dataset update statement.
CREATE TABLE m_tble_column
	(
		tble NVARCHAR (100), 
		colmn NVARCHAR (100), 
		format NVARCHAR (100) null, 
		width int not null default 100,
		CONSTRAINT PK_m_tble_column PRIMARY KEY (tble, colmn)
	)
go
	ALTER TABLE m_tble_column ADD CONSTRAINT FK_m_tble_column_tble
	FOREIGN KEY (tble) REFERENCES m_tble (tble) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_tble_column ADD CONSTRAINT FK_m_tble_column_format
	FOREIGN KEY (format) REFERENCES m_format (format) 
go
--the headers of the columns
CREATE TABLE m_tble_column_header
	(
		tble NVARCHAR (100), 
		colmn NVARCHAR (100), 
		lang VARCHAR(2) NOT NULL,
		header NVARCHAR (100), 
		CONSTRAINT PK_m_tble_column_header PRIMARY KEY (tble, colmn, lang)
	)
go
	ALTER TABLE m_tble_column_header ADD CONSTRAINT FK_m_tble_column_header_tble_colmn
	FOREIGN KEY (tble, colmn) REFERENCES m_tble_column (tble, colmn) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_tble_column_header ADD CONSTRAINT FK_m_tble_column_header_lang
	FOREIGN KEY (lang) REFERENCES m_lang (lang) --ON DELETE CASCADE ON UPDATE CASCADE
go

--the tables/grids used in the forms.
	CREATE TABLE m_form_tble
	(
		form NVARCHAR (100), 
		tble NVARCHAR (100), 
		CONSTRAINT PK_m_form_tble PRIMARY KEY (form, tble)
	)
go
	ALTER TABLE m_form_tble ADD CONSTRAINT FK_m_form_tble_tble
	FOREIGN KEY (tble) REFERENCES m_tble (tble) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_form_tble ADD CONSTRAINT FK_m_form_tble_form
	FOREIGN KEY (form) REFERENCES m_form (form) ON DELETE CASCADE ON UPDATE CASCADE
go
--the visibility in the grid in a form.
--20100330 RPB added default_filter to m_form_tble_column__visibility
CREATE TABLE m_form_tble_column__visibility
	(
		form NVARCHAR (100), 
		tble NVARCHAR (100), 
		colmn NVARCHAR (100), 
		visible bit not null default(1),
		prnt  bit not null default(1),
		sequence int,
		bold bit not null default(0),
		default_filter NVARCHAR (100) null default '', 
		CONSTRAINT PK_m_form_tble_column PRIMARY KEY (form, tble, colmn)
	)
go
	ALTER TABLE m_form_tble_column__visibility ADD CONSTRAINT FK_m_form_tble_column__visibility_form
	FOREIGN KEY (form) REFERENCES m_form (form) ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_form_tble_column__visibility ADD CONSTRAINT FK_m_form_tble_column__visibility_tble_colmn
	FOREIGN KEY (tble, colmn) REFERENCES m_tble_column (tble, colmn) ON DELETE CASCADE ON UPDATE CASCADE
go

CREATE TABLE m_form_grp_groupbox 	(
		grp NVARCHAR (100),
		form NVARCHAR (100),
		groupbox NVARCHAR (100),
		RO bit not null default 1,
		PRIMARY KEY (grp, form, groupbox),
		)
GO
	ALTER TABLE m_form_grp_groupbox ADD CONSTRAINT FK_m_form_grp_groupbox_grp_form
	FOREIGN KEY (grp, form) REFERENCES m_form_grp (grp, form) ON DELETE CASCADE ON UPDATE CASCADE
go

---------------------------------------------
--Lookup tables for any text.
--20201207 Added createtime DATETIME NULL DEFAULT (GETDATE()),
---------------------------------------------
	CREATE TABLE m_txt
	(
		txt NVARCHAR (100), 
		typ  NVARCHAR (100) null, --just for free sorting
		descr nvarchar(max),
		createtime DATETIME NULL DEFAULT (GETDATE()),
		CONSTRAINT PK_m_txt PRIMARY KEY (txt)
	)
go
--	DROP TABLE m_txt_header
	CREATE TABLE m_txt_header
	(
		txt NVARCHAR (100), 
		lang VARCHAR(2) NOT NULL,
		header nvarchar(250), 
		CONSTRAINT PK_m_txt_header PRIMARY KEY (txt, lang)
	)
go
	ALTER TABLE m_txt_header ADD CONSTRAINT FK_m_txt_header_txt
	FOREIGN KEY (txt) REFERENCES m_txt (txt)  ON DELETE CASCADE ON UPDATE CASCADE
go
	ALTER TABLE m_txt_header ADD CONSTRAINT FK_m_txt_lang
	FOREIGN KEY (lang) REFERENCES m_lang (lang) --ON DELETE CASCADE ON UPDATE CASCADE
go
--20141230
--a temporary table so that updating m_form_grp fills the log in the expected way.
CREATE TABLE m_form_grp_temp 	(
		grp NVARCHAR (100),
		form NVARCHAR (100),
		RO bit not null default 1,
		PRIMARY KEY (grp, form),
		)
GO
------------------------------------------------------------
--VIEWS
-------------------------------------------------

--Used in v_usr_log
CREATE view [v_usr] as
	select u.usr
		,u.grp
		,u.lang
		,u.blocking
		,u.name
		,u.email
		,u.telephone
		,cast(1 as bit) as Excel2003
	from m_usr u 
GO

create view v_usr_grp as
	select u.usr
		,u.grp
		,u.lang
		,f.form
		,f.RO
		,1 as Excel2003
	from m_usr u inner join m_form_grp f on u.grp = f.grp
		inner join m_version on multi_lang = 1
go

create view v_tble_column_header  WITH VIEW_METADATA, SCHEMABINDING AS
	select c.tble
		, c.colmn
		, l.lang
		, isnull(h.header, c.colmn) as header 
	from dbo.m_tble_column c cross join dbo.m_lang l
		left join dbo.m_tble_column_header h on c.tble=h.tble and c.colmn = h.colmn and l.lang = h.lang
go
CREATE TRIGGER [dbo].[v_tble_column_header_Trg_MergeInsert] ON [dbo].[v_tble_column_header] INSTEAD OF INSERT, UPDATE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON

	delete from m_tble_column_header
	where exists (SELECT 1 FROM inserted I
	where i.tble = m_tble_column_header.tble
		and i.colmn = m_tble_column_header.colmn
		and i.lang = m_tble_column_header.lang) 

	insert into m_tble_column_header
           ([tble]
           ,[colmn]
           ,[lang]
           ,[header])
     select 
           i.tble
           ,i.colmn
           ,i.lang
           ,i.header
           from inserted i 
GO
CREATE TRIGGER [dbo].[v_tble_column_header_Trg_delete] ON [dbo].[v_tble_column_header] INSTEAD OF DELETE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON

	delete from m_tble_column_header
	where exists (SELECT 1 FROM deleted d
	where d.tble = m_tble_column_header.tble
		and d.colmn = m_tble_column_header.colmn
		and d.lang = m_tble_column_header.lang) 
go

--drop view v_txt_header
create view v_txt_header  WITH VIEW_METADATA, SCHEMABINDING AS

	select c.txt
	--, c.typ, c.descr
	, l.lang, isnull(h.header, c.txt) as header from dbo.m_txt c cross join dbo.m_lang l
		left join dbo.m_txt_header h on c.txt=h.txt and l.lang = h.lang
go
CREATE TRIGGER [dbo].[v_txt_header_Trg_MergeInsert] ON [dbo].[v_txt_header] INSTEAD OF INSERT, UPDATE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON

	delete from m_txt_header
	where exists (SELECT 1 FROM inserted I
		where i.txt = m_txt_header.txt
		and i.lang = m_txt_header.lang) 

	insert into m_txt_header
           (txt
           ,[lang]
           ,[header])
     select 
           i.txt
           ,i.lang
           ,i.header
           from inserted i 
GO
CREATE TRIGGER [dbo].[v_txt_header_Trg_delete] ON [dbo].[v_txt_header] INSTEAD OF DELETE AS
	SET NOCOUNT ON
	SET XACT_ABORT ON

	delete from m_txt_header
	where exists (SELECT 1 FROM deleted d
	where d.txt = m_txt_header.txt
		and d.lang = m_txt_header.lang) 
go

--a view to show the groupbox ro (read only) setting for a user. Is used in the Utilities.dll to 
--check whether the controls in a groupbox should be read only or not.

create view v_usr_grp_groupbox as
	select u.usr
		,u.grp
		,f.form
		,m.groupbox
		,m.RO
	from m_usr u inner join m_form_grp f on u.grp = f.grp
	inner join m_form_grp_groupbox m on m.grp = f.grp and m.form=f.form
GO		
		


--DROP PROC [dbo].[p_insert_m_form_grp_temp]
GO
-- called from DialogSelectForms
CREATE PROC [dbo].[p_insert_m_form_grp_temp]

		@grp NVARCHAR (100),
		@form NVARCHAR (100),
		@RO bit 
	AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
	INSERT INTO m_form_grp_temp(grp, form, RO)
	SELECT @grp, @form, @RO
GO

--DROP PROC [dbo].[p_update_m_form_grp]
GO
-- called from DialogSelectForms
CREATE PROC [dbo].[p_update_m_form_grp]
	@grp NVARCHAR (100)
	AS

	--DECLARE @grp  NVARCHAR (100)
	--SET @grp = 'Admin'

	DELETE 
	--SELECT * 
	FROM m_form_grp WHERE 
	grp = @grp AND
	NOT EXISTS (SELECT 1 FROM m_form_grp_temp
		WHERE m_form_grp_temp.grp = m_form_grp.grp AND m_form_grp_temp.form = m_form_grp.form)

	INSERT INTO m_form_grp(grp, form, RO)
	SELECT grp, form, RO FROM m_form_grp_temp
	WHERE NOT EXISTS (SELECT 1 FROM m_form_grp WHERE
	m_form_grp_temp.grp = m_form_grp.grp AND m_form_grp_temp.form = m_form_grp.form)

	UPDATE g SET g.RO = t.RO
	--SELECT *  
	FROM m_form_grp g JOIN m_form_grp_temp t ON t.grp = g.grp AND t.form = g.form
	WHERE t.RO <> g.RO

	DELETE FROM m_form_grp_temp
GO
