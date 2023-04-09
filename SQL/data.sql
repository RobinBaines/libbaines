------------------------------------------------------------------
--
--@Robin Baines 2008
--Create Security tables and log table.
--use utilities
INSERT INTO [m_version]
           ([version]
           ,[multi_lang])
     SELECT 2, 1
     WHERE not exists (SELECT 1 FROM [m_version] WHERE id = 1)
go
UPDATE [m_version]
   SET [version] = 2
GO
INSERT INTO [b_app_parameter]
           ([Parameter]
           ,[ValueString]
           ,[Remark])
     VALUES
           ('auto_logout'
           ,'0'
           ,'Application will stop itself after this time in minutes.')
go

--still to do: update utilities.
INSERT INTO [b_app_parameter]
           ([Parameter]
           ,[ValueString]
           ,[Remark])
     VALUES
           ('auto_logout_warning'
           ,'3'
           ,'Application will warn about stopping itself in this time in minutes before it does stop.')
go
INSERT INTO [b_app_parameter]
           ([Parameter]
           ,[ValueString]
           ,[Remark])
    SELECT 'block_after_days_inactivity'
           ,'5'
           ,'Application will block a user after this many days.'
           WHERE not exists (SELECT 1 FROM [b_app_parameter] WHERE parameter = 'block_after_days_inactivity')
           
go
INSERT INTO [b_app_parameter]
           ([Parameter]
           ,[ValueString]
           ,[Remark])
    SELECT 'delete_usr_log_days'
           ,'50'
           ,'Delete messages in the usr log which are older than this number of days. 0 = do not delete.'
		    WHERE not exists (SELECT 1 FROM [b_app_parameter] WHERE parameter = 'delete_usr_log_days')
go

--use TPIData
go
INSERT INTO [m_lang]
           ([lang])
    SELECT 'FR'
	UNION SELECT 'EN'
	UNION SELECT 'NL'
go
------------------------------------------------------------------
--Create basic data FROM scratch.
------------------------------------------------------------------
--SELECT * FROM m_grp
INSERT INTO [m_grp]
           ([grp])
    SELECT 'Admin' 
    UNION SELECT 'Maintenance'
go

INSERT INTO [m_usr]
           ([usr]
           ,[grp]
           ,[lang]
           )
 SELECT nt_user_name, 'Admin', 'EN' FROM sys.dm_exec_sessions WHERE session_id = @@spid
GO
------delete FROM m_form
INSERT INTO [m_form] (form)
           SELECT 'All' 
go
INSERT INTO [m_form_grp]
           ([grp]
           ,[form]
           ,[RO])
	SELECT 'Admin' ,'All' , 0 
    go
    INSERT INTO [m_format]
           ([format], descr)
     SELECT 'N0', 'An integer.'
		UNION SELECT '', 'Not formatted; just a string.'
		UNION SELECT 'N1', 'Decimal with 1 number after decimal separator.'
		UNION SELECT 'N2', 'Decimal with 2 numbers after decimal separator.'
		UNION SELECT 'N3', 'Decimal with 3 numbers after decimal separator.'
		UNION SELECT 'N4', 'Decimal with 4 numbers after decimal separator.'
		UNION SELECT 'N5', 'Decimal with 5 numbers after decimal separator.'
		UNION SELECT 'N6', 'Decimal with 6 numbers after decimal separator.'
		UNION SELECT '0\%', 'Integer Percentage.'
		UNION SELECT '0.0\%', 'Percentage with 1 number after decimal separator.'
		UNION SELECT '0.00\%', 'Percentage with 2 numbers after decimal separator.'
		UNION SELECT '0.000\%', 'Percentage with 3 numbers after decimal separator.'
		UNION SELECT '0.0000\%', 'Percentage with 4 numbers after decimal separator.'
		
GO
	
--Or for existing datatabases.
--alter table m_format  add descr nvarchar(max)
update m_format set descr = 'Not formatted; just a string.' WHERE [format] =  ''

update m_format set descr = 'An integer.' WHERE [format] = 'N0'
update m_format set descr =  'Decimal with 1 number after decimal separator.' WHERE [format] = 'N1'
update m_format set descr =  'Decimal with 2 numbers after decimal separator.' WHERE [format] = 'N2'
update m_format set descr =  'Decimal with 3 numbers after decimal separator.' WHERE [format] = 'N3'
update m_format set descr =  'Decimal with 4 numbers after decimal separator.' WHERE [format] = 'N4'
update m_format set descr =  'Decimal with 5 numbers after decimal separator.' WHERE [format] = 'N5'
update m_format set descr =  'Decimal with 6 numbers after decimal separator.' WHERE [format] = 'N6'
update m_format set descr =  'Integer Percentage.'  WHERE [format] = '0\%' 
update m_format set descr =   'Percentage with 1 number after decimal separator.' WHERE [format] = '0.0\%'
update m_format set descr =   'Percentage with 2 numbers after decimal separator.' WHERE [format] =  '0.00\%' 
update m_format set descr =  'Percentage with 3 numbers after decimal separator.' WHERE [format] = '0.000\%' 
update m_format set descr =  'Percentage with 4 numbers after decimal separator.' WHERE [format] = '0.0000\%'
 

------------------------------------------------------------------
--Create the utilities texts and tooltips.
------------------------------------------------------------------
---------------------------------------
--Use this to copy FROM one database to another.
---------------------------------------  

--delete FROM [m_form_grp]
--go
--delete FROM [m_usr]
--go
--delete FROM m_tble
--go
--delete FROM m_form
--go
--delete FROM [m_grp]
--go
--INSERT INTO [m_grp]
--           ([grp])
--    SELECT 'Admin' 
--    UNION SELECT 'Maintenance'
--go

    
------delete FROM m_form
--INSERT INTO [m_form] (form)
--           SELECT 'All' 
--go
--INSERT INTO [m_form_grp]
--           ([grp]
--           ,[form]
--           ,[RO])
--	SELECT 'Admin' ,'All' , 0 
--    go
insert into m_tble ([tble]) values ( 'b_app_color')
insert into m_tble ([tble]) values ( 'b_app_parameter')
insert into m_tble ([tble]) values ( 'm_form')
insert into m_tble ([tble]) values ( 'm_form_grp')
insert into m_tble ([tble]) values ( 'm_form_tble')
insert into m_tble ([tble]) values ( 'm_form_tble_column__visibility')
insert into m_tble ([tble]) values ( 'm_format')
insert into m_tble ([tble]) values ( 'm_grp')
insert into m_tble ([tble]) values ( 'm_lang')
insert into m_tble ([tble]) values ( 'm_tble')
insert into m_tble ([tble]) values ( 'm_tble_column')
insert into m_tble ([tble]) values ( 'm_txt')
insert into m_tble ([tble]) values ( 'm_usr')
insert into m_tble ([tble]) values ( 'm_usr_log')
insert into m_tble ([tble]) values ( 'v_tble_column_header')
insert into m_tble ([tble]) values ( 'v_txt_header')
insert into m_tble ([tble]) values ( 'v_usr_log')
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_color','Color','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_color','Parameter','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_color','Remark','',       300)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_color','ValueString','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_parameter','Parameter','',       150)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_parameter','Remark','',       450)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'b_app_parameter','ValueString','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form','form','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_grp','form','',       200)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_grp','grp','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_grp','RO','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble','form','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble','tble','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','bold','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','colmn','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','default_filter','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','form','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','prnt','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','sequence','',        65)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','tble','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_form_tble_column__visibility','visible','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_format','descr','',       300)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_format','format','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_grp','grp','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_lang','lang','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_tble','tble','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_tble_column','colmn','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_tble_column','format','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_tble_column','tble','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_tble_column','width','',        65)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_txt','descr','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_txt','txt','',       200)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_txt','typ','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','blocking','',        30)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','email','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','grp','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','lang','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','name','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','telephone','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr','usr','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','app','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','authorized','',        40)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','createTime','',       120)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','Id','',        50)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','logout','',        40)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','minutes','N0',        40)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','sql_usr','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','usr','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'm_usr_log','windowsIdentity','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_tble_column_header','colmn','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_tble_column_header','header','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_tble_column_header','lang','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_tble_column_header','tble','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_txt_header','descr','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_txt_header','header','',       200)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_txt_header','lang','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_txt_header','txt','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_txt_header','typ','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','app','',       100)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','blocked','',        30)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','blocking','',        30)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','days_since_log','N0',        65)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','logged_in','',        30)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','month_average','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','month_loggedin','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','month_loggedin_unauthorized','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','month_logout','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','month_minutes','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','name','',        80)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','usr','',        80)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','year_average','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','year_loggedin','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','year_loggedin_unauthorized','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','year_logout','N0',        55)
insert into m_tble_column ([tble],[colmn],[format],[width]) values ( 'v_usr_log','year_minutes','N0',        55)

insert into m_tble_column_header (tble, colmn, lang, header) values ( 'b_app_color','Color','EN','Colour')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr','name','EN','name')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr','usr','EN','Windows User')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr_log','authorized','EN','auth')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr_log','createTime','EN','When')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr_log','minutes','EN','mins')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr_log','sql_usr','EN','SQL User')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'm_usr_log','windowsIdentity','EN','Windows Id')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','app','EN','application')

insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','blocked','EN','Is Blocked')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','blocking','EN','blocking activated')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','days_since_log','EN','last logged in days')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','logged_in','EN','Is logged in')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','month_average','EN','Av mins in month')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','month_loggedin','EN','log-in month')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','month_loggedin_unauthorized','EN','log-in unauth month')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','month_minutes','EN','mins month')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','usr','EN','Windows User')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','year_average','EN','Av mins in year')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','year_loggedin','EN','log-in year')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','year_loggedin_unauthorized','EN','log-in unauth year')
insert into m_tble_column_header (tble, colmn, lang, header) values ( 'v_usr_log','year_minutes','EN','mins year')

insert into m_form ([form]) values ( 'User Log')
insert into m_form ([form]) values ( 'Security')

insert into m_form_tble ([form], tble) values ( 'Security','m_form')
insert into m_form_tble ([form], tble) values ( 'Security','m_form_grp')
insert into m_form_tble ([form], tble) values ( 'Security','m_form_tble')
insert into m_form_tble ([form], tble) values ( 'Security','m_form_tble_column__visibility')
insert into m_form_tble ([form], tble) values ( 'Security','m_format')
insert into m_form_tble ([form], tble) values ( 'Security','m_grp')
insert into m_form_tble ([form], tble) values ( 'Security','m_lang')
insert into m_form_tble ([form], tble) values ( 'Security','m_tble')
insert into m_form_tble ([form], tble) values ( 'Security','m_tble_column')
insert into m_form_tble ([form], tble) values ( 'Security','m_txt')
insert into m_form_tble ([form], tble) values ( 'Security','m_usr')
insert into m_form_tble ([form], tble) values ( 'Security','v_tble_column_header')
insert into m_form_tble ([form], tble) values ( 'Security','v_txt_header')
insert into m_form_tble ([form], tble) values ( 'Security','v_usr_log')
insert into m_form_tble ([form], tble) values ( 'User Log','m_usr_log')
insert into m_form_tble ([form], tble) values ( 'User Log','v_usr_log')


insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form','form',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_grp','form',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_grp','grp',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_grp','RO',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble','form',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble','tble',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','bold',         1,         1,        70,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','colmn',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','default_filter',         1,         1,        60,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','form',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','prnt',         1,         1,        40,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','sequence',         1,         1,        50,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','tble',         0,         0,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_form_tble_column__visibility','visible',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_format','descr',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_format','format',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_grp','grp',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_lang','lang',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_tble','tble',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_tble_column','colmn',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_tble_column','format',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_tble_column','tble',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_tble_column','width',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_txt','descr',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_txt','txt',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_txt','typ',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','blocking',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','email',         0,         1,        50,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','grp',         1,         1,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','lang',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','name',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','telephone',         0,         1,        60,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','m_usr','usr',         1,         1,         5,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_tble_column_header','colmn',         0,         0,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_tble_column_header','header',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_tble_column_header','lang',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_tble_column_header','tble',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_txt_header','descr',         0,         0,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_txt_header','header',         1,         1,        40,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_txt_header','lang',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_txt_header','txt',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_txt_header','typ',         0,         0,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','app',         0,         0,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','blocked',         1,         1,       140,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','blocking',         0,         0,       150,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','days_since_log',         1,         1,       130,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','logged_in',         1,         1,       120,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','month_average',         1,         1,       110,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','month_loggedin',         1,         1,        80,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','month_loggedin_unauthorized',         1,         1,        90,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','month_logout',         1,         1,       100,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','month_minutes',         1,         1,        70,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','name',         1,         1,       160,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','usr',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','year_average',         1,         1,        60,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','year_loggedin',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','year_loggedin_unauthorized',         1,         1,        40,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','year_logout',         1,         1,        50,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'Security','v_usr_log','year_minutes',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','app',         0,         0,        10,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','authorized',         1,         1,        60,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','createTime',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','Id',         0,         0,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','logout',         1,         1,        20,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','minutes',         1,         1,        80,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','sql_usr',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','usr',         0,         0,        50,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','m_usr_log','windowsIdentity',         1,         1,        40,         0,'')

insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','blocked',         1,         1,       140,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','blocking',         1,         1,       150,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','days_since_log',         1,         1,       130,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','logged_in',         1,         1,       120,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','month_average',         0,         1,       110,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','month_loggedin',         1,         1,        80,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','month_loggedin_unauthorized',         1,         1,        90,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','month_logout',         0,         0,       100,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','month_minutes',         1,         1,        70,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','name',         1,         1,         5,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','usr',         1,         1,         0,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','year_average',         0,         1,        60,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','year_loggedin',         1,         1,        30,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','year_loggedin_unauthorized',         1,         1,        40,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','year_logout',         0,         0,        50,         0,'')
insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( 'User Log','v_usr_log','year_minutes',         1,         1,        20,         0,'')       
   
 
---------------------------------------------------------------------
--The above can be generated FROM the following or by using 
--procs in align_table_form_formatting.sql
---------------------------------------------------------------------
 -- SELECT 'insert into m_tble ([tble]) values ( ''' + 
	--[tble]
 --      		+ ''')'
 -- FROM m_tble
 -- WHERE tble = 'm_usr_log' or tble = 'v_usr_log'
 --or tble = 'b_app_color'
 --or tble = 'b_app_parameter'
 --or tble = 'm_form'
 --or tble = 'm_form_grp'
 --or tble = 'm_form_tble'
 --or tble = 'm_form_tble_column__visibility'
 --or tble = 'm_format'
 --or tble = 'm_grp'
 -- or tble = 'm_lang'
 --  or tble = 'm_tble'
 --   or tble = 'm_tble_column'
 --    or tble = 'm_tble_column_header'
 --     or tble = 'm_txt'
 --      or tble = 'm_txt_header'
 --       or tble = 'm_usr'
 --         or tble = 'v_usr_log'
 --        or tble = 'v_tble_column_header'
 --         or tble = 'v_txt_header'
  

--INSERT INTO [m_txt_header]
--           ([txt]
--           ,[lang]
--           ,[header])
--SELECT 'Security', 'EN', 'Management'
--UNION SELECT 'Security', 'FR', 'Le Management'
--UNION SELECT 'Events', 'EN', 'Errors etc'
--UNION SELECT 'Events', 'FR', 'Hello'
--UNION SELECT 'Refresh', 'EN', ':-Refresh'
--UNION SELECT 'Refresh', 'FR', ':-Refraiche'
--UNION SELECT 'Production Status', 'EN', 'SFC'
--UNION SELECT 'Production Status', 'FR', 'SFC op zijn Frans'
--UNION SELECT 'Ack Event', 'EN', 'Ack'
--UNION SELECT 'Ack Event', 'FR', 'Le Ack'
--UNION SELECT 'Ack ALL Events', 'EN', 'Ack ALL'
--UNION SELECT 'Ack ALL Events', 'FR', 'Le Ack ALL'


--SELECT 'insert into m_tble_column ([tble],[colmn],[format],[width]) values ( ''' + 
--	[tble]
--       + ''',''' + [colmn]
--       + ''',''' + [format]
--       + ''',' + str([width])
--		+ ')'
--  FROM m_tble_column
--  WHERE tble = 'm_usr_log' or tble = 'v_usr_log'
-- or tble = 'b_app_color'
-- or tble = 'b_app_parameter'
-- or tble = 'm_form'
-- or tble = 'm_form_grp'
-- or tble = 'm_form_tble'
-- or tble = 'm_form_tble_column__visibility'
-- or tble = 'm_format'
-- or tble = 'm_grp'
--  or tble = 'm_lang'
--   or tble = 'm_tble'
--    or tble = 'm_tble_column'
--     or tble = 'm_tble_column_header'
--      or tble = 'm_txt'
--       or tble = 'm_txt_header'
--        or tble = 'm_usr'
--          or tble = 'v_usr_log'
--         or tble = 'v_tble_column_header'
--          or tble = 'v_txt_header'
        
--  --WHERE tble = 'm_usr_log' or tble = 'v_usr_log'
--  go
----go
------------------------------------- 
--  SELECT 'insert into m_tble_column_header (tble, colmn, lang, header) values ( ''' +
--  tble
--       + ''',''' + colmn
--       + ''',''' + lang
--       + ''',''' + header
--       + ''')'
--       FROM m_tble_column_header
--  WHERE tble = 'm_usr_log' or tble = 'v_usr_log'
-- or tble = 'b_app_color'
-- or tble = 'b_app_parameter'
-- or tble = 'm_form'
-- or tble = 'm_form_grp'
-- or tble = 'm_form_tble'
-- or tble = 'm_form_tble_column__visibility'
-- or tble = 'm_format'
-- or tble = 'm_grp'
--  or tble = 'm_lang'
--   or tble = 'm_tble'
--    or tble = 'm_tble_column'
--     or tble = 'm_tble_column_header'
--      or tble = 'm_txt'
--       or tble = 'm_txt_header'
--        or tble = 'm_usr'
--          or tble = 'v_usr_log'
--         or tble = 'v_tble_column_header'
--          or tble = 'v_txt_header'
       
       
---------------------------------------       
--  SELECT 'insert into m_txt (txt,[typ],[descr]) values ( ''' + 
--	[txt]
--       + ''',''' + [typ]
--       + ''',''' + [descr]
--       + ''')'
--       --[form]
--  FROM [m_txt]
----go

--SELECT 'insert into m_form ([form]) values ( ''' + 
--	[form]
--       + ''')'
--       [form]
--  FROM m_form
--    WHERE form = 'User Log' or form = 'Security'
--go
  
  
--  SELECT 'insert into m_form_tble ([form], tble) values ( ''' + 
--	[form]
--	+ ''',''' + [tble]
--       + ''')'
       
--  FROM m_form_tble
--  WHERE form = 'User Log' or form = 'Security'
--go
  

--SELECT 'insert into m_form_tble_column__visibility ([form],[tble],[colmn],[visible],[prnt],[sequence],[bold],[default_filter]) values ( ''' + 
--	[form]
--	+ ''',''' + [tble]
--	+ ''',''' + [colmn]
--	+ ''',' + str(visible)
--	+ ',' + str(prnt)
--	+ ',' + str(sequence)
--	+ ',' + str(bold)
--	+ ',''' + default_filter
--	+ ''')'
--  FROM m_form_tble_column__visibility
--  WHERE form = 'User Log' or form = 'Security'

-- go
 
 
-- -- delete FROM m_tble_column
-- --WHERE tble = 'm_usr_log' or tble = 'v_usr_log'
-- --or tble = 'b_app_color'
-- --or tble = 'b_app_parameter'
-- --or tble = 'm_form'
-- --or tble = 'm_form_grp'
-- --or tble = 'm_form_tble'
-- --or tble = 'm_form_tble_column__visibility'
-- --or tble = 'm_format'
-- --or tble = 'm_grp'
-- -- or tble = 'm_lang'
-- --  or tble = 'm_tble'
-- --   or tble = 'm_tble_column'
-- --    or tble = 'm_tble_column_header'
-- --     or tble = 'm_txt'
-- --      or tble = 'm_txt_header'
-- --       or tble = 'm_usr'
-- --       or tble = 'v_usr_log'
-- --        or tble = 'v_tble_column_header'
-- --         or tble = 'v_txt_header'
        
   
 