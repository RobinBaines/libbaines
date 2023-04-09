------------------------------------------------------------------
--@Robin Baines 2011
--UTILITIES application log
------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_app_log') AND type in (N'U'))
DROP table m_app_log
go
	CREATE TABLE m_app_log 	(
		Id [int] IDENTITY(1,1) NOT NULL,
		app [nvarchar](50) null,
		usr [nvarchar](50) null,
		error nvarchar(max) null,
		raised_in [nvarchar](50) null,
		priority int default 0,
		createTime datetime NOT NULL DEFAULT (getdate()),
		PRIMARY KEY (Id)
	)
GO
INSERT INTO [b_app_parameter]
           ([Parameter]
           ,[ValueString]
           ,[Remark])
    select 'delete_app_log_days'
           ,'50'
           ,'Delete messages in the application log which are older than this number of days. 0 = do not delete.'
		    where not exists (select 1 from [b_app_parameter] where parameter = 'delete_app_log_days')
go

--Delete usr log entries if they are older than the parameter in delete_usr_log_days and if delete_usr_log_days is not 0 or any none integer.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].p_delete_app_log') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].p_delete_app_log
GO
CREATE proc p_delete_app_log
	as
	SET NOCOUNT ON
	SET XACT_ABORT ON
	
	--use of function prevents errors if parameter in valuestring is not an integer.	
	delete from m_app_log where id in (select id from  m_app_log
		inner join  b_app_parameter on parameter = 'delete_app_log_days' and [dbo].UDF_ParseNumChars(valuestring) <> '0'
		and [dbo].UDF_ParseNumChars(valuestring) <> ''
	where datediff(day, createtime, getdate()) > cast(valuestring as int))
go
INSERT INTO [b_semaphore]
           ([app]
           ,[tble]
           ,[semaphore])
    select 'Utilities', 'app_log', 1  
go 

CREATE TRIGGER m_app_log_insert_update_delete ON m_app_log AFTER INSERT, UPDATE, DELETE
AS
	SET NOCOUNT ON
	SET XACT_ABORT ON
	update b_semaphore set semaphore = semaphore + 1 where
	app='Utilities' and tble = 'app_log' 
GO