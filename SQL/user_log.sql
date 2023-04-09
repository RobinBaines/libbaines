------------------------------------------------------------------
--@Robin Baines 2008
--User log
--20100318
--20100817 RPB Modified v_usr_log. Only set blocked if the m_usr.blocking is 1; meaning that blocking is activated for the user.
--DEPENDS ON Utilities.sql and appParameters.sql and Functions.sql
--Also requires SQL 2005 compatability level for example [v_usr_log0].
--20120716 RPB removed all references to bUsers table.
------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].m_usr_log') AND type in (N'U'))
DROP table m_usr_log
go
	CREATE TABLE m_usr_log 	(
		Id [int] IDENTITY(1,1) NOT NULL,
		[app] [nvarchar](50) not null,
		logout bit default 0,
		sql_usr NVARCHAR (100) default left(SUSER_SNAME(), 128),
		windowsIdentity NVARCHAR (100),
		usr NVARCHAR (100),
		authorized bit default 0,
		createTime datetime NOT NULL DEFAULT (getdate()),
		minutes	int default 0,
		PRIMARY KEY (Id)
	)
GO
--Delete usr log entries if they are older than the parameter in delete_usr_log_days and if delete_usr_log_days is not 0 or any none integer.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].p_delete_usr_log') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[p_delete_usr_log]
GO
CREATE proc p_delete_usr_log
	as
	SET NOCOUNT ON
	SET XACT_ABORT ON
	
	--use of function prevents errors if parameter in valuestring is not an integer.	
	delete from m_usr_log where id in (select id from  m_usr_log
		inner join  b_app_parameter on parameter = 'delete_usr_log_days' and [dbo].UDF_ParseNumChars(valuestring) <> '0'
		and [dbo].UDF_ParseNumChars(valuestring) <> ''
	where datediff(day, createtime, getdate()) > cast(valuestring as int))
go
--Insert a new log message. 
--20100818 But only set authorized if not blocked.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].p_usr_log') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[p_usr_log]
GO
CREATE proc p_usr_log
	@app as [nvarchar](50),
	@logout as bit,
	@windowsIdentity NVARCHAR (100),
	@usr NVARCHAR (100),
	@minutes as int,
	@blocked as bit 
	as
	SET NOCOUNT ON
	SET XACT_ABORT ON
	Declare @authorized as bit
	Set @authorized = 0
	Set @authorized = (case when @blocked = 1 then 0 else isnull((select 1 from m_usr where usr = @usr), 0) end)
		
	insert into m_usr_log (app, logout, windowsIdentity, usr, authorized, minutes)
	values(@app, @logout, @windowsIdentity, @usr, @authorized, @minutes)
	select @@rowcount
go
--exec [p_usr_log] 'TestApp', 1, 'RPB4\Robin', 'Robin', 57, 1

--Calculate several counts on the usr log.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].fn_Usr_usage')  AND type in (N'IF'))
DROP FUNCTION dbo.fn_Usr_usage
GO
CREATE FUNCTION dbo.fn_Usr_usage(@fromdate as datetime)
RETURNS TABLE 
AS RETURN 

	SELECT distinct u.usr, u.app 
		, (select sum(minutes) from [m_usr_log]	where createtime >= @fromdate and usr = u.usr and app = u.app) as minutes
		, (select count(*) from [m_usr_log]	where createtime >= @fromdate and usr = u.usr and app = u.app and logout = 0) as loggedin
		, (select count(*) from [m_usr_log]	where createtime >= @fromdate and usr = u.usr and app = u.app and authorized = 0 and logout = 1) as loggedin_unauthorized
		, (select count(*) from [m_usr_log]	where createtime >= @fromdate and usr = u.usr and app = u.app and logout = 1) as logout
		, (select avg(minutes) from [m_usr_log]	where createtime >= @fromdate and usr = u.usr and app = u.app and logout = 1) as average
	from [m_usr_log] u 
	where u.createtime >=@fromdate
go
--declare @fromdate as datetime
--set @fromdate = dateadd(day, -10, getdate())
--select * from fn_Usr_usage( @fromdate)

--declare @fromdate as datetime
--set @fromdate = dateadd(year, -1, getdate())
--declare @fromdate_month as datetime
--set @fromdate_month = dateadd(month, -1, getdate())
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr_log') AND type in (N'V'))
drop view v_usr_log 
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr_log0') AND type in (N'V'))
drop view v_usr_log0 
go
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create view v_usr_log0 as
	select 
	u1.usr
	, u1.app
	, u1.minutes as year_minutes
	, u1.loggedin as year_loggedin
	, u1.loggedin_unauthorized as year_loggedin_unauthorized
	, u1.logout as year_logout
	, u1.average as year_average
	, isnull(u2.minutes,0) as month_minutes
	, isnull(u2.loggedin,0) as month_loggedin
	, isnull(u2.loggedin_unauthorized,0) as month_loggedin_unauthorized
	, isnull(u2.logout,0) as month_logout
	, isnull(u2.average,0) as month_average
	from 
	fn_Usr_usage( dateadd(year, -1, getdate()) ) u1
	left join fn_Usr_usage(dateadd(month, -1, getdate())) u2
	on u1.usr = u2.usr and u1.app = u2.app
	
go
--Get last log action of each usr/app today.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_user_log_last_action_today') AND type in (N'V'))
drop view v_user_log_last_action_today 
go
create view v_user_log_last_action_today as 
	select max(createtime) as createtime, app, usr
	from [m_usr_log]
	where datediff(day, createtime ,getdate()) = 0
	and authorized = 1
	group by app, usr
go
--Get the last authorized log action for a usr, app if it was a login action.
--Therefore finds those users who are still logged into an app today.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_user_logged_in') AND type in (N'V'))
drop view v_user_logged_in 
go
create view v_user_logged_in as 
	select b.id, a.createtime, a.app, a.usr
	, b.logout from v_user_log_last_action_today a inner join [m_usr_log] b
	on a.createtime =b.createtime and a.app = b.app and a.usr = b.usr
	where logout = 0 
go
--last authorized log action (in or out).
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_user_log_last_action') AND type in (N'V'))
drop view v_user_log_last_action 
go
create view v_user_log_last_action as 
	select max(createtime) as createtime, app, usr
	from [m_usr_log]
	where authorized = 1
	group by app, usr
go
--create some fields such as blocked, logged_in
--20100817 RPB Modified v_usr_log. Only set blocked if the m_usr.blocking is 1; meaning that blocking is activated for the user.
--Do not set blocked if user is not authorized.
create view v_usr_log as
	SELECT l.[usr]
	  ,l.app
      ,[year_minutes]
      ,[year_loggedin]
      ,[year_loggedin_unauthorized]
      ,[year_logout]
      ,[year_average]
      ,[month_minutes]
      ,[month_loggedin]
      ,[month_loggedin_unauthorized]
      ,[month_logout]
      ,[month_average]
	  --set to 1 if the usr is still logged into the app.
	  , cast(case when len(isnull(i.usr, '')) > 0 then 1 else 0 end as bit) as logged_in
	  , case when a.createtime is null then -1 else datediff(day, a.createtime, getdate()) end as days_since_log
	  
	  --Only set blocked if the m_usr.blocking is 1.
	  , cast(
			case when p.parameter is null then 0 else 
				case when a.createtime is null then 0 else
					case when datediff(day, a.createtime, getdate()) > cast(p.valuestring as int) 
					then isnull(m.blocking, 1) 
					else 0 
					end
				end 
			end
		as bit) as blocked
	  , cast(isnull(m.blocking, 0) as bit) as blocking
	  , isnull(m.name, '') as name
	FROM [v_usr_log0] l
	left join v_user_logged_in i
		on l.usr= i.usr and l.app = i.app
	left join v_user_log_last_action a
		on l.usr= a.usr and l.app = a.app
	left join [b_app_parameter] p on p.parameter = 'block_after_days_inactivity'
	left join m_usr m on m.usr = l.usr
		union 
	select a.usr, '', 0,0,0,0,0,0,0,0,0,0,cast(0 as bit),-1, cast(0 as bit)
	,cast(0 as bit), ''
	 from v_usr a left join [v_usr_log0] b
	on a.usr = b.usr
	where b.usr is null
go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_usr_blocked') AND type in (N'V'))
drop view v_usr_blocked
go
create view v_usr_blocked as
	select u.usr
		,l.app
		,u.grp
		,u.lang
		,u.blocking
		,u.name
		,u.email
		,u.telephone
		,cast(isnull(l.logged_in, 0) as bit) as logged_in
		,cast(isnull(l.days_since_log, 0) as int) as days_since_log
		,cast(isnull(l.blocked, 0) as bit) as blocked
	from m_usr u 
	inner join m_version on multi_lang = 1
	left join v_usr_log l on u.usr = l.usr
go
