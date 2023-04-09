/*
Filename: AppParameters.sql
Project: 
Date: 20090501
Author and Copyright: Robin Baines
Purpose: The application paramters.
Modifications:
20100920 Added b_app_color
20101115 RPB Modified p_get_app_parameter to use v_app_parameter.
*/
--use Utilities
--go
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].b_app_parameter') AND type in (N'U'))
DROP TABLE b_app_parameter 	
go
	CREATE TABLE b_app_parameter 	(
		Parameter varchar(32) PRIMARY KEY,
		ValueString nvarchar(MAX) default '',
		Remark nvarchar(MAX) NULL
		)
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].v_app_parameter') 
AND type in (N'V'))
drop view v_app_parameter
GO
------------------------------------
--v_app_parameter
--RPB Created 20101115
------------------------------------
Create view [dbo].[v_app_parameter] as 
	select 
		parameter
		--removed this method of having parameter with server name which needs to
		--be remapped to back up server if that is running.
		--,case when charindex('\\xxxxxx005m', valuestring) > 0 then 
		--	left(valuestring,charindex('\\xxxxxx005m', valuestring) - 1 ) 
		--	+ '\\' + cast (SERVERPROPERTY('ComputerNamePhysicalNetBIOS') as nvarchar)
		--	+ right(valuestring, len(valuestring) - (len('\\xxxxxx005m') + charindex('\\xxxxxx005m', valuestring) - 1) ) 
		--	else valuestring end as valuestring
		, valuestring
		, valuestring as org_valuestring
		, remark
	from b_app_parameter
	GO
------------------------------------
--p_get_app_parameter
--get a parameter and used from gui.
--20101115 RPB Modified p_get_app_parameter to use v_app_parameter.
---------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[p_get_app_parameter]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[p_get_app_parameter]
GO
CREATE proc p_get_app_parameter 
	@Parameter as varchar(32),
	@ValueString as nvarchar(max) OUTPUT
	as
	set @ValueString = isnull((select ValueString from v_app_parameter where Parameter =@Parameter), '')

go
---------------------------
--
---------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].b_app_color') AND type in (N'U'))
DROP TABLE b_app_color
go
	CREATE TABLE b_app_color 	(
		Parameter varchar(32) PRIMARY KEY,
		ValueString nvarchar(50) default '',
		Remark nvarchar(MAX) NULL
		)
GO

------------------------------------
--p_get_app_color
--get a parameter and used from gui.
---------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[p_get_app_color]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[p_get_app_color]
GO
CREATE proc p_get_app_color 
	@Parameter as varchar(32),
	@ValueString as nvarchar(max) OUTPUT
	as
	set @ValueString = isnull((select ValueString from b_app_color where Parameter =@Parameter), '')
go

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_get_app_parameter_float]')  AND type in (N'IF', 'FN'))
drop FUNCTION [dbo].f_get_app_parameter_float
go
create FUNCTION [dbo].f_get_app_parameter_float (@Parameter as varchar(32),
	@Default as float)
RETURNS float
AS
BEGIN
	DECLARE @ResultVar float
	SEt @ResultVar= isnull((select case when PATINDEX('%[^0-9.]%', [ValueString])= 0 then cast([ValueString] as float) else 0 end
		from b_app_parameter where Parameter = @Parameter), @Default)
	RETURN @ResultVar
END
go
--select [dbo].f_get_app_parameter_float('margin_cfl', 100.501)
--select [dbo].f_get_app_parameter_float('undefined', 100.501)





