/*
	Filename: functions.sql
	Project: 
	Date: 20100222
	Author: RPB
	20130131 RPB added GetDayOfWeek which accepts dow integer and returns translated text.
	-- 20200202 Modified [f_get_translatedtext] and [f_get_translated_column_header]

*/
--Use Utilities
GO

-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Function to translate a string using m_txt_header. 
-- 20180219 This function was originally defined in utilities. Was modified to use usr GIS as a 2nd effort to translate because GISEvents service uses 
-- a strange alias and was defaulting to the original @text (English).
-- 20200202 Replaced RIGHT(suser_name(), len(u.usr)) = u.usr by RIGHT(SUSER_NAME(), LEN(SUSER_NAME()) - CHARINDEX('\',SUSER_NAME())) = u.usr
-- because the first effort will match baines if SUSER_NAME is oaBaines.
-- Used By FUNCTION [dbo].[GetDayOfWeek] and application views.
--SELECT [dbo].GetDayOfWeek(1)
-- =============================================
CREATE FUNCTION [dbo].[f_get_translatedtext](@text AS NVARCHAR(100))
RETURNS NVARCHAR(250)
BEGIN
	DECLARE @translatedtext NVARCHAR(250)
	SET @translatedtext= ISNULL(
			(SELECT header FROM m_txt_header h
				INNER JOIN v_usr u ON  RIGHT(SUSER_NAME(), LEN(SUSER_NAME()) - CHARINDEX('\',SUSER_NAME())) = u.usr
				--RIGHT(suser_name(), len(u.usr)) = u.usr
				WHERE u.lang = h.lang AND h.txt = @text)

				, ISNULL(
					(SELECT TOP 1 h.header FROM dbo.m_txt_header h INNER JOIN v_usr u ON u.usr = 'GIS' WHERE u.lang = h.lang AND h.txt = @text ORDER BY u.usr)
					,  @text))
	RETURN @translatedtext
END
GO

-- =============================================
-- Author:  RPB
-- Create date: 
-- Description: Function to translate a string using m_txt_header. 
-- 20200202 Replaced RIGHT(suser_name(), len(u.usr)) = u.usr by RIGHT(SUSER_NAME(), LEN(SUSER_NAME()) - CHARINDEX('\',SUSER_NAME())) = u.usr
-- because the first effort will match baines if SUSER_NAME is oaBaines.
-- Used by application views.
-- =============================================
CREATE FUNCTION [dbo].[f_get_translated_column_header](@tble AS NVARCHAR(100), @text AS NVARCHAR(100))
RETURNS NVARCHAR(250)
BEGIN
	DECLARE @translatedtext NVARCHAR(250)
	SET @translatedtext=ISNULL((SELECT header FROM m_tble_column_header h
		INNER JOIN v_usr u ON  RIGHT(SUSER_NAME(), LEN(SUSER_NAME()) - CHARINDEX('\',SUSER_NAME())) = u.usr
		--inner join v_usr u on right(suser_name(), len(usr)) = usr
		WHERE u.lang = h.lang AND h.tble = @tble AND h.colmn = @text ), @text)
	RETURN @translatedtext
END
GO

--COMPARE THE ABOVE WITH THE ORIGINALS WHEN SUSER_NAME() = 'oa_Baines'
--IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].f_get_translatedtext_old')  AND type in (N'IF', 'FN'))
--DROP function dbo.f_get_translatedtext_old
--go
----return translated text in users language. return the original if not found.
--CREATE FUNCTION  dbo.f_get_translatedtext_old(@text as nvarchar(100))
--RETURNS nvarchar(250)
--BEGIN
--	declare @translatedtext nvarchar(250)

--	set @translatedtext=isnull((select header from m_txt_header h
--		inner join v_usr u on right(suser_name(), len(usr)) = usr
--		where u.lang = h.lang and h.txt = @text), @text)
--	return @translatedtext
--END
--GO


--IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].f_get_translated_column_header_old')  AND type in (N'IF', 'FN'))
--DROP function dbo.f_get_translated_column_header_old
--go
----return translated column header in users language. return the original if not found.
--CREATE FUNCTION dbo.f_get_translated_column_header_old(@tble AS NVARCHAR(100), @text AS NVARCHAR(100))
--RETURNS nvarchar(250)
--BEGIN
--	DECLARE @translatedtext nvarchar(250)

--	SET @translatedtext=isnull((select header from m_tble_column_header h
--		inner join v_usr u on right(suser_name(), len(usr)) = usr
--		where u.lang = h.lang AND h.tble = @tble AND h.colmn = @text ), @text)
--	RETURN @translatedtext
--END
--GO

--SELECT [dbo].[f_get_translated_column_header_old]('b_activity', 'activity')
--SELECT [dbo].[f_get_translatedtext_old]('activity')

--DROP function dbo.f_get_translatedtext_old
--DROP function dbo.f_get_translated_column_header_old
----------------------------------------------------------------------

--get the day of the week name translating where necessary.
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].GetDayOfWeek')  AND type in (N'IF', 'FN'))
drop function dbo.GetDayOfWeek
go
CREATE FUNCTION dbo.GetDayOfWeek(@dow AS INT)
 RETURNS VARCHAR(10)
 AS
 BEGIN
 DECLARE @rtDayofWeek VARCHAR(10)
 SELECT @rtDayofWeek = CASE @dow
 WHEN 1 THEN 'sun'
 WHEN 2 THEN 'mon'
 WHEN 3 THEN 'tue'
 WHEN 4 THEN 'wed'
 WHEN 5 THEN 'thu'
 WHEN 6 THEN 'fri'
 WHEN 7 THEN 'sat'
 END
 RETURN dbo.f_get_translatedtext(@rtDayofWeek)
 END
 GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_remove_lzeros]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].[f_remove_lzeros] 
GO
CREATE FUNCTION [dbo].[f_remove_lzeros] (
	@string AS varchar(max)
)	
RETURNS varchar(max)
AS
BEGIN
	DECLARE @result AS varchar(max)
	DECLARE @int AS int
	DECLARE @IncorrectCharLoc SMALLINT
	SET @result = @string
	SET @IncorrectCharLoc = PATINDEX('%[^0-9]%', @string)
	if @IncorrectCharLoc = 0
	begin
		SET @int = cast(@string as int)
		SET @result = ltrim(str(@int))
	end
	RETURN @result
END
GO
--select [dbo].[f_remove_lzeros]('0001234')
--select [dbo].[f_remove_lzeros]('0001234.0')
--select [dbo].[f_remove_lzeros]('1234')

--20110502 Removed function to return iso 112 and 12 dates. Use a cast or convert.
--This looks better. 112 = yyyymmdd and 12 is yymmdd.
--See CAST and CONVERT Trans SQL help. 112 and 12 are the ISO sortable formats.
--select convert(varchar, GETDATE(), 112)
-------------------------------------
--Return hh:mm for example 12:30
-------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn_ShortTime]')   AND type in (N'IF', 'FN'))
DROP FUNCTION dbo.[fn_ShortTime]
GO
CREATE FUNCTION [dbo].[fn_ShortTime] (@thisDate datetime)
RETURNS varchar(50)
AS
BEGIN
	DECLARE @ResultVar varchar(50)
	SELECT @ResultVar=substring(ltrim(str(DatePart(hour, @thisDate)+100)),2, 2)
			+ ':' + substring(ltrim(str(DatePart(minute, @thisDate)+100)),2, 2)
	RETURN @ResultVar
END
go
--select dbo.[fn_ShortTime](getdate())
--select dbo.[fn_ShortTime](getdate())

-----------------------------------------------
--Filter out none num characters. Is used in user
-----------------------------------------------
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UDF_ParseNumChars]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].UDF_ParseNumChars
GO
CREATE FUNCTION dbo.UDF_ParseNumChars
(
@string VARCHAR(8000)
)
RETURNS VARCHAR(8000)
AS
BEGIN
DECLARE @IncorrectCharLoc SMALLINT
SET @IncorrectCharLoc = PATINDEX('%[^0-9]%', @string)
WHILE @IncorrectCharLoc > 0
BEGIN
SET @string = STUFF(@string, @IncorrectCharLoc, 1, '')
SET @IncorrectCharLoc = PATINDEX('%[^0-9]%', @string)
END
SET @string = @string
RETURN @string
END
GO
--SELECT dbo.UDF_ParseNumChars('ABC”_I+{D[]}4|:e;””5,<.F>/?6')
GO
--only difference is the pattern which includes a '.'
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UDF_ParseFloats]')  AND type in (N'IF', 'FN'))
DROP FUNCTION [dbo].UDF_ParseFloats
GO
CREATE FUNCTION dbo.UDF_ParseFloats
(
@string VARCHAR(8000)
)
RETURNS VARCHAR(8000)
AS
BEGIN
DECLARE @IncorrectCharLoc SMALLINT
SET @IncorrectCharLoc = PATINDEX('%[^0-9.,]%', @string)
WHILE @IncorrectCharLoc > 0
BEGIN
SET @string = STUFF(@string, @IncorrectCharLoc, 1, '')
SET @IncorrectCharLoc = PATINDEX('%[^0-9.,]%', @string)
END
SET @string = @string
RETURN @string
END
GO

--sometimes this is safer.
--select case when PATINDEX('%[^0-9]%', '123.3')= 0 then str('1233') else 0 end
--select case when PATINDEX('%[^0-9]%', '1233')= 0 then str('1233') else 0 end


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn_SubtractDates]')  AND type in (N'IF', 'FN'))
drop function dbo.fn_SubtractDates
go
create function dbo.fn_SubtractDates(@Month as varchar(8), @Month2 as varchar(8))
returns int
AS
BEGIN
	Declare @ResultVar int
	Declare @fMonth as int
	Declare @fYear as int
	Declare @fMonth2 as int
	Declare @fYear2 as int

	set @fMonth = cast(substring(@Month, 5, 2) as int)
	set @fYear = cast(substring(@Month, 1, 4) as int)
	set @fMonth2 = cast(substring(@Month2, 5, 2) as int)
	set @fYear2 = cast(substring(@Month2, 1, 4) as int)

	set @ResultVar = (@fyear - @fyear2) * 12 + (@fMonth - @fMonth2)
	RETURN  @ResultVar
END
GO
