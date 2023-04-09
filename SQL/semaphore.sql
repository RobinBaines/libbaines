------------------------------------------------------------------
--@Robin Baines 2008
--semaphore table.
--20090202 RPB created. 
------------------------------------------------------------------
--use Utilities
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[b_semaphore]') AND type in (N'U'))
drop TABLE [dbo].[b_semaphore]
go
CREATE TABLE [dbo].[b_semaphore](
       [app] [nvarchar](50) not null,
       [tble] [nvarchar](50) not null,
       [semaphore] [int] NULL,
PRIMARY KEY (app, tble)
) ON [PRIMARY]
go

------------------------------
--Example of how to set up. 
------------------------------
--INSERT INTO [b_semaphore]
--           ([app]
--           ,[tble]
--           ,[semaphore])
--    select 'SomeApp', 'SomeEvent', 1  
--go 
   
--IF  EXISTS (SELECT * FROM sys.triggers WHERE name= 'DEGUSSA_LOG_EVENT_insert_update_delete')
--DROP trigger DEGUSSA_LOG_EVENT_insert_update_delete
--go

--CREATE TRIGGER SOME_TABLE_insert_update_delete ON SOME_TABLE AFTER INSERT, UPDATE, DELETE
--AS
--	SET NOCOUNT ON
--	SET XACT_ABORT ON
--	update b_semaphore set semaphore = semaphore + 1 where
--	app='SomeApp' and tble = 'SomeEvent' 
--GO