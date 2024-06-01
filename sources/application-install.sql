-- =============================================
-- Gartle CRM
-- Version 2.1, December 13, 2022
--
-- Copyright 2020-2022 Gartle LLC
--
-- License: MIT
-- =============================================

SET NOCOUNT ON
GO

CREATE SCHEMA [gcrm]
GO

CREATE TABLE [gcrm].[busy_statuses] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_busy_statuses] PRIMARY KEY ([id])
    , CONSTRAINT [IX_busy_statuses_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[folder_types] (
      [id] tinyint NOT NULL
    , [folder_type] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_folder_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_folder_types_folder_type] UNIQUE ([folder_type])
);
GO

CREATE TABLE [gcrm].[formats] (
      [ID] int IDENTITY(1,1) NOT NULL
    , [TABLE_SCHEMA] nvarchar(128) NOT NULL
    , [TABLE_NAME] nvarchar(128) NOT NULL
    , [TABLE_EXCEL_FORMAT_XML] xml NULL
    , CONSTRAINT [PK_formats] PRIMARY KEY ([ID])
    , CONSTRAINT [IX_formats] UNIQUE ([TABLE_NAME], [TABLE_SCHEMA])
);
GO

CREATE TABLE [gcrm].[handlers] (
      [ID] int IDENTITY(1,1) NOT NULL
    , [TABLE_SCHEMA] nvarchar(20) NULL
    , [TABLE_NAME] nvarchar(128) NOT NULL
    , [COLUMN_NAME] nvarchar(128) NULL
    , [EVENT_NAME] varchar(25) NULL
    , [HANDLER_SCHEMA] nvarchar(20) NULL
    , [HANDLER_NAME] nvarchar(128) NULL
    , [HANDLER_TYPE] varchar(25) NULL
    , [HANDLER_CODE] nvarchar(max) NULL
    , [TARGET_WORKSHEET] nvarchar(128) NULL
    , [MENU_ORDER] int NULL
    , [EDIT_PARAMETERS] bit NULL
    , CONSTRAINT [PK_handlers] PRIMARY KEY ([ID])
    , CONSTRAINT [IX_handlers] UNIQUE ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME])
);
GO

CREATE TABLE [gcrm].[importance_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_importance_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_importance_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[journal_recipient_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_journal_recipient_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_journal_recipient_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[mail_recipient_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_mail_recipient_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_mail_recipient_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[meeting_recipient_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_meeting_recipient_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_meeting_recipient_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[objects] (
      [ID] int IDENTITY(1,1) NOT NULL
    , [TABLE_SCHEMA] nvarchar(128) NOT NULL
    , [TABLE_NAME] nvarchar(128) NOT NULL
    , [TABLE_TYPE] nvarchar(128) NOT NULL
    , [TABLE_CODE] nvarchar(max) NULL
    , [INSERT_OBJECT] nvarchar(max) NULL
    , [UPDATE_OBJECT] nvarchar(max) NULL
    , [DELETE_OBJECT] nvarchar(max) NULL
    , CONSTRAINT [PK_objects] PRIMARY KEY ([ID])
    , CONSTRAINT [IX_objects] UNIQUE ([TABLE_NAME], [TABLE_SCHEMA])
);
GO

CREATE TABLE [gcrm].[sensitivity_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_sensitivity_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_sensitivity_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[task_recipient_types] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_task_recipient_types] PRIMARY KEY ([id])
    , CONSTRAINT [IX_task_recipient_types_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[task_statuses] (
      [id] tinyint NOT NULL
    , [name] nvarchar(100) NOT NULL
    , CONSTRAINT [PK_task_statuses] PRIMARY KEY ([id])
    , CONSTRAINT [IX_task_statuses_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[translations] (
      [ID] int IDENTITY(1,1) NOT NULL
    , [TABLE_SCHEMA] nvarchar(128) NULL
    , [TABLE_NAME] nvarchar(128) NULL
    , [COLUMN_NAME] nvarchar(128) NULL
    , [LANGUAGE_NAME] char(2) NOT NULL
    , [TRANSLATED_NAME] nvarchar(128) NULL
    , [TRANSLATED_DESC] nvarchar(1024) NULL
    , [TRANSLATED_COMMENT] nvarchar(2000) NULL
    , CONSTRAINT [PK_translations] PRIMARY KEY ([ID])
    , CONSTRAINT [IX_translations] UNIQUE ([TABLE_NAME], [TABLE_SCHEMA], [COLUMN_NAME], [LANGUAGE_NAME])
);
GO

CREATE TABLE [gcrm].[users] (
      [id] smallint NOT NULL
    , [name] nvarchar(128) NOT NULL
    , CONSTRAINT [PK_users] PRIMARY KEY ([id])
    , CONSTRAINT [IX_users_name] UNIQUE ([name])
);
GO

CREATE TABLE [gcrm].[workbooks] (
      [ID] int IDENTITY(1,1) NOT NULL
    , [NAME] nvarchar(128) NOT NULL
    , [TEMPLATE] nvarchar(255) NULL
    , [DEFINITION] nvarchar(max) NOT NULL
    , [TABLE_SCHEMA] nvarchar(128) NULL
    , CONSTRAINT [PK_workbooks] PRIMARY KEY ([ID])
    , CONSTRAINT [IX_workbooks] UNIQUE ([NAME])
);
GO

CREATE TABLE [gcrm].[folders] (
      [id] int IDENTITY(1,1) NOT NULL
    , [user_id] smallint NOT NULL
    , [folder_type_id] tinyint NOT NULL
    , [StoreID] char(188) NOT NULL
    , [EntryID] char(48) NOT NULL
    , [Name] nvarchar(255) NOT NULL
    , [FolderPath] nvarchar(255) NOT NULL
    , CONSTRAINT [PK_folders] PRIMARY KEY ([id])
    , CONSTRAINT [IX_folders] UNIQUE ([user_id], [StoreID], [EntryID])
);
GO

ALTER TABLE [gcrm].[folders] ADD CONSTRAINT [FK_folders_folder_types] FOREIGN KEY ([folder_type_id]) REFERENCES [gcrm].[folder_types] ([id]) ON UPDATE CASCADE;
GO

ALTER TABLE [gcrm].[folders] ADD CONSTRAINT [FK_folders_users] FOREIGN KEY ([user_id]) REFERENCES [gcrm].[users] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[appointments] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [AllDayEvent] bit NULL
    , [Attachments_Count] tinyint NULL
    , [BillingInformation] nvarchar(255) NULL
    , [Body] nvarchar(max) NULL
    , [BusyStatus] tinyint NULL
    , [Categories] nvarchar(255) NULL
    , [ConversationTopic] nvarchar(255) NULL
    , [CreationTime] datetime NULL
    , [Duration] int NULL
    , [End] datetime NULL
    , [Importance] tinyint NULL
    , [IsRecurring] bit NULL
    , [LastModificationTime] datetime NULL
    , [Location] nvarchar(255) NULL
    , [MeetingStatus] tinyint NULL
    , [MessageClass] nvarchar(100) NULL
    , [Mileage] nvarchar(50) NULL
    , [NoAging] bit NULL
    , [OptionalAttendees] nvarchar(255) NULL
    , [Organizer] nvarchar(255) NULL
    , [RecurrencePattern_RecurrenceType] tinyint NULL
    , [RecurrencePattern_PatternEndDate] datetime NULL
    , [RecurrencePattern_PatternStartDate] datetime NULL
    , [ReminderMinutesBeforeStart] int NULL
    , [ReminderSet] bit NULL
    , [ReminderOverrideDefault] bit NULL
    , [ReminderPlaySound] bit NULL
    , [RequiredAttendees] nvarchar(255) NULL
    , [Resources] nvarchar(255) NULL
    , [ResponseRequested] bit NULL
    , [Sensitivity] tinyint NULL
    , [Size] int NULL
    , [Start] datetime NULL
    , [Subject] nvarchar(255) NULL
    , [UnRead] bit NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_appointments] PRIMARY KEY ([id])
    , CONSTRAINT [IX_appointments_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[appointments] ADD CONSTRAINT [FK_appointments_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[contacts] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [Account] nvarchar(255) NULL
    , [Anniversary] datetime NULL
    , [AssistantName] nvarchar(255) NULL
    , [AssistantTelephoneNumber] nvarchar(50) NULL
    , [Attachments_Count] tinyint NULL
    , [BillingInformation] nvarchar(255) NULL
    , [Birthday] datetime NULL
    , [Body] nvarchar(max) NULL
    , [BusinessAddress] nvarchar(255) NULL
    , [BusinessAddressCity] nvarchar(50) NULL
    , [BusinessAddressCountry] nvarchar(50) NULL
    , [BusinessAddressPostOfficeBox] nvarchar(50) NULL
    , [BusinessAddressPostalCode] nvarchar(50) NULL
    , [BusinessAddressState] nvarchar(50) NULL
    , [BusinessAddressStreet] nvarchar(255) NULL
    , [BusinessFaxNumber] nvarchar(50) NULL
    , [BusinessHomePage] nvarchar(50) NULL
    , [BusinessTelephoneNumber] nvarchar(50) NULL
    , [Business2TelephoneNumber] nvarchar(50) NULL
    , [CallbackTelephoneNumber] nvarchar(50) NULL
    , [CarTelephoneNumber] nvarchar(50) NULL
    , [Categories] nvarchar(255) NULL
    , [Children] nvarchar(255) NULL
    , [CompanyName] nvarchar(255) NULL
    , [CompanyMainTelephoneNumber] nvarchar(50) NULL
    , [ComputerNetworkName] nvarchar(50) NULL
    , [CreationTime] datetime NULL
    , [CustomerID] nvarchar(50) NULL
    , [Department] nvarchar(255) NULL
    , [Email1Address] nvarchar(100) NULL
    , [Email2Address] nvarchar(100) NULL
    , [Email3Address] nvarchar(100) NULL
    , [FileAs] nvarchar(255) NULL
    , [FirstName] nvarchar(50) NULL
    , [FlagRequest] nvarchar(50) NULL
    , [FTPSite] nvarchar(100) NULL
    , [FullName] nvarchar(255) NULL
    , [Gender] tinyint NULL
    , [GovernmentIDNumber] nvarchar(50) NULL
    , [Hobby] nvarchar(255) NULL
    , [HomeAddress] nvarchar(255) NULL
    , [HomeAddressCity] nvarchar(50) NULL
    , [HomeAddressCountry] nvarchar(50) NULL
    , [HomeAddressPostOfficeBox] nvarchar(50) NULL
    , [HomeAddressPostalCode] nvarchar(50) NULL
    , [HomeAddressState] nvarchar(50) NULL
    , [HomeAddressStreet] nvarchar(255) NULL
    , [HomeFaxNumber] nvarchar(50) NULL
    , [HomeTelephoneNumber] nvarchar(50) NULL
    , [Home2TelephoneNumber] nvarchar(50) NULL
    , [IMAddress] nvarchar(50) NULL
    , [Initials] nvarchar(50) NULL
    , [InternetFreeBusyAddress] nvarchar(255) NULL
    , [ISDNNumber] nvarchar(50) NULL
    , [JobTitle] nvarchar(255) NULL
    , [Journal] bit NULL
    , [Language] nvarchar(50) NULL
    , [LastModificationTime] datetime NULL
    , [LastName] nvarchar(255) NULL
    , [Location] nvarchar(255) NULL
    , [MailingAddress] nvarchar(255) NULL
    , [ManagerName] nvarchar(255) NULL
    , [MessageClass] nvarchar(100) NULL
    , [MiddleName] nvarchar(255) NULL
    , [Mileage] nvarchar(50) NULL
    , [MobileTelephoneNumber] nvarchar(50) NULL
    , [NickName] nvarchar(50) NULL
    , [NoAging] bit NULL
    , [OfficeLocation] nvarchar(255) NULL
    , [OrganizationalIDNumber] nvarchar(50) NULL
    , [OtherAddress] nvarchar(255) NULL
    , [OtherAddressCity] nvarchar(50) NULL
    , [OtherAddressCountry] nvarchar(50) NULL
    , [OtherAddressPostOfficeBox] nvarchar(50) NULL
    , [OtherAddressPostalCode] nvarchar(50) NULL
    , [OtherAddressState] nvarchar(50) NULL
    , [OtherAddressStreet] nvarchar(255) NULL
    , [OtherFaxNumber] nvarchar(50) NULL
    , [OtherTelephoneNumber] nvarchar(50) NULL
    , [PagerNumber] nvarchar(50) NULL
    , [PersonalHomePage] nvarchar(255) NULL
    , [PrimaryTelephoneNumber] nvarchar(50) NULL
    , [Profession] nvarchar(50) NULL
    , [RadioTelephoneNumber] nvarchar(50) NULL
    , [ReferredBy] nvarchar(255) NULL
    , [ReminderSet] bit NULL
    , [ReminderTime] datetime NULL
    , [Sensitivity] tinyint NULL
    , [Size] int NULL
    , [Spouse] nvarchar(255) NULL
    , [Subject] nvarchar(255) NULL
    , [Suffix] nvarchar(50) NULL
    , [TelexNumber] nvarchar(50) NULL
    , [Title] nvarchar(255) NULL
    , [TTYTDDTelephoneNumber] nvarchar(50) NULL
    , [UnRead] bit NULL
    , [User1] nvarchar(255) NULL
    , [User2] nvarchar(255) NULL
    , [User3] nvarchar(255) NULL
    , [User4] nvarchar(255) NULL
    , [WebPage] nvarchar(255) NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_contacts] PRIMARY KEY ([id])
    , CONSTRAINT [IX_contacts_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[contacts] ADD CONSTRAINT [FK_contacts_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[journals] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [Attachments_Count] tinyint NULL
    , [BillingInformation] nvarchar(255) NULL
    , [Body] nvarchar(max) NULL
    , [Categories] nvarchar(255) NULL
    , [CompanyName] nvarchar(255) NULL
    , [CreationTime] datetime NULL
    , [Duration] int NULL
    , [End] datetime NULL
    , [FormDescription_ContactName] nvarchar(255) NULL
    , [LastModificationTime] datetime NULL
    , [MessageClass] nvarchar(100) NULL
    , [Mileage] nvarchar(50) NULL
    , [NoAging] bit NULL
    , [Type] nvarchar(50) NULL
    , [Sensitivity] tinyint NULL
    , [Size] int NULL
    , [Start] datetime NULL
    , [Subject] nvarchar(255) NULL
    , [UnRead] bit NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_journals] PRIMARY KEY ([id])
    , CONSTRAINT [IX_journals_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[journals] ADD CONSTRAINT [FK_journals_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[mails] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [Attachments_Count] tinyint NULL
    , [BCC] nvarchar(255) NULL
    , [BillingInformation] nvarchar(255) NULL
    , [Body] nvarchar(max) NULL
    , [Categories] nvarchar(255) NULL
    , [CC] nvarchar(255) NULL
    , [ConversationTopic] nvarchar(255) NULL
    , [CreationTime] datetime NULL
    , [DeferredDeliveryTime] datetime NULL
    , [ExpiryTime] datetime NULL
    , [FlagDueBy] datetime NULL
    , [FlagStatus] tinyint NULL
    , [FlagRequest] nvarchar(50) NULL
    , [Importance] tinyint NULL
    , [LastModificationTime] datetime NULL
    , [MessageClass] nvarchar(100) NULL
    , [Mileage] nvarchar(50) NULL
    , [NoAging] bit NULL
    , [Recipients_Item_1_Name] nvarchar(255) NULL
    , [Recipients_Item_1_Address] nvarchar(255) NULL
    , [Recipients_Item_1_Type] tinyint NULL
    , [Recipients_Item_1_EntryID] varchar(48) NULL
    , [Recipients_Item_2_Name] nvarchar(255) NULL
    , [Recipients_Item_2_Address] nvarchar(255) NULL
    , [Recipients_Item_2_EntryID] varchar(48) NULL
    , [Recipients_Item_2_Type] tinyint NULL
    , [ReplyRecipientNames] nvarchar(255) NULL
    , [ReadReceiptRequested] bit NULL
    , [ReceivedTime] datetime NULL
    , [RemoteStatus] tinyint NULL
    , [Sender_Address] nvarchar(255) NULL
    , [Sender_Name] nvarchar(255) NULL
    , [Sensitivity] tinyint NULL
    , [SentOn] datetime NULL
    , [SentOnBehalfOfName] nvarchar(255) NULL
    , [Size] int NULL
    , [Subject] nvarchar(255) NULL
    , [To] nvarchar(255) NULL
    , [TrackingStatus] tinyint NULL
    , [UnRead] bit NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_mails] PRIMARY KEY ([id])
    , CONSTRAINT [IX_mails_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[mails] ADD CONSTRAINT [FK_mails_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[notes] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [Categories] nvarchar(255) NULL
    , [Color] tinyint NULL
    , [Body] nvarchar(max) NULL
    , [CreationTime] datetime NULL
    , [NoAging] bit NULL
    , [MessageClass] nvarchar(100) NULL
    , [LastModificationTime] datetime NULL
    , [UnRead] bit NULL
    , [Size] int NULL
    , [Subject] nvarchar(255) NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_notes] PRIMARY KEY ([id])
    , CONSTRAINT [IX_notes_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[notes] ADD CONSTRAINT [FK_notes_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

CREATE TABLE [gcrm].[tasks] (
      [id] int IDENTITY(1,1) NOT NULL
    , [folder_id] int NOT NULL
    , [delete_in_outlook] bit NULL
    , [EntryID] char(48) NULL
    , [ActualWork] int NULL
    , [Attachments_Count] tinyint NULL
    , [BillingInformation] nvarchar(255) NULL
    , [Body] nvarchar(max) NULL
    , [Categories] nvarchar(255) NULL
    , [CompanyName] nvarchar(255) NULL
    , [Complete] bit NULL
    , [ConversationTopic] nvarchar(255) NULL
    , [CreationTime] datetime NULL
    , [DateCompleted] datetime NULL
    , [DelegationState] tinyint NULL
    , [DueDate] datetime NULL
    , [Importance] tinyint NULL
    , [IsRecurring] bit NULL
    , [LastModificationTime] datetime NULL
    , [MessageClass] nvarchar(100) NULL
    , [Mileage] nvarchar(50) NULL
    , [NoAging] bit NULL
    , [Owner] nvarchar(255) NULL
    , [PercentComplete] float NULL
    , [ReminderOverrideDefault] bit NULL
    , [ReminderSet] bit NULL
    , [ReminderPlaySound] bit NULL
    , [ReminderTime] datetime NULL
    , [Role] nvarchar(50) NULL
    , [SchedulePlusPriority] nvarchar(50) NULL
    , [Sensitivity] tinyint NULL
    , [Size] int NULL
    , [StartDate] datetime NULL
    , [Status] tinyint NULL
    , [Subject] nvarchar(255) NULL
    , [TeamTask] bit NULL
    , [To] nvarchar(255) NULL
    , [TotalWork] int NULL
    , [UnRead] bit NULL
    , [last_import_time] datetime NULL
    , [last_update_time] datetime NULL
    , CONSTRAINT [PK_tasks] PRIMARY KEY ([id])
    , CONSTRAINT [IX_tasks_folderid_entryid] UNIQUE ([folder_id], [EntryID])
);
GO

ALTER TABLE [gcrm].[tasks] ADD CONSTRAINT [FK_tasks_folders] FOREIGN KEY ([folder_id]) REFERENCES [gcrm].[folders] ([id]) ON DELETE CASCADE ON UPDATE CASCADE;
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: SaveToDB Query List
-- =============================================

CREATE VIEW [gcrm].[queries]
AS

SELECT
    s.name AS TABLE_SCHEMA
    , o.name AS TABLE_NAME
    , CASE o.[type] WHEN 'U' THEN 'BASE TABLE' WHEN 'V' THEN 'VIEW' WHEN 'P' THEN 'PROCEDURE' WHEN 'IF' THEN 'FUNCTION' ELSE o.type_desc COLLATE Latin1_General_CI_AS END AS TABLE_TYPE
    , CAST(NULL AS nvarchar(max)) AS TABLE_CODE
    , CAST(NULL AS nvarchar(max)) AS INSERT_PROCEDURE
    , CAST(NULL AS nvarchar(max)) AS UPDATE_PROCEDURE
    , CAST(NULL AS nvarchar(max)) AS DELETE_PROCEDURE
    , CAST(NULL AS nvarchar(50)) AS PROCEDURE_TYPE
FROM
    sys.objects o
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
    o.[type] IN ('U', 'V', 'P', 'IF')
    AND (HAS_PERMS_BY_NAME(DB_NAME() + '.' + s.name + '.' + o.name, 'OBJECT', 'SELECT') = 1
        OR HAS_PERMS_BY_NAME(DB_NAME() + '.' + s.name + '.' + o.name, 'OBJECT', 'EXECUTE') = 1)
    AND NOT s.name IN ('sys', 'gcrm', 'etl01', 'dbo01', 'gcrm01', 'savetodb_dev', 'savetodb_gcrm', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL')
    AND NOT (s.name = 'dbo' AND (
           o.name LIKE 'sp_%'
        OR o.name LIKE 'fn_%'
        OR o.name LIKE 'sys%'
        ))
    AND NOT o.name LIKE 'xl_%'
    AND NOT (o.[type] = 'V' AND (
           o.name LIKE 'viewQueryList%'
        OR o.name LIKE 'viewParameterValues%'
        OR o.name LIKE 'viewValidationList%'
        OR o.name LIKE 'view_query_list%'
        OR o.name LIKE 'view_xl_%'
        OR o.name LIKE 'xl_%'
        ))
    AND NOT (o.[type] = 'P' AND (
           o.name LIKE '%_insert'
        OR o.name LIKE '%_update'
        OR o.name LIKE '%_delete'
        OR o.name LIKE '%_merge'
        OR o.name LIKE '%_change'
        OR o.name LIKE 'uspExcelEvent%'
        OR o.name LIKE 'uspParameterValues%'
        OR o.name LIKE 'uspValidationList%'
        OR o.name LIKE 'uspAdd%'
        OR o.name LIKE 'uspSet%'
        OR o.name LIKE 'uspInsert%'
        OR o.name LIKE 'uspUpdate%'
        OR o.name LIKE 'uspDelete%'
        OR o.name LIKE 'usp_insert_%'
        OR o.name LIKE 'usp_update_%'
        OR o.name LIKE 'usp_delete_%'
        OR o.name LIKE 'Add%'
        OR o.name LIKE 'Set%'
        OR o.name LIKE 'Insert%'
        OR o.name LIKE 'Update%'
        OR o.name LIKE 'Delete%'
        OR o.name LIKE 'usp_xl_%'
        OR o.name LIKE 'usp_import_%'
        OR o.name LIKE 'usp_export_%'
        OR o.name LIKE 'usp_sync_%'
        ))
    AND NOT (o.[type] = 'IF' AND (
           o.name LIKE 'Has%'
        OR o.name LIKE 'ufnExcelEvent%'
        OR o.name LIKE 'ufnParameterValues%'
        OR o.name LIKE 'ufn_xl_%'
        ))
UNION ALL
SELECT
    o.TABLE_SCHEMA
    , o.TABLE_NAME
    , o.TABLE_TYPE COLLATE Latin1_General_CI_AS
    , o.TABLE_CODE
    , o.INSERT_OBJECT AS INSERT_PROCEDURE
    , o.UPDATE_OBJECT AS UPDATE_PROCEDURE
    , o.DELETE_OBJECT AS DELETE_PROCEDURE
    , CAST(NULL AS nvarchar(50)) AS PROCEDURE_TYPE
FROM
    gcrm.[objects] o
WHERE
    o.TABLE_TYPE IN ('CODE', 'HTTP', 'TEXT')
    AND o.TABLE_SCHEMA IS NOT NULL
    AND o.TABLE_NAME IS NOT NULL
    AND o.TABLE_CODE IS NOT NULL
    AND NOT o.TABLE_NAME LIKE 'xl_%'


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the appointments table
-- =============================================

CREATE VIEW [gcrm].[view_appointments]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Importance] AS [Importance]
    , t.[Subject] AS [Subject]

    , t.[AllDayEvent] AS [AllDayEvent]
    , t.[Start] AS [Start]
    , t.[End] AS [End]
    , t.[Duration] AS [Duration]
    , t.[Location] AS [Location]
    , t.[BusyStatus] AS [BusyStatus]
    , t.[Organizer] AS [Organizer]
    , t.[RequiredAttendees] AS [RequiredAttendees]
    , t.[OptionalAttendees] AS [OptionalAttendees]
    , t.[Body] AS [Body]

    , t.[BillingInformation] AS [BillingInformation]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[MeetingStatus] AS [MeetingStatus]
    , t.[Mileage] AS [Mileage]
    , t.[RecurrencePattern_RecurrenceType] AS [RecurrencePattern_RecurrenceType]
    , t.[RecurrencePattern_PatternEndDate] AS [RecurrencePattern_PatternEndDate]
    , t.[RecurrencePattern_PatternStartDate] AS [RecurrencePattern_PatternStartDate]
    , t.[IsRecurring] AS [IsRecurring]
    , t.[ReminderMinutesBeforeStart] AS [ReminderMinutesBeforeStart]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[ReminderOverrideDefault] AS [ReminderOverrideDefault]
    , t.[ReminderPlaySound] AS [ReminderPlaySound]
    , t.[Resources] AS [Resources]
    , t.[ResponseRequested] AS [ResponseRequested]

    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[Size] AS [Size]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.appointments t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the contacts table
-- =============================================

CREATE VIEW [gcrm].[view_contacts]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Subject] AS [Subject]

    , t.[CompanyName] AS [CompanyName]
    , t.[JobTitle] AS [JobTitle]
    , t.[Title] AS [Title]
    , t.[FirstName] AS [FirstName]
    , t.[MiddleName] AS [MiddleName]
    , t.[LastName] AS [LastName]
    , t.[FullName] AS [FullName]
    , t.[FileAs] AS [FileAs]
    , t.[Email1Address] AS [Email1Address]
    , t.[Email2Address] AS [Email2Address]
    , t.[Email3Address] AS [Email3Address]
    , t.[PrimaryTelephoneNumber] AS [PrimaryTelephoneNumber]
    , t.[MobileTelephoneNumber] AS [MobileTelephoneNumber]

    , t.[Account] AS [Account]
    , t.[Anniversary] AS [Anniversary]
    , t.[AssistantName] AS [AssistantName]
    , t.[AssistantTelephoneNumber] AS [AssistantTelephoneNumber]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Birthday] AS [Birthday]
    , t.[BusinessAddress] AS [BusinessAddress]
    , t.[BusinessAddressCity] AS [BusinessAddressCity]
    , t.[BusinessAddressCountry] AS [BusinessAddressCountry]
    , t.[BusinessAddressPostOfficeBox] AS [BusinessAddressPostOfficeBox]
    , t.[BusinessAddressPostalCode] AS [BusinessAddressPostalCode]
    , t.[BusinessAddressState] AS [BusinessAddressState]
    , t.[BusinessAddressStreet] AS [BusinessAddressStreet]
    , t.[BusinessFaxNumber] AS [BusinessFaxNumber]
    , t.[BusinessHomePage] AS [BusinessHomePage]
    , t.[BusinessTelephoneNumber] AS [BusinessTelephoneNumber]
    , t.[Business2TelephoneNumber] AS [Business2TelephoneNumber]
    , t.[CallbackTelephoneNumber] AS [CallbackTelephoneNumber]
    , t.[CarTelephoneNumber] AS [CarTelephoneNumber]
    , t.[Children] AS [Children]
    , t.[CompanyMainTelephoneNumber] AS [CompanyMainTelephoneNumber]
    , t.[ComputerNetworkName] AS [ComputerNetworkName]
    , t.[CustomerID] AS [CustomerID]
    , t.[Department] AS [Department]
    , t.[FlagRequest] AS [FlagRequest]
    , t.[FTPSite] AS [FTPSite]
    , t.[Gender] AS [Gender]
    , t.[GovernmentIDNumber] AS [GovernmentIDNumber]
    , t.[Hobby] AS [Hobby]
    , t.[HomeAddress] AS [HomeAddress]
    , t.[HomeAddressCity] AS [HomeAddressCity]
    , t.[HomeAddressCountry] AS [HomeAddressCountry]
    , t.[HomeAddressPostOfficeBox] AS [HomeAddressPostOfficeBox]
    , t.[HomeAddressPostalCode] AS [HomeAddressPostalCode]
    , t.[HomeAddressState] AS [HomeAddressState]
    , t.[HomeAddressStreet] AS [HomeAddressStreet]
    , t.[HomeFaxNumber] AS [HomeFaxNumber]
    , t.[HomeTelephoneNumber] AS [HomeTelephoneNumber]
    , t.[Home2TelephoneNumber] AS [Home2TelephoneNumber]
    , t.[IMAddress] AS [IMAddress]
    , t.[Initials] AS [Initials]
    , t.[InternetFreeBusyAddress] AS [InternetFreeBusyAddress]
    , t.[ISDNNumber] AS [ISDNNumber]
    , t.[Journal] AS [Journal]
    , t.[Language] AS [Language]
    , t.[Location] AS [Location]
    , t.[MailingAddress] AS [MailingAddress]
    , t.[ManagerName] AS [ManagerName]
    , t.[Mileage] AS [Mileage]
    , t.[NickName] AS [NickName]
    , t.[Body] AS [Body]
    , t.[OfficeLocation] AS [OfficeLocation]
    , t.[OrganizationalIDNumber] AS [OrganizationalIDNumber]
    , t.[OtherAddress] AS [OtherAddress]
    , t.[OtherAddressCity] AS [OtherAddressCity]
    , t.[OtherAddressCountry] AS [OtherAddressCountry]
    , t.[OtherAddressPostOfficeBox] AS [OtherAddressPostOfficeBox]
    , t.[OtherAddressPostalCode] AS [OtherAddressPostalCode]
    , t.[OtherAddressState] AS [OtherAddressState]
    , t.[OtherAddressStreet] AS [OtherAddressStreet]
    , t.[OtherFaxNumber] AS [OtherFaxNumber]
    , t.[OtherTelephoneNumber] AS [OtherTelephoneNumber]
    , t.[PagerNumber] AS [PagerNumber]
    , t.[PersonalHomePage] AS [PersonalHomePage]
    , t.[Profession] AS [Profession]
    , t.[RadioTelephoneNumber] AS [RadioTelephoneNumber]
    , t.[ReferredBy] AS [ReferredBy]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[ReminderTime] AS [ReminderTime]
    , t.[Spouse] AS [Spouse]
    , t.[Suffix] AS [Suffix]
    , t.[TelexNumber] AS [TelexNumber]
    , t.[TTYTDDTelephoneNumber] AS [TTYTDDTelephoneNumber]
    , t.[User1] AS [User1]
    , t.[User2] AS [User2]
    , t.[User3] AS [User3]
    , t.[User4] AS [User4]
    , t.[WebPage] AS [WebPage]

    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[Size] AS [Size]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.contacts t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the juornals table
-- =============================================

CREATE VIEW [gcrm].[view_journals]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Subject] AS [Subject]

    , t.[Type] AS [Type]
    , t.[Start] AS [Start]
    , t.[End] AS [End]
    , t.[Duration] AS [Duration]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[CompanyName] AS [CompanyName]
    , t.[FormDescription_ContactName] AS [FormDescription_ContactName]
    , t.[Mileage] AS [Mileage]
    , t.[Body] AS [Body]

    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[Size] AS [Size]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.journals t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the mails table
-- =============================================

CREATE VIEW [gcrm].[view_mails]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Importance] AS [Importance]
    , t.[Subject] AS [Subject]

    , t.[SentOn] AS [SentOn]
    , t.[ReceivedTime] AS [ReceivedTime]
    , t.[Sender_Name]
    , t.[Sender_Address]
    , t.[Recipients_Item_1_Name]
    , t.[Recipients_Item_1_Address]
    , t.[Recipients_Item_2_Name]
    , t.[Recipients_Item_2_Address]
    , t.[To] AS [To]
    , t.[CC] AS [CC]
    , t.[BCC] AS [BCC]
    , t.[Body] AS [Body]

    , t.[BillingInformation] AS [BillingInformation]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[DeferredDeliveryTime] AS [DeferredDeliveryTime]
    , t.[FlagDueBy] AS [FlagDueBy]
    , t.[ExpiryTime] AS [ExpiryTime]
    , t.[FlagStatus] AS [FlagStatus]
    , t.[FlagRequest] AS [FlagRequest]
    , t.[SentOnBehalfOfName] AS [SentOnBehalfOfName]
    , t.[ReplyRecipientNames] AS [ReplyRecipientNames]
    , t.[Mileage] AS [Mileage]
    , t.[ReadReceiptRequested] AS [ReadReceiptRequested]
    , t.[RemoteStatus] AS [RemoteStatus]
    , t.[TrackingStatus] AS [TrackingStatus]

    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[Size] AS [Size]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.mails t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the notes table
-- =============================================

CREATE VIEW [gcrm].[view_notes]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Subject] AS [Subject]

    , t.[Body] AS [Body]
    , t.[Color] AS [Color]

    , t.[Size] AS [Size]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.notes t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view selects columns of Outlook data tables
-- =============================================

CREATE VIEW [gcrm].[view_outlook_columns]
AS

SELECT
    c.COLUMN_NAME
    , c2.DATA_TYPE + CASE WHEN c2.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c2.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c2.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END AS [contacts]
    , c4.DATA_TYPE + CASE WHEN c4.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c4.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c4.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS [mails]
    , c1.DATA_TYPE + CASE WHEN c1.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c1.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c1.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS [appointments]
    , c6.DATA_TYPE + CASE WHEN c6.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c6.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c6.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS [tasks]
    , c3.DATA_TYPE + CASE WHEN c3.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c3.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c3.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS [journals]
    , c5.DATA_TYPE + CASE WHEN c5.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c5.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c5.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS [notes]
FROM
    (
    SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'appointments'
    UNION SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'contacts'
    UNION SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'journals'
    UNION SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'mails'
    UNION SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'notes'
    UNION SELECT c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA = 'gcrm' AND c.TABLE_NAME = 'tasks'
    ) c
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c1 ON c1.TABLE_SCHEMA = 'gcrm' AND c1.TABLE_NAME = 'appointments' AND c1.COLUMN_NAME = c.COLUMN_NAME
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c2 ON c2.TABLE_SCHEMA = 'gcrm' AND c2.TABLE_NAME = 'contacts' AND c2.COLUMN_NAME = c.COLUMN_NAME
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c3 ON c3.TABLE_SCHEMA = 'gcrm' AND c3.TABLE_NAME = 'journals' AND c3.COLUMN_NAME = c.COLUMN_NAME
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c4 ON c4.TABLE_SCHEMA = 'gcrm' AND c4.TABLE_NAME = 'mails' AND c4.COLUMN_NAME = c.COLUMN_NAME
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c5 ON c5.TABLE_SCHEMA = 'gcrm' AND c5.TABLE_NAME = 'notes' AND c5.COLUMN_NAME = c.COLUMN_NAME
    LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS c6 ON c6.TABLE_SCHEMA = 'gcrm' AND c6.TABLE_NAME = 'tasks' AND c6.COLUMN_NAME = c.COLUMN_NAME


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view selects columns of Outlook data tables and update procedure parameterss
-- =============================================

CREATE VIEW [gcrm].[view_outlook_parameters]
AS

SELECT
    c.TABLE_SCHEMA
    , c.TABLE_NAME
    , c.COLUMN_NAME
    , c.DATA_TYPE + CASE WHEN c.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN c.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(c.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS DATA_TYPE
    , p1.DATA_TYPE + CASE WHEN p1.CHARACTER_MAXIMUM_LENGTH = -1 THEN '(max)' WHEN p1.CHARACTER_MAXIMUM_LENGTH IS NOT NULL THEN '(' + CAST(p1.CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')' ELSE '' END  AS 'UPDATE_DATA_TYPE'
FROM
    INFORMATION_SCHEMA.COLUMNS c
    LEFT OUTER JOIN INFORMATION_SCHEMA.PARAMETERS p1 ON p1.SPECIFIC_SCHEMA = c.TABLE_SCHEMA AND p1.SPECIFIC_NAME = 'usp_sync_' + c.TABLE_NAME + '_update' AND p1.PARAMETER_NAME = '@' + c.COLUMN_NAME
WHERE
    c.TABLE_NAME IN ('appointments', 'contacts', 'journals', 'mails', 'notes', 'tasks')


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: A view of the tasks table
-- =============================================

CREATE VIEW [gcrm].[view_tasks]
AS

SELECT
      t.[id]
    , t.[folder_id]
    , t.[delete_in_outlook]
    , t.[MessageClass] AS [MessageClass]
    , t.[Categories] AS [Categories]
    , t.[Importance] AS [Importance]
    , t.[Subject] AS [Subject]

    , t.[StartDate] AS [StartDate]
    , t.[PercentComplete] AS [PercentComplete]
    , t.[ActualWork] AS [ActualWork]
    , t.[DelegationState] AS [DelegationState]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[CompanyName] AS [CompanyName]
    , t.[Complete] AS [Complete]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[DateCompleted] AS [DateCompleted]
    , t.[DueDate] AS [DueDate]
    , t.[Mileage] AS [Mileage]
    , t.[Body] AS [Body]
    , t.[Owner] AS [Owner]
    , t.[IsRecurring] AS [IsRecurring]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[ReminderOverrideDefault] AS [ReminderOverrideDefault]
    , t.[ReminderPlaySound] AS [ReminderPlaySound]
    , t.[ReminderTime] AS [ReminderTime]
    , t.[Role] AS [Role]
    , t.[SchedulePlusPriority] AS [SchedulePlusPriority]
    , t.[Status] AS [Status]
    , t.[TeamTask] AS [TeamTask]
    , t.[To] AS [To]
    , t.[TotalWork] AS [TotalWork]

    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[Size] AS [Size]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[NoAging] AS [NoAging]
    , t.[UnRead] AS [UnRead]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[last_import_time]
    , t.[last_update_time]
FROM
    gcrm.tasks t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Returns @folder_id and updates the users and folders tables if required
-- =============================================

CREATE PROCEDURE [gcrm].[usp_get_folder_id]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @folder_type_id tinyint = NULL
    , @folder_id int OUTPUT
AS
BEGIN

SET NOCOUNT ON

DECLARE @name nvarchar(255)
DECLARE @path nvarchar(255)
DECLARE @entry_id char(48)
DECLARE @user_id smallint
DECLARE @user_name nvarchar(128)

SELECT @folder_id = f.id, @entry_id = f.EntryID, @name = f.Name, @path = f.FolderPath FROM gcrm.folders f WHERE f.[user_id] = USER_ID() AND f.StoreID = @StoreID AND f.EntryID = @FolderEntryID

-- The folder is not specified, return the @folder_id as is

IF @FolderName IS NULL OR @FolderPath IS NULL
    RETURN

-- The folder exists with an unchanged name and folder path

IF @folder_id IS NOT NULL AND @name = @FolderName AND @path = @FolderPath
    RETURN

-- The folder exists but it has a new name or folder path

IF @folder_id IS NOT NULL
    BEGIN
    UPDATE gcrm.folders SET Name = @FolderName, FolderPath = @FolderPath WHERE id = @folder_id
    RETURN
    END

-- The folder_type_id is required to insert a new folder

IF @folder_type_id IS NULL
    RETURN

SELECT @user_id = u.id, @user_name = u.name FROM gcrm.users u WHERE u.id = USER_ID() OR u.name = USER_NAME()

IF @user_id IS NULL
    BEGIN
    -- New user
    INSERT INTO gcrm.users (id, name) VALUES (USER_ID(), USER_NAME())
    END
ELSE IF NOT @user_id = USER_ID()
    BEGIN
    -- Dropped and created user

    -- This update updates the folder table also via the cascade foreagn key
    UPDATE gcrm.users SET id = USER_ID() WHERE name = USER_NAME()

    -- The recursive call is OK as the user exists now with a correct user_id
    EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, @folder_type_id, @folder_id
    RETURN
    END
ELSE IF NOT @user_name = USER_NAME()
    BEGIN
    -- Renamed user
    UPDATE gcrm.users SET name = USER_NAME() WHERE id = @user_id
    END

INSERT INTO gcrm.folders ([user_id], folder_type_id, StoreID, EntryID, Name, FolderPath) VALUES (USER_ID(), @folder_type_id, @StoreID, @FolderEntryID, @FolderName, @FolderPath)

SET @folder_id = SCOPE_IDENTITY()

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the appointments table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_appointments_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[AllDayEvent] AS [AllDayEvent]
    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Body] AS [Body]
    , t.[BusyStatus] AS [BusyStatus]
    , t.[Categories] AS [Categories]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[CreationTime] AS [CreationTime]
    , t.[Duration] AS [Duration]
    , t.[End] AS [End]
    , t.[Importance] AS [Importance]
    , t.[IsRecurring] AS [IsRecurring]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[Location] AS [Location]
    , t.[MeetingStatus] AS [MeetingStatus]
    , t.[MessageClass] AS [MessageClass]
    , t.[Mileage] AS [Mileage]
    , t.[NoAging] AS [NoAging]
    , t.[OptionalAttendees] AS [OptionalAttendees]
    , t.[Organizer] AS [Organizer]
    , t.[RecurrencePattern_PatternEndDate] AS [RecurrencePattern_PatternEndDate]
    , t.[RecurrencePattern_PatternStartDate] AS [RecurrencePattern_PatternStartDate]
    , t.[RecurrencePattern_RecurrenceType] AS [RecurrencePattern_RecurrenceType]
    , t.[ReminderMinutesBeforeStart] AS [ReminderMinutesBeforeStart]
    , t.[ReminderOverrideDefault] AS [ReminderOverrideDefault]
    , t.[ReminderPlaySound] AS [ReminderPlaySound]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[RequiredAttendees] AS [RequiredAttendees]
    , t.[Resources] AS [Resources]
    , t.[ResponseRequested] AS [ResponseRequested]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[Size] AS [Size]
    , t.[Start] AS [Start]
    , t.[Subject] AS [Subject]
    , t.[UnRead] AS [UnRead]
FROM
    gcrm.appointments t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the appointments table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_appointments_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @AllDayEvent bit = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @BusyStatus tinyint = NULL
    , @Categories nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @Location nvarchar(255) = NULL
    , @MeetingStatus tinyint = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage varchar(50) = NULL
    , @NoAging bit = NULL
    , @OptionalAttendees nvarchar(255) = NULL
    , @Organizer nvarchar(255) = NULL
    , @RecurrencePattern_PatternEndDate datetime = NULL
    , @RecurrencePattern_PatternStartDate datetime = NULL
    , @RecurrencePattern_RecurrenceType tinyint = NULL
    , @ReminderMinutesBeforeStart int = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @RequiredAttendees nvarchar(255) = NULL
    , @Resources nvarchar(255) = NULL
    , @ResponseRequested bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 3, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.appointments WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.appointments
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [AllDayEvent] = @AllDayEvent
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [BusyStatus] = @BusyStatus
    , [Categories] = @Categories
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [Duration] = @Duration
    , [End] = @End
    , [Importance] = @Importance
    , [IsRecurring] = @IsRecurring
    , [LastModificationTime] = @LastModificationTime
    , [Location] = @Location
    , [MeetingStatus] = @MeetingStatus
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [OptionalAttendees] = @OptionalAttendees
    , [Organizer] = @Organizer
    , [RecurrencePattern_PatternEndDate] = @RecurrencePattern_PatternEndDate
    , [RecurrencePattern_PatternStartDate] = @RecurrencePattern_PatternStartDate
    , [RecurrencePattern_RecurrenceType] = @RecurrencePattern_RecurrenceType
    , [ReminderMinutesBeforeStart] = @ReminderMinutesBeforeStart
    , [ReminderOverrideDefault] = @ReminderOverrideDefault
    , [ReminderPlaySound] = @ReminderPlaySound
    , [ReminderSet] = @ReminderSet
    , [RequiredAttendees] = @RequiredAttendees
    , [Resources] = @Resources
    , [ResponseRequested] = @ResponseRequested
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Start] = @Start
    , [Subject] = @Subject
    , [UnRead] = @UnRead
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.appointments
    ( [folder_id]
    , [EntryID]
    , [AllDayEvent]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [BusyStatus]
    , [Categories]
    , [ConversationTopic]
    , [CreationTime]
    , [Duration]
    , [End]
    , [Importance]
    , [IsRecurring]
    , [LastModificationTime]
    , [Location]
    , [MeetingStatus]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [OptionalAttendees]
    , [Organizer]
    , [RecurrencePattern_PatternEndDate]
    , [RecurrencePattern_PatternStartDate]
    , [RecurrencePattern_RecurrenceType]
    , [ReminderMinutesBeforeStart]
    , [ReminderOverrideDefault]
    , [ReminderPlaySound]
    , [ReminderSet]
    , [RequiredAttendees]
    , [Resources]
    , [ResponseRequested]
    , [Sensitivity]
    , [Size]
    , [Start]
    , [Subject]
    , [UnRead]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @AllDayEvent
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @BusyStatus
    , @Categories
    , @ConversationTopic
    , @CreationTime
    , @Duration
    , @End
    , @Importance
    , @IsRecurring
    , @LastModificationTime
    , @Location
    , @MeetingStatus
    , @MessageClass
    , @Mileage
    , @NoAging
    , @OptionalAttendees
    , @Organizer
    , @RecurrencePattern_PatternEndDate
    , @RecurrencePattern_PatternStartDate
    , @RecurrencePattern_RecurrenceType
    , @ReminderMinutesBeforeStart
    , @ReminderOverrideDefault
    , @ReminderPlaySound
    , @ReminderSet
    , @RequiredAttendees
    , @Resources
    , @ResponseRequested
    , @Sensitivity
    , @Size
    , @Start
    , @Subject
    , @UnRead
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )


END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the contacts table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_contacts_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[Account] AS [Account]
    , t.[Anniversary] AS [Anniversary]
    , t.[AssistantName] AS [AssistantName]
    , t.[AssistantTelephoneNumber] AS [AssistantTelephoneNumber]
    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Birthday] AS [Birthday]
    , t.[Body] AS [Body]
    , t.[Business2TelephoneNumber] AS [Business2TelephoneNumber]
    , t.[BusinessAddress] AS [BusinessAddress]
    , t.[BusinessAddressCity] AS [BusinessAddressCity]
    , t.[BusinessAddressCountry] AS [BusinessAddressCountry]
    , t.[BusinessAddressPostalCode] AS [BusinessAddressPostalCode]
    , t.[BusinessAddressPostOfficeBox] AS [BusinessAddressPostOfficeBox]
    , t.[BusinessAddressState] AS [BusinessAddressState]
    , t.[BusinessAddressStreet] AS [BusinessAddressStreet]
    , t.[BusinessFaxNumber] AS [BusinessFaxNumber]
    , t.[BusinessHomePage] AS [BusinessHomePage]
    , t.[BusinessTelephoneNumber] AS [BusinessTelephoneNumber]
    , t.[CallbackTelephoneNumber] AS [CallbackTelephoneNumber]
    , t.[CarTelephoneNumber] AS [CarTelephoneNumber]
    , t.[Categories] AS [Categories]
    , t.[Children] AS [Children]
    , t.[CompanyMainTelephoneNumber] AS [CompanyMainTelephoneNumber]
    , t.[CompanyName] AS [CompanyName]
    , t.[ComputerNetworkName] AS [ComputerNetworkName]
    , t.[CreationTime] AS [CreationTime]
    , t.[CustomerID] AS [CustomerID]
    , t.[Department] AS [Department]
    , t.[Email1Address] AS [Email1Address]
    , t.[Email2Address] AS [Email2Address]
    , t.[Email3Address] AS [Email3Address]
    , t.[FileAs] AS [FileAs]
    , t.[FirstName] AS [FirstName]
    , t.[FlagRequest] AS [FlagRequest]
    , t.[FTPSite] AS [FTPSite]
    , t.[FullName] AS [FullName]
    , t.[Gender] AS [Gender]
    , t.[GovernmentIDNumber] AS [GovernmentIDNumber]
    , t.[Hobby] AS [Hobby]
    , t.[Home2TelephoneNumber] AS [Home2TelephoneNumber]
    , t.[HomeAddress] AS [HomeAddress]
    , t.[HomeAddressCity] AS [HomeAddressCity]
    , t.[HomeAddressCountry] AS [HomeAddressCountry]
    , t.[HomeAddressPostalCode] AS [HomeAddressPostalCode]
    , t.[HomeAddressPostOfficeBox] AS [HomeAddressPostOfficeBox]
    , t.[HomeAddressState] AS [HomeAddressState]
    , t.[HomeAddressStreet] AS [HomeAddressStreet]
    , t.[HomeFaxNumber] AS [HomeFaxNumber]
    , t.[HomeTelephoneNumber] AS [HomeTelephoneNumber]
    , t.[IMAddress] AS [IMAddress]
    , t.[Initials] AS [Initials]
    , t.[InternetFreeBusyAddress] AS [InternetFreeBusyAddress]
    , t.[ISDNNumber] AS [ISDNNumber]
    , t.[JobTitle] AS [JobTitle]
    , t.[Journal] AS [Journal]
    , t.[Language] AS [Language]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[LastName] AS [LastName]
    , t.[Location] AS [Location]
    , t.[MailingAddress] AS [MailingAddress]
    , t.[ManagerName] AS [ManagerName]
    , t.[MessageClass] AS [MessageClass]
    , t.[MiddleName] AS [MiddleName]
    , t.[Mileage] AS [Mileage]
    , t.[MobileTelephoneNumber] AS [MobileTelephoneNumber]
    , t.[NickName] AS [NickName]
    , t.[OfficeLocation] AS [OfficeLocation]
    , t.[OrganizationalIDNumber] AS [OrganizationalIDNumber]
    , t.[OtherAddress] AS [OtherAddress]
    , t.[OtherAddressCity] AS [OtherAddressCity]
    , t.[OtherAddressCountry] AS [OtherAddressCountry]
    , t.[OtherAddressPostalCode] AS [OtherAddressPostalCode]
    , t.[OtherAddressPostOfficeBox] AS [OtherAddressPostOfficeBox]
    , t.[OtherAddressState] AS [OtherAddressState]
    , t.[OtherAddressStreet] AS [OtherAddressStreet]
    , t.[OtherFaxNumber] AS [OtherFaxNumber]
    , t.[OtherTelephoneNumber] AS [OtherTelephoneNumber]
    , t.[PagerNumber] AS [PagerNumber]
    , t.[PersonalHomePage] AS [PersonalHomePage]
    , t.[PrimaryTelephoneNumber] AS [PrimaryTelephoneNumber]
    , t.[Profession] AS [Profession]
    , t.[RadioTelephoneNumber] AS [RadioTelephoneNumber]
    , t.[ReferredBy] AS [ReferredBy]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[ReminderTime] AS [ReminderTime]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[Size] AS [Size]
    , t.[Spouse] AS [Spouse]
    , t.[Subject] AS [Subject]
    , t.[Suffix] AS [Suffix]
    , t.[TelexNumber] AS [TelexNumber]
    , t.[Title] AS [Title]
    , t.[TTYTDDTelephoneNumber] AS [TTYTDDTelephoneNumber]
    , t.[UnRead] AS [UnRead]
    , t.[User1] AS [User1]
    , t.[User2] AS [User2]
    , t.[User3] AS [User3]
    , t.[User4] AS [User4]
    , t.[WebPage] AS [WebPage]
FROM
    gcrm.contacts t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the contacts table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_contacts_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @Account nvarchar(255) = NULL
    , @Anniversary datetime = NULL
    , @AssistantName nvarchar(255) = NULL
    , @AssistantTelephoneNumber nvarchar(50) = NULL
    , @Attachments_Count tinyint = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Birthday datetime = NULL
    , @Body nvarchar(max) = NULL
    , @Business2TelephoneNumber nvarchar(50) = NULL
    , @BusinessAddress nvarchar(255) = NULL
    , @BusinessAddressCity nvarchar(50) = NULL
    , @BusinessAddressCountry nvarchar(50) = NULL
    , @BusinessAddressPostalCode nvarchar(50) = NULL
    , @BusinessAddressPostOfficeBox nvarchar(50) = NULL
    , @BusinessAddressState nvarchar(50) = NULL
    , @BusinessAddressStreet nvarchar(255) = NULL
    , @BusinessFaxNumber nvarchar(50) = NULL
    , @BusinessHomePage nvarchar(50) = NULL
    , @BusinessTelephoneNumber nvarchar(50) = NULL
    , @CallbackTelephoneNumber nvarchar(50) = NULL
    , @CarTelephoneNumber nvarchar(50) = NULL
    , @Categories nvarchar(255) = NULL
    , @Children nvarchar(255) = NULL
    , @CompanyMainTelephoneNumber nvarchar(50) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @ComputerNetworkName nvarchar(50) = NULL
    , @CreationTime datetime = NULL
    , @CustomerID varchar(50) = NULL
    , @Department nvarchar(255) = NULL
    , @Email1Address varchar(100) = NULL
    , @Email2Address varchar(100) = NULL
    , @Email3Address varchar(100) = NULL
    , @FileAs nvarchar(255) = NULL
    , @FirstName nvarchar(50) = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FTPSite nvarchar(100) = NULL
    , @FullName nvarchar(255) = NULL
    , @Gender tinyint = NULL
    , @GovernmentIDNumber varchar(50) = NULL
    , @Hobby nvarchar(255) = NULL
    , @Home2TelephoneNumber nvarchar(50) = NULL
    , @HomeAddress nvarchar(255) = NULL
    , @HomeAddressCity nvarchar(50) = NULL
    , @HomeAddressCountry nvarchar(50) = NULL
    , @HomeAddressPostalCode nvarchar(50) = NULL
    , @HomeAddressPostOfficeBox nvarchar(50) = NULL
    , @HomeAddressState nvarchar(50) = NULL
    , @HomeAddressStreet nvarchar(255) = NULL
    , @HomeFaxNumber nvarchar(50) = NULL
    , @HomeTelephoneNumber nvarchar(50) = NULL
    , @IMAddress nvarchar(50) = NULL
    , @Initials nvarchar(50) = NULL
    , @InternetFreeBusyAddress varchar(255) = NULL
    , @ISDNNumber varchar(50) = NULL
    , @JobTitle nvarchar(255) = NULL
    , @Journal bit = NULL
    , @Language nvarchar(50) = NULL
    , @LastModificationTime datetime = NULL
    , @LastName nvarchar(255) = NULL
    , @Location nvarchar(255) = NULL
    , @MailingAddress nvarchar(255) = NULL
    , @ManagerName nvarchar(255) = NULL
    , @MessageClass nvarchar(100) = NULL
    , @MiddleName nvarchar(255) = NULL
    , @Mileage nvarchar(50) = NULL
    , @MobileTelephoneNumber nvarchar(50) = NULL
    , @NickName nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @OfficeLocation nvarchar(255) = NULL
    , @OrganizationalIDNumber varchar(50) = NULL
    , @OtherAddress nvarchar(255) = NULL
    , @OtherAddressCity nvarchar(50) = NULL
    , @OtherAddressCountry nvarchar(50) = NULL
    , @OtherAddressPostalCode nvarchar(50) = NULL
    , @OtherAddressPostOfficeBox nvarchar(50) = NULL
    , @OtherAddressState nvarchar(50) = NULL
    , @OtherAddressStreet nvarchar(255) = NULL
    , @OtherFaxNumber nvarchar(50) = NULL
    , @OtherTelephoneNumber nvarchar(50) = NULL
    , @PagerNumber nvarchar(50) = NULL
    , @PersonalHomePage nvarchar(255) = NULL
    , @PrimaryTelephoneNumber nvarchar(50) = NULL
    , @Profession nvarchar(50) = NULL
    , @RadioTelephoneNumber nvarchar(50) = NULL
    , @ReferredBy nvarchar(255) = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Spouse nvarchar(255) = NULL
    , @Subject nvarchar(255) = NULL
    , @Suffix nvarchar(50) = NULL
    , @TelexNumber nvarchar(50) = NULL
    , @Title nvarchar(255) = NULL
    , @TTYTDDTelephoneNumber nvarchar(50) = NULL
    , @UnRead bit = NULL
    , @User1 nvarchar(255) = NULL
    , @User2 nvarchar(255) = NULL
    , @User3 nvarchar(255) = NULL
    , @User4 nvarchar(255) = NULL
    , @WebPage nvarchar(255) = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 1, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.contacts WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.contacts
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [Account] = @Account
    , [Anniversary] = @Anniversary
    , [AssistantName] = @AssistantName
    , [AssistantTelephoneNumber] = @AssistantTelephoneNumber
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Birthday] = @Birthday
    , [Body] = @Body
    , [Business2TelephoneNumber] = @Business2TelephoneNumber
    , [BusinessAddress] = @BusinessAddress
    , [BusinessAddressCity] = @BusinessAddressCity
    , [BusinessAddressCountry] = @BusinessAddressCountry
    , [BusinessAddressPostalCode] = @BusinessAddressPostalCode
    , [BusinessAddressPostOfficeBox] = @BusinessAddressPostOfficeBox
    , [BusinessAddressState] = @BusinessAddressState
    , [BusinessAddressStreet] = @BusinessAddressStreet
    , [BusinessFaxNumber] = @BusinessFaxNumber
    , [BusinessHomePage] = @BusinessHomePage
    , [BusinessTelephoneNumber] = @BusinessTelephoneNumber
    , [CallbackTelephoneNumber] = @CallbackTelephoneNumber
    , [CarTelephoneNumber] = @CarTelephoneNumber
    , [Categories] = @Categories
    , [Children] = @Children
    , [CompanyMainTelephoneNumber] = @CompanyMainTelephoneNumber
    , [CompanyName] = @CompanyName
    , [ComputerNetworkName] = @ComputerNetworkName
    , [CreationTime] = @CreationTime
    , [CustomerID] = @CustomerID
    , [Department] = @Department
    , [Email1Address] = @Email1Address
    , [Email2Address] = @Email2Address
    , [Email3Address] = @Email3Address
    , [FileAs] = @FileAs
    , [FirstName] = @FirstName
    , [FlagRequest] = @FlagRequest
    , [FTPSite] = @FTPSite
    , [FullName] = @FullName
    , [Gender] = @Gender
    , [GovernmentIDNumber] = @GovernmentIDNumber
    , [Hobby] = @Hobby
    , [Home2TelephoneNumber] = @Home2TelephoneNumber
    , [HomeAddress] = @HomeAddress
    , [HomeAddressCity] = @HomeAddressCity
    , [HomeAddressCountry] = @HomeAddressCountry
    , [HomeAddressPostalCode] = @HomeAddressPostalCode
    , [HomeAddressPostOfficeBox] = @HomeAddressPostOfficeBox
    , [HomeAddressState] = @HomeAddressState
    , [HomeAddressStreet] = @HomeAddressStreet
    , [HomeFaxNumber] = @HomeFaxNumber
    , [HomeTelephoneNumber] = @HomeTelephoneNumber
    , [IMAddress] = @IMAddress
    , [Initials] = @Initials
    , [InternetFreeBusyAddress] = @InternetFreeBusyAddress
    , [ISDNNumber] = @ISDNNumber
    , [JobTitle] = @JobTitle
    , [Journal] = @Journal
    , [Language] = @Language
    , [LastModificationTime] = @LastModificationTime
    , [LastName] = @LastName
    , [Location] = @Location
    , [MailingAddress] = @MailingAddress
    , [ManagerName] = @ManagerName
    , [MessageClass] = @MessageClass
    , [MiddleName] = @MiddleName
    , [Mileage] = @Mileage
    , [MobileTelephoneNumber] = @MobileTelephoneNumber
    , [NickName] = @NickName
    , [NoAging] = @NoAging
    , [OfficeLocation] = @OfficeLocation
    , [OrganizationalIDNumber] = @OrganizationalIDNumber
    , [OtherAddress] = @OtherAddress
    , [OtherAddressCity] = @OtherAddressCity
    , [OtherAddressCountry] = @OtherAddressCountry
    , [OtherAddressPostalCode] = @OtherAddressPostalCode
    , [OtherAddressPostOfficeBox] = @OtherAddressPostOfficeBox
    , [OtherAddressState] = @OtherAddressState
    , [OtherAddressStreet] = @OtherAddressStreet
    , [OtherFaxNumber] = @OtherFaxNumber
    , [OtherTelephoneNumber] = @OtherTelephoneNumber
    , [PagerNumber] = @PagerNumber
    , [PersonalHomePage] = @PersonalHomePage
    , [PrimaryTelephoneNumber] = @PrimaryTelephoneNumber
    , [Profession] = @Profession
    , [RadioTelephoneNumber] = @RadioTelephoneNumber
    , [ReferredBy] = @ReferredBy
    , [ReminderSet] = @ReminderSet
    , [ReminderTime] = @ReminderTime
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Spouse] = @Spouse
    , [Subject] = @Subject
    , [Suffix] = @Suffix
    , [TelexNumber] = @TelexNumber
    , [Title] = @Title
    , [TTYTDDTelephoneNumber] = @TTYTDDTelephoneNumber
    , [UnRead] = @UnRead
    , [User1] = @User1
    , [User2] = @User2
    , [User3] = @User3
    , [User4] = @User4
    , [WebPage] = @WebPage
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.contacts
    ( [folder_id]
    , [EntryID]
    , [Account]
    , [Anniversary]
    , [AssistantName]
    , [AssistantTelephoneNumber]
    , [Attachments_Count]
    , [BillingInformation]
    , [Birthday]
    , [Body]
    , [Business2TelephoneNumber]
    , [BusinessAddress]
    , [BusinessAddressCity]
    , [BusinessAddressCountry]
    , [BusinessAddressPostalCode]
    , [BusinessAddressPostOfficeBox]
    , [BusinessAddressState]
    , [BusinessAddressStreet]
    , [BusinessFaxNumber]
    , [BusinessHomePage]
    , [BusinessTelephoneNumber]
    , [CallbackTelephoneNumber]
    , [CarTelephoneNumber]
    , [Categories]
    , [Children]
    , [CompanyMainTelephoneNumber]
    , [CompanyName]
    , [ComputerNetworkName]
    , [CreationTime]
    , [CustomerID]
    , [Department]
    , [Email1Address]
    , [Email2Address]
    , [Email3Address]
    , [FileAs]
    , [FirstName]
    , [FlagRequest]
    , [FTPSite]
    , [FullName]
    , [Gender]
    , [GovernmentIDNumber]
    , [Hobby]
    , [Home2TelephoneNumber]
    , [HomeAddress]
    , [HomeAddressCity]
    , [HomeAddressCountry]
    , [HomeAddressPostalCode]
    , [HomeAddressPostOfficeBox]
    , [HomeAddressState]
    , [HomeAddressStreet]
    , [HomeFaxNumber]
    , [HomeTelephoneNumber]
    , [IMAddress]
    , [Initials]
    , [InternetFreeBusyAddress]
    , [ISDNNumber]
    , [JobTitle]
    , [Journal]
    , [Language]
    , [LastModificationTime]
    , [LastName]
    , [Location]
    , [MailingAddress]
    , [ManagerName]
    , [MessageClass]
    , [MiddleName]
    , [Mileage]
    , [MobileTelephoneNumber]
    , [NickName]
    , [NoAging]
    , [OfficeLocation]
    , [OrganizationalIDNumber]
    , [OtherAddress]
    , [OtherAddressCity]
    , [OtherAddressCountry]
    , [OtherAddressPostalCode]
    , [OtherAddressPostOfficeBox]
    , [OtherAddressState]
    , [OtherAddressStreet]
    , [OtherFaxNumber]
    , [OtherTelephoneNumber]
    , [PagerNumber]
    , [PersonalHomePage]
    , [PrimaryTelephoneNumber]
    , [Profession]
    , [RadioTelephoneNumber]
    , [ReferredBy]
    , [ReminderSet]
    , [ReminderTime]
    , [Sensitivity]
    , [Size]
    , [Spouse]
    , [Subject]
    , [Suffix]
    , [TelexNumber]
    , [Title]
    , [TTYTDDTelephoneNumber]
    , [UnRead]
    , [User1]
    , [User2]
    , [User3]
    , [User4]
    , [WebPage]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @Account
    , @Anniversary
    , @AssistantName
    , @AssistantTelephoneNumber
    , @Attachments_Count
    , @BillingInformation
    , @Birthday
    , @Body
    , @Business2TelephoneNumber
    , @BusinessAddress
    , @BusinessAddressCity
    , @BusinessAddressCountry
    , @BusinessAddressPostalCode
    , @BusinessAddressPostOfficeBox
    , @BusinessAddressState
    , @BusinessAddressStreet
    , @BusinessFaxNumber
    , @BusinessHomePage
    , @BusinessTelephoneNumber
    , @CallbackTelephoneNumber
    , @CarTelephoneNumber
    , @Categories
    , @Children
    , @CompanyMainTelephoneNumber
    , @CompanyName
    , @ComputerNetworkName
    , @CreationTime
    , @CustomerID
    , @Department
    , @Email1Address
    , @Email2Address
    , @Email3Address
    , @FileAs
    , @FirstName
    , @FlagRequest
    , @FTPSite
    , @FullName
    , @Gender
    , @GovernmentIDNumber
    , @Hobby
    , @Home2TelephoneNumber
    , @HomeAddress
    , @HomeAddressCity
    , @HomeAddressCountry
    , @HomeAddressPostalCode
    , @HomeAddressPostOfficeBox
    , @HomeAddressState
    , @HomeAddressStreet
    , @HomeFaxNumber
    , @HomeTelephoneNumber
    , @IMAddress
    , @Initials
    , @InternetFreeBusyAddress
    , @ISDNNumber
    , @JobTitle
    , @Journal
    , @Language
    , @LastModificationTime
    , @LastName
    , @Location
    , @MailingAddress
    , @ManagerName
    , @MessageClass
    , @MiddleName
    , @Mileage
    , @MobileTelephoneNumber
    , @NickName
    , @NoAging
    , @OfficeLocation
    , @OrganizationalIDNumber
    , @OtherAddress
    , @OtherAddressCity
    , @OtherAddressCountry
    , @OtherAddressPostalCode
    , @OtherAddressPostOfficeBox
    , @OtherAddressState
    , @OtherAddressStreet
    , @OtherFaxNumber
    , @OtherTelephoneNumber
    , @PagerNumber
    , @PersonalHomePage
    , @PrimaryTelephoneNumber
    , @Profession
    , @RadioTelephoneNumber
    , @ReferredBy
    , @ReminderSet
    , @ReminderTime
    , @Sensitivity
    , @Size
    , @Spouse
    , @Subject
    , @Suffix
    , @TelexNumber
    , @Title
    , @TTYTDDTelephoneNumber
    , @UnRead
    , @User1
    , @User2
    , @User3
    , @User4
    , @WebPage
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the juornals table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_journals_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Body] AS [Body]
    , t.[Categories] AS [Categories]
    , t.[CompanyName] AS [CompanyName]
    , t.[CreationTime] AS [CreationTime]
    , t.[Duration] AS [Duration]
    , t.[End] AS [End]
    , t.[FormDescription_ContactName] AS [FormDescription_ContactName]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[MessageClass] AS [MessageClass]
    , t.[Mileage] AS [Mileage]
    , t.[NoAging] AS [NoAging]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[Size] AS [Size]
    , t.[Start] AS [Start]
    , t.[Subject] AS [Subject]
    , t.[Type] AS [Type]
    , t.[UnRead] AS [UnRead]
FROM
    gcrm.journals t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the appointments table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_journals_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @FormDescription_ContactName nvarchar(255) = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @Type nvarchar(50) = NULL
    , @UnRead bit = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 5, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.journals WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.journals
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CompanyName] = @CompanyName
    , [CreationTime] = @CreationTime
    , [Duration] = @Duration
    , [End] = @End
    , [FormDescription_ContactName] = @FormDescription_ContactName
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Start] = @Start
    , [Subject] = @Subject
    , [Type] = @Type
    , [UnRead] = @UnRead
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.journals
    ( [folder_id]
    , [EntryID]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CompanyName]
    , [CreationTime]
    , [Duration]
    , [End]
    , [FormDescription_ContactName]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [Sensitivity]
    , [Size]
    , [Start]
    , [Subject]
    , [Type]
    , [UnRead]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @Categories
    , @CompanyName
    , @CreationTime
    , @Duration
    , @End
    , @FormDescription_ContactName
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @Sensitivity
    , @Size
    , @Start
    , @Subject
    , @Type
    , @UnRead
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the mails table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_mails_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[BCC] AS [BCC]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Body] AS [Body]
    , t.[Categories] AS [Categories]
    , t.[CC] AS [CC]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[CreationTime] AS [CreationTime]
    , t.[DeferredDeliveryTime] AS [DeferredDeliveryTime]
    , t.[ExpiryTime] AS [ExpiryTime]
    , t.[FlagDueBy] AS [FlagDueBy]
    , t.[FlagRequest] AS [FlagRequest]
    , t.[FlagStatus] AS [FlagStatus]
    , t.[Importance] AS [Importance]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[MessageClass] AS [MessageClass]
    , t.[Mileage] AS [Mileage]
    , t.[NoAging] AS [NoAging]
    , t.[ReadReceiptRequested] AS [ReadReceiptRequested]
    , t.[ReceivedTime] AS [ReceivedTime]
    , t.[RemoteStatus] AS [RemoteStatus]
    , t.[ReplyRecipientNames] AS [ReplyRecipientNames]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[SentOn] AS [SentOn]
    , t.[SentOnBehalfOfName] AS [SentOnBehalfOfName]
    , t.[Size] AS [Size]
    , t.[Subject] AS [Subject]
    , t.[UnRead] AS [UnRead]
    , t.[To] AS [To]
    , t.[TrackingStatus] AS [TrackingStatus]
FROM
    gcrm.mails t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the mails table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_mails_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @Attachments_Count bit = NULL
    , @BCC nvarchar(255) = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CC nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DeferredDeliveryTime datetime = NULL
    , @ExpiryTime datetime = NULL
    , @FlagDueBy datetime = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FlagStatus tinyint = NULL
    , @Importance tinyint = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @ReadReceiptRequested bit = NULL
    , @ReceivedTime datetime = NULL
    , @Recipients_Item_1_Address nvarchar(255) = NULL
    , @Recipients_Item_1_EntryID varchar(48) = NULL
    , @Recipients_Item_1_Name nvarchar(255) = NULL
    , @Recipients_Item_1_Type tinyint = NULL
    , @Recipients_Item_2_Address nvarchar(255) = NULL
    , @Recipients_Item_2_EntryID varchar(48) = NULL
    , @Recipients_Item_2_Name nvarchar(255) = NULL
    , @Recipients_Item_2_Type tinyint = NULL
    , @RemoteStatus tinyint = NULL
    , @ReplyRecipientNames nvarchar(255) = NULL
    , @Sender_Address nvarchar(255) = NULL
    , @Sender_Name nvarchar(255) = NULL
    , @Sensitivity tinyint = NULL
    , @SentOn datetime = NULL
    , @SentOnBehalfOfName nvarchar(255) = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @To nvarchar(255) = NULL
    , @TrackingStatus tinyint = NULL
    , @UnRead bit = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 2, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.mails WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.mails
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [Attachments_Count] = @Attachments_Count
    , [BCC] = @BCC
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CC] = @CC
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [DeferredDeliveryTime] = @DeferredDeliveryTime
    , [ExpiryTime] = @ExpiryTime
    , [FlagDueBy] = @FlagDueBy
    , [FlagRequest] = @FlagRequest
    , [FlagStatus] = @FlagStatus
    , [Importance] = @Importance
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [ReadReceiptRequested] = @ReadReceiptRequested
    , [ReceivedTime] = @ReceivedTime
    , [Recipients_Item_1_Address] = @Recipients_Item_1_Address
    , [Recipients_Item_1_EntryID] = @Recipients_Item_1_EntryID
    , [Recipients_Item_1_Name] = @Recipients_Item_1_Name
    , [Recipients_Item_1_Type] = @Recipients_Item_1_Type
    , [Recipients_Item_2_Address] = @Recipients_Item_2_Address
    , [Recipients_Item_2_EntryID] = @Recipients_Item_2_EntryID
    , [Recipients_Item_2_Name] = @Recipients_Item_2_Name
    , [Recipients_Item_2_Type] = @Recipients_Item_2_Type
    , [RemoteStatus] = @RemoteStatus
    , [ReplyRecipientNames] = @ReplyRecipientNames
    , [Sender_Address] = @Sender_Address
    , [Sender_Name] = @Sender_Name
    , [Sensitivity] = @Sensitivity
    , [SentOn] = @SentOn
    , [SentOnBehalfOfName] = @SentOnBehalfOfName
    , [Size] = @Size
    , [Subject] = @Subject
    , [To] = @To
    , [TrackingStatus] = @TrackingStatus
    , [UnRead] = @UnRead
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.mails
    ( [folder_id]
    , [EntryID]
    , [Attachments_Count]
    , [BCC]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CC]
    , [ConversationTopic]
    , [CreationTime]
    , [DeferredDeliveryTime]
    , [ExpiryTime]
    , [FlagDueBy]
    , [FlagRequest]
    , [FlagStatus]
    , [Importance]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [ReadReceiptRequested]
    , [ReceivedTime]
    , [Recipients_Item_1_Address]
    , [Recipients_Item_1_EntryID]
    , [Recipients_Item_1_Name]
    , [Recipients_Item_1_Type]
    , [Recipients_Item_2_Address]
    , [Recipients_Item_2_EntryID]
    , [Recipients_Item_2_Name]
    , [Recipients_Item_2_Type]
    , [RemoteStatus]
    , [ReplyRecipientNames]
    , [Sender_Address]
    , [Sender_Name]
    , [Sensitivity]
    , [SentOn]
    , [SentOnBehalfOfName]
    , [Size]
    , [Subject]
    , [To]
    , [TrackingStatus]
    , [UnRead]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @Attachments_Count
    , @BCC
    , @BillingInformation
    , @Body
    , @Categories
    , @CC
    , @ConversationTopic
    , @CreationTime
    , @DeferredDeliveryTime
    , @ExpiryTime
    , @FlagDueBy
    , @FlagRequest
    , @FlagStatus
    , @Importance
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @ReadReceiptRequested
    , @ReceivedTime
    , @Recipients_Item_1_Address
    , @Recipients_Item_1_EntryID
    , @Recipients_Item_1_Name
    , @Recipients_Item_1_Type
    , @Recipients_Item_2_Address
    , @Recipients_Item_2_EntryID
    , @Recipients_Item_2_Name
    , @Recipients_Item_2_Type
    , @RemoteStatus
    , @ReplyRecipientNames
    , @Sender_Address
    , @Sender_Name
    , @Sensitivity
    , @SentOn
    , @SentOnBehalfOfName
    , @Size
    , @Subject
    , @To
    , @TrackingStatus
    , @UnRead
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the notes table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_notes_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[Body] AS [Body]
    , t.[Categories] AS [Categories]
    , t.[Color] AS [Color]
    , t.[CreationTime] AS [CreationTime]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[MessageClass] AS [MessageClass]
    , t.[NoAging] AS [NoAging]
    , t.[Size] AS [Size]
    , t.[Subject] AS [Subject]
    , t.[UnRead] AS [UnRead]
FROM
    gcrm.notes t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the notes table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_notes_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @Color tinyint = NULL
    , @CreationTime datetime = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @NoAging bit = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 6, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.notes WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.notes
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [Body] = @Body
    , [Categories] = @Categories
    , [Color] = @Color
    , [CreationTime] = @CreationTime
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [NoAging] = @NoAging
    , [Size] = @Size
    , [Subject] = @Subject
    , [UnRead] = @UnRead
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.notes
    ( [folder_id]
    , [EntryID]
    , [Body]
    , [Categories]
    , [Color]
    , [CreationTime]
    , [LastModificationTime]
    , [MessageClass]
    , [NoAging]
    , [Size]
    , [Subject]
    , [UnRead]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @Body
    , @Categories
    , @Color
    , @CreationTime
    , @LastModificationTime
    , @MessageClass
    , @NoAging
    , @Size
    , @Subject
    , @UnRead
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )


END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The select procedure of the tasks table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_tasks_select]
    @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @EntryID char(48) = NULL
AS
BEGIN

SET NOCOUNT ON

SELECT
      t.[id]
    , t.[delete_in_outlook]
    , f.[StoreID] AS [StoreID]
    , t.[EntryID] AS [EntryID]
    , t.[ActualWork] AS [ActualWork]
    , t.[Attachments_Count] AS [Attachments_Count]
    , t.[BillingInformation] AS [BillingInformation]
    , t.[Body] AS [Body]
    , t.[Categories] AS [Categories]
    , t.[CompanyName] AS [CompanyName]
    , t.[Complete] AS [Complete]
    , t.[ConversationTopic] AS [ConversationTopic]
    , t.[CreationTime] AS [CreationTime]
    , t.[DateCompleted] AS [DateCompleted]
    , t.[DelegationState] AS [DelegationState]
    , t.[DueDate] AS [DueDate]
    , t.[Importance] AS [Importance]
    , t.[IsRecurring] AS [IsRecurring]
    , t.[LastModificationTime] AS [LastModificationTime]
    , t.[MessageClass] AS [MessageClass]
    , t.[Mileage] AS [Mileage]
    , t.[NoAging] AS [NoAging]
    , t.[Owner] AS [Owner]
    , t.[PercentComplete] AS [PercentComplete]
    , t.[ReminderOverrideDefault] AS [ReminderOverrideDefault]
    , t.[ReminderPlaySound] AS [ReminderPlaySound]
    , t.[ReminderSet] AS [ReminderSet]
    , t.[ReminderTime] AS [ReminderTime]
    , t.[Role] AS [Role]
    , t.[SchedulePlusPriority] AS [SchedulePlusPriority]
    , t.[Sensitivity] AS [Sensitivity]
    , t.[Size] AS [Size]
    , t.[StartDate] AS [StartDate]
    , t.[Status] AS [Status]
    , t.[Subject] AS [Subject]
    , t.[TeamTask] AS [TeamTask]
    , t.[To] AS [To]
    , t.[TotalWork] AS [TotalWork]
    , t.[UnRead] AS [UnRead]
FROM
    gcrm.tasks t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id
    INNER JOIN gcrm.users u ON u.id = f.[user_id]
WHERE
    u.id = USER_ID()
    AND f.StoreID = @StoreID
    AND f.EntryID = @FolderEntryID
    AND (@EntryID IS NULL OR t.EntryID = @EntryID)
    AND (t.last_import_time IS NULL OR t.last_import_time < t.last_update_time)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the tasks table
-- =============================================

CREATE PROCEDURE [gcrm].[usp_sync_tasks_update]
    @id int = NULL
    , @delete_in_outlook bit = NULL
    , @StoreID char(188) = NULL
    , @FolderEntryID char(48) = NULL
    , @FolderName nvarchar(255) = NULL
    , @FolderPath nvarchar(255) = NULL
    , @EntryID char(48) = NULL
    , @ActualWork int = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @Complete bit = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DateCompleted datetime = NULL
    , @DelegationState tinyint = NULL
    , @DueDate datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Owner nvarchar(255) = NULL
    , @PercentComplete float = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Role nvarchar(50) = NULL
    , @SchedulePlusPriority nvarchar(50) = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @StartDate datetime = NULL
    , @Status tinyint = NULL
    , @Subject nvarchar(255) = NULL
    , @TeamTask bit = NULL
    , @To nvarchar(255) = NULL
    , @TotalWork int = NULL
    , @UnRead bit = NULL
    , @export_time datetime = NULL
AS
BEGIN

SET NOCOUNT ON

DECLARE @folder_id int

EXEC gcrm.usp_get_folder_id @StoreID, @FolderEntryID, @FolderName, @FolderPath, 4, @folder_id OUTPUT

IF @folder_id IS NULL RETURN

IF @delete_in_outlook = 1
    BEGIN
    DELETE FROM gcrm.tasks WHERE id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID
    RETURN
    END

UPDATE gcrm.tasks
SET
    folder_id = @folder_id
    , [EntryID] = @EntryID
    , [ActualWork] = @ActualWork
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CompanyName] = @CompanyName
    , [Complete] = @Complete
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [DateCompleted] = @DateCompleted
    , [DelegationState] = @DelegationState
    , [DueDate] = @DueDate
    , [Importance] = @Importance
    , [IsRecurring] = @IsRecurring
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [Owner] = @Owner
    , [PercentComplete] = @PercentComplete
    , [ReminderOverrideDefault] = @ReminderOverrideDefault
    , [ReminderPlaySound] = @ReminderPlaySound
    , [ReminderSet] = @ReminderSet
    , [ReminderTime] = @ReminderTime
    , [Role] = @Role
    , [SchedulePlusPriority] = @SchedulePlusPriority
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [StartDate] = @StartDate
    , [Status] = @Status
    , [Subject] = @Subject
    , [TeamTask] = @TeamTask
    , [To] = @To
    , [TotalWork] = @TotalWork
    , [UnRead] = @UnRead
    , [last_import_time] = COALESCE([last_import_time], CONVERT(datetime2(0), @export_time))
    , [last_update_time] = CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
WHERE
    id = @id OR @id IS NULL AND folder_id = @folder_id AND [EntryID] = @EntryID

IF @@ROWCOUNT = 0
INSERT INTO gcrm.tasks
    ( [folder_id]
    , [EntryID]
    , [ActualWork]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CompanyName]
    , [Complete]
    , [ConversationTopic]
    , [CreationTime]
    , [DateCompleted]
    , [DelegationState]
    , [DueDate]
    , [Importance]
    , [IsRecurring]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [Owner]
    , [PercentComplete]
    , [ReminderOverrideDefault]
    , [ReminderPlaySound]
    , [ReminderSet]
    , [ReminderTime]
    , [Role]
    , [SchedulePlusPriority]
    , [Sensitivity]
    , [Size]
    , [StartDate]
    , [Status]
    , [Subject]
    , [TeamTask]
    , [To]
    , [TotalWork]
    , [UnRead]
    , [last_import_time]
    , [last_update_time]
    )
VALUES
    ( @folder_id
    , @EntryID
    , @ActualWork
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @Categories
    , @CompanyName
    , @Complete
    , @ConversationTopic
    , @CreationTime
    , @DateCompleted
    , @DelegationState
    , @DueDate
    , @Importance
    , @IsRecurring
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @Owner
    , @PercentComplete
    , @ReminderOverrideDefault
    , @ReminderPlaySound
    , @ReminderSet
    , @ReminderTime
    , @Role
    , @SchedulePlusPriority
    , @Sensitivity
    , @Size
    , @StartDate
    , @Status
    , @Subject
    , @TeamTask
    , @To
    , @TotalWork
    , @UnRead
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    , CONVERT(datetime2(0), COALESCE(@export_time, GETDATE()))
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the appointments table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_appointments_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.appointments
SET
    delete_in_outlook = 1
FROM
    gcrm.appointments t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the appointments table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_appointments_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @AllDayEvent bit = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @BusyStatus tinyint = NULL
    , @Categories nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @Location nvarchar(255) = NULL
    , @MeetingStatus tinyint = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage varchar(50) = NULL
    , @NoAging bit = NULL
    , @OptionalAttendees nvarchar(255) = NULL
    , @Organizer nvarchar(255) = NULL
    , @RecurrencePattern_PatternEndDate datetime = NULL
    , @RecurrencePattern_PatternStartDate datetime = NULL
    , @RecurrencePattern_RecurrenceType tinyint = NULL
    , @ReminderMinutesBeforeStart int = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @RequiredAttendees nvarchar(255) = NULL
    , @Resources nvarchar(255) = NULL
    , @ResponseRequested bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.appointments
    ( [folder_id]
    , [delete_in_outlook]
    , [AllDayEvent]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [BusyStatus]
    , [Categories]
    , [ConversationTopic]
    , [CreationTime]
    , [Duration]
    , [End]
    , [Importance]
    , [IsRecurring]
    , [LastModificationTime]
    , [Location]
    , [MeetingStatus]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [OptionalAttendees]
    , [Organizer]
    , [RecurrencePattern_PatternEndDate]
    , [RecurrencePattern_PatternStartDate]
    , [RecurrencePattern_RecurrenceType]
    , [ReminderMinutesBeforeStart]
    , [ReminderOverrideDefault]
    , [ReminderPlaySound]
    , [ReminderSet]
    , [RequiredAttendees]
    , [Resources]
    , [ResponseRequested]
    , [Sensitivity]
    , [Size]
    , [Start]
    , [Subject]
    , [UnRead]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @AllDayEvent
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @BusyStatus
    , @Categories
    , @ConversationTopic
    , @CreationTime
    , @Duration
    , @End
    , @Importance
    , @IsRecurring
    , @LastModificationTime
    , @Location
    , @MeetingStatus
    , @MessageClass
    , @Mileage
    , @NoAging
    , @OptionalAttendees
    , @Organizer
    , @RecurrencePattern_PatternEndDate
    , @RecurrencePattern_PatternStartDate
    , @RecurrencePattern_RecurrenceType
    , @ReminderMinutesBeforeStart
    , @ReminderOverrideDefault
    , @ReminderPlaySound
    , @ReminderSet
    , @RequiredAttendees
    , @Resources
    , @ResponseRequested
    , @Sensitivity
    , @Size
    , @Start
    , @Subject
    , @UnRead
    )


END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the appointments table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_appointments_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @AllDayEvent bit = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @BusyStatus tinyint = NULL
    , @Categories nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @Location nvarchar(255) = NULL
    , @MeetingStatus tinyint = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage varchar(50) = NULL
    , @NoAging bit = NULL
    , @OptionalAttendees nvarchar(255) = NULL
    , @Organizer nvarchar(255) = NULL
    , @RecurrencePattern_PatternEndDate datetime = NULL
    , @RecurrencePattern_PatternStartDate datetime = NULL
    , @RecurrencePattern_RecurrenceType tinyint = NULL
    , @ReminderMinutesBeforeStart int = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @RequiredAttendees nvarchar(255) = NULL
    , @Resources nvarchar(255) = NULL
    , @ResponseRequested bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.appointments
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [AllDayEvent] = @AllDayEvent
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [BusyStatus] = @BusyStatus
    , [Categories] = @Categories
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [Duration] = @Duration
    , [End] = @End
    , [Importance] = @Importance
    , [IsRecurring] = @IsRecurring
    , [LastModificationTime] = @LastModificationTime
    , [Location] = @Location
    , [MeetingStatus] = @MeetingStatus
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [OptionalAttendees] = @OptionalAttendees
    , [Organizer] = @Organizer
    , [RecurrencePattern_PatternEndDate] = @RecurrencePattern_PatternEndDate
    , [RecurrencePattern_PatternStartDate] = @RecurrencePattern_PatternStartDate
    , [RecurrencePattern_RecurrenceType] = @RecurrencePattern_RecurrenceType
    , [ReminderMinutesBeforeStart] = @ReminderMinutesBeforeStart
    , [ReminderOverrideDefault] = @ReminderOverrideDefault
    , [ReminderPlaySound] = @ReminderPlaySound
    , [ReminderSet] = @ReminderSet
    , [RequiredAttendees] = @RequiredAttendees
    , [Resources] = @Resources
    , [ResponseRequested] = @ResponseRequested
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Start] = @Start
    , [Subject] = @Subject
    , [UnRead] = @UnRead
FROM
    gcrm.appointments t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the contacts table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_contacts_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.contacts
SET
    delete_in_outlook = 1
FROM
    gcrm.contacts t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the contacts table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_contacts_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Account nvarchar(255) = NULL
    , @Anniversary datetime = NULL
    , @AssistantName nvarchar(255) = NULL
    , @AssistantTelephoneNumber nvarchar(50) = NULL
    , @Attachments_Count tinyint = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Birthday datetime = NULL
    , @Body nvarchar(max) = NULL
    , @Business2TelephoneNumber nvarchar(50) = NULL
    , @BusinessAddress nvarchar(255) = NULL
    , @BusinessAddressCity nvarchar(50) = NULL
    , @BusinessAddressCountry nvarchar(50) = NULL
    , @BusinessAddressPostalCode nvarchar(50) = NULL
    , @BusinessAddressPostOfficeBox nvarchar(50) = NULL
    , @BusinessAddressState nvarchar(50) = NULL
    , @BusinessAddressStreet nvarchar(255) = NULL
    , @BusinessFaxNumber nvarchar(50) = NULL
    , @BusinessHomePage nvarchar(50) = NULL
    , @BusinessTelephoneNumber nvarchar(50) = NULL
    , @CallbackTelephoneNumber nvarchar(50) = NULL
    , @CarTelephoneNumber nvarchar(50) = NULL
    , @Categories nvarchar(255) = NULL
    , @Children nvarchar(255) = NULL
    , @CompanyMainTelephoneNumber nvarchar(50) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @ComputerNetworkName nvarchar(50) = NULL
    , @CreationTime datetime = NULL
    , @CustomerID varchar(50) = NULL
    , @Department nvarchar(255) = NULL
    , @Email1Address varchar(100) = NULL
    , @Email2Address varchar(100) = NULL
    , @Email3Address varchar(100) = NULL
    , @FileAs nvarchar(255) = NULL
    , @FirstName nvarchar(50) = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FTPSite nvarchar(100) = NULL
    , @FullName nvarchar(255) = NULL
    , @Gender tinyint = NULL
    , @GovernmentIDNumber varchar(50) = NULL
    , @Hobby nvarchar(255) = NULL
    , @Home2TelephoneNumber nvarchar(50) = NULL
    , @HomeAddress nvarchar(255) = NULL
    , @HomeAddressCity nvarchar(50) = NULL
    , @HomeAddressCountry nvarchar(50) = NULL
    , @HomeAddressPostalCode nvarchar(50) = NULL
    , @HomeAddressPostOfficeBox nvarchar(50) = NULL
    , @HomeAddressState nvarchar(50) = NULL
    , @HomeAddressStreet nvarchar(255) = NULL
    , @HomeFaxNumber nvarchar(50) = NULL
    , @HomeTelephoneNumber nvarchar(50) = NULL
    , @IMAddress nvarchar(50) = NULL
    , @Initials nvarchar(50) = NULL
    , @InternetFreeBusyAddress varchar(255) = NULL
    , @ISDNNumber varchar(50) = NULL
    , @JobTitle nvarchar(255) = NULL
    , @Journal bit = NULL
    , @Language nvarchar(50) = NULL
    , @LastModificationTime datetime = NULL
    , @LastName nvarchar(255) = NULL
    , @Location nvarchar(255) = NULL
    , @MailingAddress nvarchar(255) = NULL
    , @ManagerName nvarchar(255) = NULL
    , @MessageClass nvarchar(100) = NULL
    , @MiddleName nvarchar(255) = NULL
    , @Mileage nvarchar(50) = NULL
    , @MobileTelephoneNumber nvarchar(50) = NULL
    , @NickName nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @OfficeLocation nvarchar(255) = NULL
    , @OrganizationalIDNumber varchar(50) = NULL
    , @OtherAddress nvarchar(255) = NULL
    , @OtherAddressCity nvarchar(50) = NULL
    , @OtherAddressCountry nvarchar(50) = NULL
    , @OtherAddressPostalCode nvarchar(50) = NULL
    , @OtherAddressPostOfficeBox nvarchar(50) = NULL
    , @OtherAddressState nvarchar(50) = NULL
    , @OtherAddressStreet nvarchar(255) = NULL
    , @OtherFaxNumber nvarchar(50) = NULL
    , @OtherTelephoneNumber nvarchar(50) = NULL
    , @PagerNumber nvarchar(50) = NULL
    , @PersonalHomePage nvarchar(255) = NULL
    , @PrimaryTelephoneNumber nvarchar(50) = NULL
    , @Profession nvarchar(50) = NULL
    , @RadioTelephoneNumber nvarchar(50) = NULL
    , @ReferredBy nvarchar(255) = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Spouse nvarchar(255) = NULL
    , @Subject nvarchar(255) = NULL
    , @Suffix nvarchar(50) = NULL
    , @TelexNumber nvarchar(50) = NULL
    , @Title nvarchar(255) = NULL
    , @TTYTDDTelephoneNumber nvarchar(50) = NULL
    , @UnRead bit = NULL
    , @User1 nvarchar(255) = NULL
    , @User2 nvarchar(255) = NULL
    , @User3 nvarchar(255) = NULL
    , @User4 nvarchar(255) = NULL
    , @WebPage nvarchar(255) = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.contacts
    ( [folder_id]
    , [delete_in_outlook]
    , [Account]
    , [Anniversary]
    , [AssistantName]
    , [AssistantTelephoneNumber]
    , [Attachments_Count]
    , [BillingInformation]
    , [Birthday]
    , [Body]
    , [Business2TelephoneNumber]
    , [BusinessAddress]
    , [BusinessAddressCity]
    , [BusinessAddressCountry]
    , [BusinessAddressPostalCode]
    , [BusinessAddressPostOfficeBox]
    , [BusinessAddressState]
    , [BusinessAddressStreet]
    , [BusinessFaxNumber]
    , [BusinessHomePage]
    , [BusinessTelephoneNumber]
    , [CallbackTelephoneNumber]
    , [CarTelephoneNumber]
    , [Categories]
    , [Children]
    , [CompanyMainTelephoneNumber]
    , [CompanyName]
    , [ComputerNetworkName]
    , [CreationTime]
    , [CustomerID]
    , [Department]
    , [Email1Address]
    , [Email2Address]
    , [Email3Address]
    , [FileAs]
    , [FirstName]
    , [FlagRequest]
    , [FTPSite]
    , [FullName]
    , [Gender]
    , [GovernmentIDNumber]
    , [Hobby]
    , [Home2TelephoneNumber]
    , [HomeAddress]
    , [HomeAddressCity]
    , [HomeAddressCountry]
    , [HomeAddressPostalCode]
    , [HomeAddressPostOfficeBox]
    , [HomeAddressState]
    , [HomeAddressStreet]
    , [HomeFaxNumber]
    , [HomeTelephoneNumber]
    , [IMAddress]
    , [Initials]
    , [InternetFreeBusyAddress]
    , [ISDNNumber]
    , [JobTitle]
    , [Journal]
    , [Language]
    , [LastModificationTime]
    , [LastName]
    , [Location]
    , [MailingAddress]
    , [ManagerName]
    , [MessageClass]
    , [MiddleName]
    , [Mileage]
    , [MobileTelephoneNumber]
    , [NickName]
    , [NoAging]
    , [OfficeLocation]
    , [OrganizationalIDNumber]
    , [OtherAddress]
    , [OtherAddressCity]
    , [OtherAddressCountry]
    , [OtherAddressPostalCode]
    , [OtherAddressPostOfficeBox]
    , [OtherAddressState]
    , [OtherAddressStreet]
    , [OtherFaxNumber]
    , [OtherTelephoneNumber]
    , [PagerNumber]
    , [PersonalHomePage]
    , [PrimaryTelephoneNumber]
    , [Profession]
    , [RadioTelephoneNumber]
    , [ReferredBy]
    , [ReminderSet]
    , [ReminderTime]
    , [Sensitivity]
    , [Size]
    , [Spouse]
    , [Subject]
    , [Suffix]
    , [TelexNumber]
    , [Title]
    , [TTYTDDTelephoneNumber]
    , [UnRead]
    , [User1]
    , [User2]
    , [User3]
    , [User4]
    , [WebPage]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @Account
    , @Anniversary
    , @AssistantName
    , @AssistantTelephoneNumber
    , @Attachments_Count
    , @BillingInformation
    , @Birthday
    , @Body
    , @Business2TelephoneNumber
    , @BusinessAddress
    , @BusinessAddressCity
    , @BusinessAddressCountry
    , @BusinessAddressPostalCode
    , @BusinessAddressPostOfficeBox
    , @BusinessAddressState
    , @BusinessAddressStreet
    , @BusinessFaxNumber
    , @BusinessHomePage
    , @BusinessTelephoneNumber
    , @CallbackTelephoneNumber
    , @CarTelephoneNumber
    , @Categories
    , @Children
    , @CompanyMainTelephoneNumber
    , @CompanyName
    , @ComputerNetworkName
    , @CreationTime
    , @CustomerID
    , @Department
    , @Email1Address
    , @Email2Address
    , @Email3Address
    , @FileAs
    , @FirstName
    , @FlagRequest
    , @FTPSite
    , @FullName
    , @Gender
    , @GovernmentIDNumber
    , @Hobby
    , @Home2TelephoneNumber
    , @HomeAddress
    , @HomeAddressCity
    , @HomeAddressCountry
    , @HomeAddressPostalCode
    , @HomeAddressPostOfficeBox
    , @HomeAddressState
    , @HomeAddressStreet
    , @HomeFaxNumber
    , @HomeTelephoneNumber
    , @IMAddress
    , @Initials
    , @InternetFreeBusyAddress
    , @ISDNNumber
    , @JobTitle
    , @Journal
    , @Language
    , @LastModificationTime
    , @LastName
    , @Location
    , @MailingAddress
    , @ManagerName
    , @MessageClass
    , @MiddleName
    , @Mileage
    , @MobileTelephoneNumber
    , @NickName
    , @NoAging
    , @OfficeLocation
    , @OrganizationalIDNumber
    , @OtherAddress
    , @OtherAddressCity
    , @OtherAddressCountry
    , @OtherAddressPostalCode
    , @OtherAddressPostOfficeBox
    , @OtherAddressState
    , @OtherAddressStreet
    , @OtherFaxNumber
    , @OtherTelephoneNumber
    , @PagerNumber
    , @PersonalHomePage
    , @PrimaryTelephoneNumber
    , @Profession
    , @RadioTelephoneNumber
    , @ReferredBy
    , @ReminderSet
    , @ReminderTime
    , @Sensitivity
    , @Size
    , @Spouse
    , @Subject
    , @Suffix
    , @TelexNumber
    , @Title
    , @TTYTDDTelephoneNumber
    , @UnRead
    , @User1
    , @User2
    , @User3
    , @User4
    , @WebPage
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the contacts table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_contacts_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Account nvarchar(255) = NULL
    , @Anniversary datetime = NULL
    , @AssistantName nvarchar(255) = NULL
    , @AssistantTelephoneNumber nvarchar(50) = NULL
    , @Attachments_Count tinyint = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Birthday datetime = NULL
    , @Body nvarchar(max) = NULL
    , @Business2TelephoneNumber nvarchar(50) = NULL
    , @BusinessAddress nvarchar(255) = NULL
    , @BusinessAddressCity nvarchar(50) = NULL
    , @BusinessAddressCountry nvarchar(50) = NULL
    , @BusinessAddressPostalCode nvarchar(50) = NULL
    , @BusinessAddressPostOfficeBox nvarchar(50) = NULL
    , @BusinessAddressState nvarchar(50) = NULL
    , @BusinessAddressStreet nvarchar(255) = NULL
    , @BusinessFaxNumber nvarchar(50) = NULL
    , @BusinessHomePage nvarchar(50) = NULL
    , @BusinessTelephoneNumber nvarchar(50) = NULL
    , @CallbackTelephoneNumber nvarchar(50) = NULL
    , @CarTelephoneNumber nvarchar(50) = NULL
    , @Categories nvarchar(255) = NULL
    , @Children nvarchar(255) = NULL
    , @CompanyMainTelephoneNumber nvarchar(50) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @ComputerNetworkName nvarchar(50) = NULL
    , @CreationTime datetime = NULL
    , @CustomerID varchar(50) = NULL
    , @Department nvarchar(255) = NULL
    , @Email1Address varchar(100) = NULL
    , @Email2Address varchar(100) = NULL
    , @Email3Address varchar(100) = NULL
    , @FileAs nvarchar(255) = NULL
    , @FirstName nvarchar(50) = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FTPSite nvarchar(100) = NULL
    , @FullName nvarchar(255) = NULL
    , @Gender tinyint = NULL
    , @GovernmentIDNumber varchar(50) = NULL
    , @Hobby nvarchar(255) = NULL
    , @Home2TelephoneNumber nvarchar(50) = NULL
    , @HomeAddress nvarchar(255) = NULL
    , @HomeAddressCity nvarchar(50) = NULL
    , @HomeAddressCountry nvarchar(50) = NULL
    , @HomeAddressPostalCode nvarchar(50) = NULL
    , @HomeAddressPostOfficeBox nvarchar(50) = NULL
    , @HomeAddressState nvarchar(50) = NULL
    , @HomeAddressStreet nvarchar(255) = NULL
    , @HomeFaxNumber nvarchar(50) = NULL
    , @HomeTelephoneNumber nvarchar(50) = NULL
    , @IMAddress nvarchar(50) = NULL
    , @Initials nvarchar(50) = NULL
    , @InternetFreeBusyAddress varchar(255) = NULL
    , @ISDNNumber varchar(50) = NULL
    , @JobTitle nvarchar(255) = NULL
    , @Journal bit = NULL
    , @Language nvarchar(50) = NULL
    , @LastModificationTime datetime = NULL
    , @LastName nvarchar(255) = NULL
    , @Location nvarchar(255) = NULL
    , @MailingAddress nvarchar(255) = NULL
    , @ManagerName nvarchar(255) = NULL
    , @MessageClass nvarchar(100) = NULL
    , @MiddleName nvarchar(255) = NULL
    , @Mileage nvarchar(50) = NULL
    , @MobileTelephoneNumber nvarchar(50) = NULL
    , @NickName nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @OfficeLocation nvarchar(255) = NULL
    , @OrganizationalIDNumber varchar(50) = NULL
    , @OtherAddress nvarchar(255) = NULL
    , @OtherAddressCity nvarchar(50) = NULL
    , @OtherAddressCountry nvarchar(50) = NULL
    , @OtherAddressPostalCode nvarchar(50) = NULL
    , @OtherAddressPostOfficeBox nvarchar(50) = NULL
    , @OtherAddressState nvarchar(50) = NULL
    , @OtherAddressStreet nvarchar(255) = NULL
    , @OtherFaxNumber nvarchar(50) = NULL
    , @OtherTelephoneNumber nvarchar(50) = NULL
    , @PagerNumber nvarchar(50) = NULL
    , @PersonalHomePage nvarchar(255) = NULL
    , @PrimaryTelephoneNumber nvarchar(50) = NULL
    , @Profession nvarchar(50) = NULL
    , @RadioTelephoneNumber nvarchar(50) = NULL
    , @ReferredBy nvarchar(255) = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Spouse nvarchar(255) = NULL
    , @Subject nvarchar(255) = NULL
    , @Suffix nvarchar(50) = NULL
    , @TelexNumber nvarchar(50) = NULL
    , @Title nvarchar(255) = NULL
    , @TTYTDDTelephoneNumber nvarchar(50) = NULL
    , @UnRead bit = NULL
    , @User1 nvarchar(255) = NULL
    , @User2 nvarchar(255) = NULL
    , @User3 nvarchar(255) = NULL
    , @User4 nvarchar(255) = NULL
    , @WebPage nvarchar(255) = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.contacts
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [Account] = @Account
    , [Anniversary] = @Anniversary
    , [AssistantName] = @AssistantName
    , [AssistantTelephoneNumber] = @AssistantTelephoneNumber
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Birthday] = @Birthday
    , [Body] = @Body
    , [Business2TelephoneNumber] = @Business2TelephoneNumber
    , [BusinessAddress] = @BusinessAddress
    , [BusinessAddressCity] = @BusinessAddressCity
    , [BusinessAddressCountry] = @BusinessAddressCountry
    , [BusinessAddressPostalCode] = @BusinessAddressPostalCode
    , [BusinessAddressPostOfficeBox] = @BusinessAddressPostOfficeBox
    , [BusinessAddressState] = @BusinessAddressState
    , [BusinessAddressStreet] = @BusinessAddressStreet
    , [BusinessFaxNumber] = @BusinessFaxNumber
    , [BusinessHomePage] = @BusinessHomePage
    , [BusinessTelephoneNumber] = @BusinessTelephoneNumber
    , [CallbackTelephoneNumber] = @CallbackTelephoneNumber
    , [CarTelephoneNumber] = @CarTelephoneNumber
    , [Categories] = @Categories
    , [Children] = @Children
    , [CompanyMainTelephoneNumber] = @CompanyMainTelephoneNumber
    , [CompanyName] = @CompanyName
    , [ComputerNetworkName] = @ComputerNetworkName
    , [CreationTime] = @CreationTime
    , [CustomerID] = @CustomerID
    , [Department] = @Department
    , [Email1Address] = @Email1Address
    , [Email2Address] = @Email2Address
    , [Email3Address] = @Email3Address
    , [FileAs] = @FileAs
    , [FirstName] = @FirstName
    , [FlagRequest] = @FlagRequest
    , [FTPSite] = @FTPSite
    , [FullName] = @FullName
    , [Gender] = @Gender
    , [GovernmentIDNumber] = @GovernmentIDNumber
    , [Hobby] = @Hobby
    , [Home2TelephoneNumber] = @Home2TelephoneNumber
    , [HomeAddress] = @HomeAddress
    , [HomeAddressCity] = @HomeAddressCity
    , [HomeAddressCountry] = @HomeAddressCountry
    , [HomeAddressPostalCode] = @HomeAddressPostalCode
    , [HomeAddressPostOfficeBox] = @HomeAddressPostOfficeBox
    , [HomeAddressState] = @HomeAddressState
    , [HomeAddressStreet] = @HomeAddressStreet
    , [HomeFaxNumber] = @HomeFaxNumber
    , [HomeTelephoneNumber] = @HomeTelephoneNumber
    , [IMAddress] = @IMAddress
    , [Initials] = @Initials
    , [InternetFreeBusyAddress] = @InternetFreeBusyAddress
    , [ISDNNumber] = @ISDNNumber
    , [JobTitle] = @JobTitle
    , [Journal] = @Journal
    , [Language] = @Language
    , [LastModificationTime] = @LastModificationTime
    , [LastName] = @LastName
    , [Location] = @Location
    , [MailingAddress] = @MailingAddress
    , [ManagerName] = @ManagerName
    , [MessageClass] = @MessageClass
    , [MiddleName] = @MiddleName
    , [Mileage] = @Mileage
    , [MobileTelephoneNumber] = @MobileTelephoneNumber
    , [NickName] = @NickName
    , [NoAging] = @NoAging
    , [OfficeLocation] = @OfficeLocation
    , [OrganizationalIDNumber] = @OrganizationalIDNumber
    , [OtherAddress] = @OtherAddress
    , [OtherAddressCity] = @OtherAddressCity
    , [OtherAddressCountry] = @OtherAddressCountry
    , [OtherAddressPostalCode] = @OtherAddressPostalCode
    , [OtherAddressPostOfficeBox] = @OtherAddressPostOfficeBox
    , [OtherAddressState] = @OtherAddressState
    , [OtherAddressStreet] = @OtherAddressStreet
    , [OtherFaxNumber] = @OtherFaxNumber
    , [OtherTelephoneNumber] = @OtherTelephoneNumber
    , [PagerNumber] = @PagerNumber
    , [PersonalHomePage] = @PersonalHomePage
    , [PrimaryTelephoneNumber] = @PrimaryTelephoneNumber
    , [Profession] = @Profession
    , [RadioTelephoneNumber] = @RadioTelephoneNumber
    , [ReferredBy] = @ReferredBy
    , [ReminderSet] = @ReminderSet
    , [ReminderTime] = @ReminderTime
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Spouse] = @Spouse
    , [Subject] = @Subject
    , [Suffix] = @Suffix
    , [TelexNumber] = @TelexNumber
    , [Title] = @Title
    , [TTYTDDTelephoneNumber] = @TTYTDDTelephoneNumber
    , [UnRead] = @UnRead
    , [User1] = @User1
    , [User2] = @User2
    , [User3] = @User3
    , [User4] = @User4
    , [WebPage] = @WebPage
FROM
    gcrm.contacts t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the journals table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_journals_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.journals
SET
    delete_in_outlook = 1
FROM
    gcrm.journals t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the appointments table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_journals_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @FormDescription_ContactName nvarchar(255) = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @Type nvarchar(50) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.journals
    ( [folder_id]
    , [delete_in_outlook]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CompanyName]
    , [CreationTime]
    , [Duration]
    , [End]
    , [FormDescription_ContactName]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [Sensitivity]
    , [Size]
    , [Start]
    , [Subject]
    , [Type]
    , [UnRead]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @Categories
    , @CompanyName
    , @CreationTime
    , @Duration
    , @End
    , @FormDescription_ContactName
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @Sensitivity
    , @Size
    , @Start
    , @Subject
    , @Type
    , @UnRead
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the appointments table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_journals_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @Duration int = NULL
    , @End datetime = NULL
    , @FormDescription_ContactName nvarchar(255) = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @Start datetime = NULL
    , @Subject nvarchar(255) = NULL
    , @Type nvarchar(50) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.journals
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CompanyName] = @CompanyName
    , [CreationTime] = @CreationTime
    , [Duration] = @Duration
    , [End] = @End
    , [FormDescription_ContactName] = @FormDescription_ContactName
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [Start] = @Start
    , [Subject] = @Subject
    , [Type] = @Type
    , [UnRead] = @UnRead
FROM
    gcrm.journals t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the busy_status_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_busy_status_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.busy_statuses t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the folder_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_folder_id]
    @folder_type_id tinyint = NULL
AS
BEGIN

SET NOCOUNT ON

IF (SELECT COUNT(DISTINCT StoreID) FROM gcrm.folders WHERE [user_id] = USER_ID()) = 1
    SELECT
        f.id
        , f.Name AS name
    FROM
        gcrm.folders f
    WHERE    
        f.[user_id] = USER_ID()
        AND f.folder_type_id = COALESCE(@folder_type_id, f.folder_type_id)
ELSE
    SELECT
        f.id
        , f.FolderPath AS name
    FROM
        gcrm.folders f
    WHERE    
        f.[user_id] = USER_ID()
        AND f.folder_type_id = COALESCE(@folder_type_id, f.folder_type_id)

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the importance_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_importance_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.importance_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the journal_recipient_type_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_journal_recipient_type_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.journal_recipient_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the mail_recipient_type_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_mail_recipient_type_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.mail_recipient_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the meeting_recipient_type_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_meeting_recipient_type_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.meeting_recipient_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the sensitivity_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_sensitivity_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.sensitivity_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the task_recipient_type_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_task_recipient_type_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.task_recipient_types t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The procedure selects the task_status_id validation list for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_list_task_status_id]
AS
BEGIN

SET NOCOUNT ON

SELECT
    t.id
    , t.name
FROM
    gcrm.task_statuses t
ORDER BY
    t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the mails table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_mails_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.mails
SET
    delete_in_outlook = 1
FROM
    gcrm.mails t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the mails table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_mails_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Attachments_Count bit = NULL
    , @BCC nvarchar(255) = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CC nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DeferredDeliveryTime datetime = NULL
    , @ExpiryTime datetime = NULL
    , @FlagDueBy datetime = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FlagStatus tinyint = NULL
    , @Importance tinyint = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @ReadReceiptRequested bit = NULL
    , @ReceivedTime datetime = NULL
    , @Recipients_Item_1_Address nvarchar(255) = NULL
    , @Recipients_Item_1_Name nvarchar(255) = NULL
    , @RemoteStatus tinyint = NULL
    , @ReplyRecipientNames nvarchar(255) = NULL
    , @Sender_Address nvarchar(255) = NULL
    , @Sender_Name nvarchar(255) = NULL
    , @Sensitivity tinyint = NULL
    , @SentOn datetime = NULL
    , @SentOnBehalfOfName nvarchar(255) = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @To nvarchar(255) = NULL
    , @TrackingStatus tinyint = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.mails
    ( [folder_id]
    , [delete_in_outlook]
    , [Attachments_Count]
    , [BCC]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CC]
    , [ConversationTopic]
    , [CreationTime]
    , [DeferredDeliveryTime]
    , [ExpiryTime]
    , [FlagDueBy]
    , [FlagRequest]
    , [FlagStatus]
    , [Importance]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [ReadReceiptRequested]
    , [ReceivedTime]
    , [Recipients_Item_1_Address]
    , [Recipients_Item_1_Name]
    , [RemoteStatus]
    , [ReplyRecipientNames]
    , [Sender_Address]
    , [Sender_Name]
    , [Sensitivity]
    , [SentOn]
    , [SentOnBehalfOfName]
    , [Size]
    , [Subject]
    , [To]
    , [TrackingStatus]
    , [UnRead]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @Attachments_Count
    , @BCC
    , @BillingInformation
    , @Body
    , @Categories
    , @CC
    , @ConversationTopic
    , @CreationTime
    , @DeferredDeliveryTime
    , @ExpiryTime
    , @FlagDueBy
    , @FlagRequest
    , @FlagStatus
    , @Importance
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @ReadReceiptRequested
    , @ReceivedTime
    , @Recipients_Item_1_Address
    , @Recipients_Item_1_Name
    , @RemoteStatus
    , @ReplyRecipientNames
    , @Sender_Address
    , @Sender_Name
    , @Sensitivity
    , @SentOn
    , @SentOnBehalfOfName
    , @Size
    , @Subject
    , @To
    , @TrackingStatus
    , @UnRead
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the mails table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_mails_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Attachments_Count bit = NULL
    , @BCC nvarchar(255) = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CC nvarchar(255) = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DeferredDeliveryTime datetime = NULL
    , @ExpiryTime datetime = NULL
    , @FlagDueBy datetime = NULL
    , @FlagRequest nvarchar(50) = NULL
    , @FlagStatus tinyint = NULL
    , @Importance tinyint = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @ReadReceiptRequested bit = NULL
    , @ReceivedTime datetime = NULL
    , @Recipients_Item_1_Address nvarchar(255) = NULL
    , @Recipients_Item_1_Name nvarchar(255) = NULL
    , @RemoteStatus tinyint = NULL
    , @ReplyRecipientNames nvarchar(255) = NULL
    , @Sender_Address nvarchar(255) = NULL
    , @Sender_Name nvarchar(255) = NULL
    , @Sensitivity tinyint = NULL
    , @SentOn datetime = NULL
    , @SentOnBehalfOfName nvarchar(255) = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @To nvarchar(255) = NULL
    , @TrackingStatus tinyint = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.mails
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [Attachments_Count] = @Attachments_Count
    , [BCC] = @BCC
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CC] = @CC
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [DeferredDeliveryTime] = @DeferredDeliveryTime
    , [ExpiryTime] = @ExpiryTime
    , [FlagDueBy] = @FlagDueBy
    , [FlagRequest] = @FlagRequest
    , [FlagStatus] = @FlagStatus
    , [Importance] = @Importance
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [ReadReceiptRequested] = @ReadReceiptRequested
    , [ReceivedTime] = @ReceivedTime
    , [Recipients_Item_1_Address] = @Recipients_Item_1_Address
    , [Recipients_Item_1_Name] = @Recipients_Item_1_Name
    , [RemoteStatus] = @RemoteStatus
    , [ReplyRecipientNames] = @ReplyRecipientNames
    , [Sender_Address] = @Sender_Address
    , [Sender_Name] = @Sender_Name
    , [Sensitivity] = @Sensitivity
    , [SentOn] = @SentOn
    , [SentOnBehalfOfName] = @SentOnBehalfOfName
    , [Size] = @Size
    , [Subject] = @Subject
    , [To] = @To
    , [TrackingStatus] = @TrackingStatus
    , [UnRead] = @UnRead
FROM
    gcrm.mails t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the notes table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_notes_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.notes
SET
    delete_in_outlook = 1
FROM
    gcrm.notes t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the notes table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_notes_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @Color tinyint = NULL
    , @CreationTime datetime = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @NoAging bit = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.notes
    ( [folder_id]
    , [delete_in_outlook]
    , [Body]
    , [Categories]
    , [Color]
    , [CreationTime]
    , [LastModificationTime]
    , [MessageClass]
    , [NoAging]
    , [Size]
    , [Subject]
    , [UnRead]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @Body
    , @Categories
    , @Color
    , @CreationTime
    , @LastModificationTime
    , @MessageClass
    , @NoAging
    , @Size
    , @Subject
    , @UnRead
    )


END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the notes table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_notes_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @Color tinyint = NULL
    , @CreationTime datetime = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @NoAging bit = NULL
    , @Size int = NULL
    , @Subject nvarchar(255) = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.notes
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [Body] = @Body
    , [Categories] = @Categories
    , [Color] = @Color
    , [CreationTime] = @CreationTime
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [NoAging] = @NoAging
    , [Size] = @Size
    , [Subject] = @Subject
    , [UnRead] = @UnRead
FROM
    gcrm.notes t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id


END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The delete procedure of the tasks table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_tasks_delete]
    @id int = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.tasks
SET
    delete_in_outlook = 1
FROM
    gcrm.tasks t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id
      
END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The insert procedure of the tasks table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_tasks_insert]
    @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @ActualWork int = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @Complete bit = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DateCompleted datetime = NULL
    , @DelegationState tinyint = NULL
    , @DueDate datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Owner nvarchar(255) = NULL
    , @PercentComplete float = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Role nvarchar(50) = NULL
    , @SchedulePlusPriority nvarchar(50) = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @StartDate datetime = NULL
    , @Status tinyint = NULL
    , @Subject nvarchar(255) = NULL
    , @TeamTask bit = NULL
    , @To nvarchar(255) = NULL
    , @TotalWork int = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

INSERT INTO gcrm.tasks
    ( [folder_id]
    , [delete_in_outlook]
    , [ActualWork]
    , [Attachments_Count]
    , [BillingInformation]
    , [Body]
    , [Categories]
    , [CompanyName]
    , [Complete]
    , [ConversationTopic]
    , [CreationTime]
    , [DateCompleted]
    , [DelegationState]
    , [DueDate]
    , [Importance]
    , [IsRecurring]
    , [LastModificationTime]
    , [MessageClass]
    , [Mileage]
    , [NoAging]
    , [Owner]
    , [PercentComplete]
    , [ReminderOverrideDefault]
    , [ReminderPlaySound]
    , [ReminderSet]
    , [ReminderTime]
    , [Role]
    , [SchedulePlusPriority]
    , [Sensitivity]
    , [Size]
    , [StartDate]
    , [Status]
    , [Subject]
    , [TeamTask]
    , [To]
    , [TotalWork]
    , [UnRead]
    )
VALUES
    ( @folder_id
    , @delete_in_outlook
    , @ActualWork
    , @Attachments_Count
    , @BillingInformation
    , @Body
    , @Categories
    , @CompanyName
    , @Complete
    , @ConversationTopic
    , @CreationTime
    , @DateCompleted
    , @DelegationState
    , @DueDate
    , @Importance
    , @IsRecurring
    , @LastModificationTime
    , @MessageClass
    , @Mileage
    , @NoAging
    , @Owner
    , @PercentComplete
    , @ReminderOverrideDefault
    , @ReminderPlaySound
    , @ReminderSet
    , @ReminderTime
    , @Role
    , @SchedulePlusPriority
    , @Sensitivity
    , @Size
    , @StartDate
    , @Status
    , @Subject
    , @TeamTask
    , @To
    , @TotalWork
    , @UnRead
    )

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: The update procedure of the tasks table for the SaveToDB add-in
-- =============================================

CREATE PROCEDURE [gcrm].[usp_xl_tasks_update]
    @id int = NULL
    , @folder_id int = NULL
    , @delete_in_outlook bit = NULL
    , @ActualWork int = NULL
    , @Attachments_Count bit = NULL
    , @BillingInformation nvarchar(255) = NULL
    , @Body nvarchar(max) = NULL
    , @Categories nvarchar(255) = NULL
    , @CompanyName nvarchar(255) = NULL
    , @Complete bit = NULL
    , @ConversationTopic nvarchar(255) = NULL
    , @CreationTime datetime = NULL
    , @DateCompleted datetime = NULL
    , @DelegationState tinyint = NULL
    , @DueDate datetime = NULL
    , @Importance tinyint = NULL
    , @IsRecurring bit = NULL
    , @LastModificationTime datetime = NULL
    , @MessageClass nvarchar(100) = NULL
    , @Mileage nvarchar(50) = NULL
    , @NoAging bit = NULL
    , @Owner nvarchar(255) = NULL
    , @PercentComplete float = NULL
    , @ReminderOverrideDefault bit = NULL
    , @ReminderPlaySound bit = NULL
    , @ReminderSet bit = NULL
    , @ReminderTime datetime = NULL
    , @Role nvarchar(50) = NULL
    , @SchedulePlusPriority nvarchar(50) = NULL
    , @Sensitivity tinyint = NULL
    , @Size int = NULL
    , @StartDate datetime = NULL
    , @Status tinyint = NULL
    , @Subject nvarchar(255) = NULL
    , @TeamTask bit = NULL
    , @To nvarchar(255) = NULL
    , @TotalWork int = NULL
    , @UnRead bit = NULL
AS
BEGIN

SET NOCOUNT ON

UPDATE gcrm.tasks
SET
    folder_id = @folder_id
    , delete_in_outlook = @delete_in_outlook
    , [ActualWork] = @ActualWork
    , [Attachments_Count] = @Attachments_Count
    , [BillingInformation] = @BillingInformation
    , [Body] = @Body
    , [Categories] = @Categories
    , [CompanyName] = @CompanyName
    , [Complete] = @Complete
    , [ConversationTopic] = @ConversationTopic
    , [CreationTime] = @CreationTime
    , [DateCompleted] = @DateCompleted
    , [DelegationState] = @DelegationState
    , [DueDate] = @DueDate
    , [Importance] = @Importance
    , [IsRecurring] = @IsRecurring
    , [LastModificationTime] = @LastModificationTime
    , [MessageClass] = @MessageClass
    , [Mileage] = @Mileage
    , [NoAging] = @NoAging
    , [Owner] = @Owner
    , [PercentComplete] = @PercentComplete
    , [ReminderOverrideDefault] = @ReminderOverrideDefault
    , [ReminderPlaySound] = @ReminderPlaySound
    , [ReminderSet] = @ReminderSet
    , [ReminderTime] = @ReminderTime
    , [Role] = @Role
    , [SchedulePlusPriority] = @SchedulePlusPriority
    , [Sensitivity] = @Sensitivity
    , [Size] = @Size
    , [StartDate] = @StartDate
    , [Status] = @Status
    , [Subject] = @Subject
    , [TeamTask] = @TeamTask
    , [To] = @To
    , [TotalWork] = @TotalWork
    , [UnRead] = @UnRead
FROM
    gcrm.tasks t
    INNER JOIN gcrm.folders f ON f.id = t.folder_id AND f.[user_id] = USER_ID()
WHERE
    t.id = @id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_appointments_update]
   ON [gcrm].[appointments]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.appointments
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.appointments t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_contacts_update]
   ON [gcrm].[contacts]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.contacts
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.contacts t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_journals_update]
   ON [gcrm].[journals]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.journals
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.journals t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_mails_update]
   ON [gcrm].[mails]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.mails
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.mails t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_notes_update]
   ON [gcrm].[notes]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.notes
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.notes t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     2.0, 2022-07-05
-- Description: Updates the last_update_time field
-- =============================================

CREATE TRIGGER [gcrm].[trigger_tasks_update]
   ON [gcrm].[tasks]
   AFTER UPDATE
AS 
BEGIN

SET NOCOUNT ON

IF NOT UPDATE(last_update_time)

    UPDATE gcrm.tasks
    SET
        last_update_time = CONVERT(datetime2(0), GETDATE())
    FROM
        gcrm.tasks t
        INNER JOIN inserted ON inserted.id = t.id

END


GO

INSERT INTO [gcrm].[busy_statuses] ([id], [name]) VALUES (2, N'Busy');
INSERT INTO [gcrm].[busy_statuses] ([id], [name]) VALUES (0, N'Free');
INSERT INTO [gcrm].[busy_statuses] ([id], [name]) VALUES (3, N'OutOfOffice');
INSERT INTO [gcrm].[busy_statuses] ([id], [name]) VALUES (1, N'Tentative');
INSERT INTO [gcrm].[busy_statuses] ([id], [name]) VALUES (4, N'WorkingElsewhere');
GO

INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (3, N'Appointments');
INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (1, N'Contacts');
INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (5, N'Journlas');
INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (2, N'Mails');
INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (6, N'Notes');
INSERT INTO [gcrm].[folder_types] ([id], [folder_type]) VALUES (4, N'Tasks');
GO

INSERT INTO [gcrm].[formats] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_EXCEL_FORMAT_XML]) VALUES (N'gcrm', N'formats', N'<table name="gcrm.formats"><columnFormats><column name="" property="ListObjectName" value="formats" type="String" /><column name="" property="ShowTotals" value="False" type="Boolean" /><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String" /><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean" /><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean" /><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean" /><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean" /><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="_RowNum" property="Address" value="$B$4" type="String" /><column name="_RowNum" property="NumberFormat" value="General" type="String" /><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double" /><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="ID" property="Address" value="$C$4" type="String" /><column name="ID" property="NumberFormat" value="General" type="String" /><column name="ID" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String" /><column name="TABLE_SCHEMA" property="ColumnWidth" value="16.57" type="Double" /><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="TABLE_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_NAME" property="Address" value="$E$4" type="String" /><column name="TABLE_NAME" property="ColumnWidth" value="30" type="Double" /><column name="TABLE_NAME" property="NumberFormat" value="General" type="String" /><column name="TABLE_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_EXCEL_FORMAT_XML" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_EXCEL_FORMAT_XML" property="Address" value="$F$4" type="String" /><column name="TABLE_EXCEL_FORMAT_XML" property="ColumnWidth" value="42.29" type="Double" /><column name="TABLE_EXCEL_FORMAT_XML" property="NumberFormat" value="General" type="String" /><column name="SortFields(1)" property="KeyfieldName" value="TABLE_SCHEMA" type="String" /><column name="SortFields(1)" property="SortOn" value="0" type="Double" /><column name="SortFields(1)" property="Order" value="1" type="Double" /><column name="SortFields(1)" property="DataOption" value="0" type="Double" /><column name="SortFields(2)" property="KeyfieldName" value="TABLE_NAME" type="String" /><column name="SortFields(2)" property="SortOn" value="0" type="Double" /><column name="SortFields(2)" property="Order" value="1" type="Double" /><column name="SortFields(2)" property="DataOption" value="0" type="Double" /><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean" /><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean" /><column name="" property="ActiveWindow.Split" value="True" type="Boolean" /><column name="" property="ActiveWindow.SplitRow" value="0" type="Double" /><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double" /><column name="" property="PageSetup.Orientation" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double" /></columnFormats></table>');
INSERT INTO [gcrm].[formats] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_EXCEL_FORMAT_XML]) VALUES (N'gcrm', N'handlers', N'<table name="gcrm.handlers"><columnFormats><column name="" property="ListObjectName" value="handlers" type="String" /><column name="" property="ShowTotals" value="False" type="Boolean" /><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String" /><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean" /><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean" /><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean" /><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean" /><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="_RowNum" property="Address" value="$B$4" type="String" /><column name="_RowNum" property="NumberFormat" value="General" type="String" /><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double" /><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="ID" property="Address" value="$C$4" type="String" /><column name="ID" property="NumberFormat" value="General" type="String" /><column name="ID" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String" /><column name="TABLE_SCHEMA" property="ColumnWidth" value="16.57" type="Double" /><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="TABLE_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_NAME" property="Address" value="$E$4" type="String" /><column name="TABLE_NAME" property="ColumnWidth" value="30" type="Double" /><column name="TABLE_NAME" property="NumberFormat" value="General" type="String" /><column name="TABLE_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="COLUMN_NAME" property="Address" value="$F$4" type="String" /><column name="COLUMN_NAME" property="ColumnWidth" value="17.43" type="Double" /><column name="COLUMN_NAME" property="NumberFormat" value="General" type="String" /><column name="COLUMN_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="EVENT_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="EVENT_NAME" property="Address" value="$G$4" type="String" /><column name="EVENT_NAME" property="ColumnWidth" value="21.57" type="Double" /><column name="EVENT_NAME" property="NumberFormat" value="General" type="String" /><column name="EVENT_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="HANDLER_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="HANDLER_SCHEMA" property="Address" value="$H$4" type="String" /><column name="HANDLER_SCHEMA" property="ColumnWidth" value="19.71" type="Double" /><column name="HANDLER_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="HANDLER_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="HANDLER_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="HANDLER_NAME" property="Address" value="$I$4" type="String" /><column name="HANDLER_NAME" property="ColumnWidth" value="31.14" type="Double" /><column name="HANDLER_NAME" property="NumberFormat" value="General" type="String" /><column name="HANDLER_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="HANDLER_TYPE" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="HANDLER_TYPE" property="Address" value="$J$4" type="String" /><column name="HANDLER_TYPE" property="ColumnWidth" value="16.29" type="Double" /><column name="HANDLER_TYPE" property="NumberFormat" value="General" type="String" /><column name="HANDLER_TYPE" property="VerticalAlignment" value="-4160" type="Double" /><column name="HANDLER_CODE" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="HANDLER_CODE" property="Address" value="$K$4" type="String" /><column name="HANDLER_CODE" property="ColumnWidth" value="70.71" type="Double" /><column name="HANDLER_CODE" property="NumberFormat" value="General" type="String" /><column name="HANDLER_CODE" property="VerticalAlignment" value="-4160" type="Double" /><column name="TARGET_WORKSHEET" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TARGET_WORKSHEET" property="Address" value="$L$4" type="String" /><column name="TARGET_WORKSHEET" property="ColumnWidth" value="21.71" type="Double" /><column name="TARGET_WORKSHEET" property="NumberFormat" value="General" type="String" /><column name="TARGET_WORKSHEET" property="VerticalAlignment" value="-4160" type="Double" /><column name="MENU_ORDER" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="MENU_ORDER" property="Address" value="$M$4" type="String" /><column name="MENU_ORDER" property="ColumnWidth" value="15.43" type="Double" /><column name="MENU_ORDER" property="NumberFormat" value="General" type="String" /><column name="MENU_ORDER" property="VerticalAlignment" value="-4160" type="Double" /><column name="EDIT_PARAMETERS" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="EDIT_PARAMETERS" property="Address" value="$N$4" type="String" /><column name="EDIT_PARAMETERS" property="ColumnWidth" value="19.57" type="Double" /><column name="EDIT_PARAMETERS" property="NumberFormat" value="General" type="String" /><column name="EDIT_PARAMETERS" property="HorizontalAlignment" value="-4108" type="Double" /><column name="EDIT_PARAMETERS" property="VerticalAlignment" value="-4160" type="Double" /><column name="EDIT_PARAMETERS" property="Font.Size" value="10" type="Double" /><column name="SortFields(1)" property="KeyfieldName" value="EVENT_NAME" type="String" /><column name="SortFields(1)" property="SortOn" value="0" type="Double" /><column name="SortFields(1)" property="Order" value="1" type="Double" /><column name="SortFields(1)" property="DataOption" value="0" type="Double" /><column name="SortFields(2)" property="KeyfieldName" value="TABLE_SCHEMA" type="String" /><column name="SortFields(2)" property="SortOn" value="0" type="Double" /><column name="SortFields(2)" property="Order" value="1" type="Double" /><column name="SortFields(2)" property="DataOption" value="0" type="Double" /><column name="SortFields(3)" property="KeyfieldName" value="TABLE_NAME" type="String" /><column name="SortFields(3)" property="SortOn" value="0" type="Double" /><column name="SortFields(3)" property="Order" value="1" type="Double" /><column name="SortFields(3)" property="DataOption" value="0" type="Double" /><column name="SortFields(4)" property="KeyfieldName" value="COLUMN_NAME" type="String" /><column name="SortFields(4)" property="SortOn" value="0" type="Double" /><column name="SortFields(4)" property="Order" value="1" type="Double" /><column name="SortFields(4)" property="DataOption" value="0" type="Double" /><column name="SortFields(5)" property="KeyfieldName" value="MENU_ORDER" type="String" /><column name="SortFields(5)" property="SortOn" value="0" type="Double" /><column name="SortFields(5)" property="Order" value="1" type="Double" /><column name="SortFields(5)" property="DataOption" value="0" type="Double" /><column name="SortFields(6)" property="KeyfieldName" value="HANDLER_SCHEMA" type="String" /><column name="SortFields(6)" property="SortOn" value="0" type="Double" /><column name="SortFields(6)" property="Order" value="1" type="Double" /><column name="SortFields(6)" property="DataOption" value="0" type="Double" /><column name="SortFields(7)" property="KeyfieldName" value="HANDLER_NAME" type="String" /><column name="SortFields(7)" property="SortOn" value="0" type="Double" /><column name="SortFields(7)" property="Order" value="1" type="Double" /><column name="SortFields(7)" property="DataOption" value="0" type="Double" /><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean" /><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean" /><column name="" property="ActiveWindow.Split" value="True" type="Boolean" /><column name="" property="ActiveWindow.SplitRow" value="0" type="Double" /><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double" /><column name="" property="PageSetup.Orientation" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double" /></columnFormats></table>');
INSERT INTO [gcrm].[formats] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_EXCEL_FORMAT_XML]) VALUES (N'gcrm', N'objects', N'<table name="gcrm.objects"><columnFormats><column name="" property="ListObjectName" value="objects" type="String" /><column name="" property="ShowTotals" value="False" type="Boolean" /><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String" /><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean" /><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean" /><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean" /><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean" /><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="_RowNum" property="Address" value="$B$4" type="String" /><column name="_RowNum" property="NumberFormat" value="General" type="String" /><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double" /><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="ID" property="Address" value="$C$4" type="String" /><column name="ID" property="NumberFormat" value="General" type="String" /><column name="ID" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String" /><column name="TABLE_SCHEMA" property="ColumnWidth" value="16.57" type="Double" /><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="TABLE_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_NAME" property="Address" value="$E$4" type="String" /><column name="TABLE_NAME" property="ColumnWidth" value="30" type="Double" /><column name="TABLE_NAME" property="NumberFormat" value="General" type="String" /><column name="TABLE_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_TYPE" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_TYPE" property="Address" value="$F$4" type="String" /><column name="TABLE_TYPE" property="ColumnWidth" value="13.14" type="Double" /><column name="TABLE_TYPE" property="NumberFormat" value="General" type="String" /><column name="TABLE_TYPE" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_CODE" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_CODE" property="Address" value="$G$4" type="String" /><column name="TABLE_CODE" property="ColumnWidth" value="13.71" type="Double" /><column name="TABLE_CODE" property="NumberFormat" value="General" type="String" /><column name="TABLE_CODE" property="VerticalAlignment" value="-4160" type="Double" /><column name="INSERT_OBJECT" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="INSERT_OBJECT" property="Address" value="$H$4" type="String" /><column name="INSERT_OBJECT" property="ColumnWidth" value="27.86" type="Double" /><column name="INSERT_OBJECT" property="NumberFormat" value="General" type="String" /><column name="INSERT_OBJECT" property="VerticalAlignment" value="-4160" type="Double" /><column name="UPDATE_OBJECT" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="UPDATE_OBJECT" property="Address" value="$I$4" type="String" /><column name="UPDATE_OBJECT" property="ColumnWidth" value="27.86" type="Double" /><column name="UPDATE_OBJECT" property="NumberFormat" value="General" type="String" /><column name="UPDATE_OBJECT" property="VerticalAlignment" value="-4160" type="Double" /><column name="DELETE_OBJECT" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="DELETE_OBJECT" property="Address" value="$J$4" type="String" /><column name="DELETE_OBJECT" property="ColumnWidth" value="27.86" type="Double" /><column name="DELETE_OBJECT" property="NumberFormat" value="General" type="String" /><column name="DELETE_OBJECT" property="VerticalAlignment" value="-4160" type="Double" /><column name="SortFields(1)" property="KeyfieldName" value="TABLE_SCHEMA" type="String" /><column name="SortFields(1)" property="SortOn" value="0" type="Double" /><column name="SortFields(1)" property="Order" value="1" type="Double" /><column name="SortFields(1)" property="DataOption" value="0" type="Double" /><column name="SortFields(2)" property="KeyfieldName" value="TABLE_NAME" type="String" /><column name="SortFields(2)" property="SortOn" value="0" type="Double" /><column name="SortFields(2)" property="Order" value="1" type="Double" /><column name="SortFields(2)" property="DataOption" value="0" type="Double" /><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean" /><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean" /><column name="" property="ActiveWindow.Split" value="True" type="Boolean" /><column name="" property="ActiveWindow.SplitRow" value="0" type="Double" /><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double" /><column name="" property="PageSetup.Orientation" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double" /></columnFormats></table>');
INSERT INTO [gcrm].[formats] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_EXCEL_FORMAT_XML]) VALUES (N'gcrm', N'translations', N'<table name="gcrm.translations"><columnFormats><column name="" property="ListObjectName" value="translations" type="String" /><column name="" property="ShowTotals" value="False" type="Boolean" /><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String" /><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean" /><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean" /><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean" /><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean" /><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="_RowNum" property="Address" value="$B$4" type="String" /><column name="_RowNum" property="NumberFormat" value="General" type="String" /><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double" /><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="ID" property="Address" value="$C$4" type="String" /><column name="ID" property="NumberFormat" value="General" type="String" /><column name="ID" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_SCHEMA" property="Address" value="$D$4" type="String" /><column name="TABLE_SCHEMA" property="ColumnWidth" value="16.57" type="Double" /><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="TABLE_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="TABLE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_NAME" property="Address" value="$E$4" type="String" /><column name="TABLE_NAME" property="ColumnWidth" value="32.14" type="Double" /><column name="TABLE_NAME" property="NumberFormat" value="General" type="String" /><column name="TABLE_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="COLUMN_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="COLUMN_NAME" property="Address" value="$F$4" type="String" /><column name="COLUMN_NAME" property="ColumnWidth" value="20.71" type="Double" /><column name="COLUMN_NAME" property="NumberFormat" value="General" type="String" /><column name="COLUMN_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="LANGUAGE_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="LANGUAGE_NAME" property="Address" value="$G$4" type="String" /><column name="LANGUAGE_NAME" property="ColumnWidth" value="19.57" type="Double" /><column name="LANGUAGE_NAME" property="NumberFormat" value="General" type="String" /><column name="LANGUAGE_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="TRANSLATED_NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TRANSLATED_NAME" property="Address" value="$H$4" type="String" /><column name="TRANSLATED_NAME" property="ColumnWidth" value="30" type="Double" /><column name="TRANSLATED_NAME" property="NumberFormat" value="General" type="String" /><column name="TRANSLATED_NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="TRANSLATED_DESC" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TRANSLATED_DESC" property="Address" value="$I$4" type="String" /><column name="TRANSLATED_DESC" property="ColumnWidth" value="19.57" type="Double" /><column name="TRANSLATED_DESC" property="NumberFormat" value="General" type="String" /><column name="TRANSLATED_DESC" property="VerticalAlignment" value="-4160" type="Double" /><column name="TRANSLATED_COMMENT" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TRANSLATED_COMMENT" property="Address" value="$J$4" type="String" /><column name="TRANSLATED_COMMENT" property="ColumnWidth" value="25" type="Double" /><column name="TRANSLATED_COMMENT" property="NumberFormat" value="General" type="String" /><column name="TRANSLATED_COMMENT" property="VerticalAlignment" value="-4160" type="Double" /><column name="SortFields(1)" property="KeyfieldName" value="LANGUAGE_NAME" type="String" /><column name="SortFields(1)" property="SortOn" value="0" type="Double" /><column name="SortFields(1)" property="Order" value="1" type="Double" /><column name="SortFields(1)" property="DataOption" value="2" type="Double" /><column name="SortFields(2)" property="KeyfieldName" value="TABLE_SCHEMA" type="String" /><column name="SortFields(2)" property="SortOn" value="0" type="Double" /><column name="SortFields(2)" property="Order" value="1" type="Double" /><column name="SortFields(2)" property="DataOption" value="2" type="Double" /><column name="SortFields(3)" property="KeyfieldName" value="TABLE_NAME" type="String" /><column name="SortFields(3)" property="SortOn" value="0" type="Double" /><column name="SortFields(3)" property="Order" value="1" type="Double" /><column name="SortFields(3)" property="DataOption" value="2" type="Double" /><column name="SortFields(4)" property="KeyfieldName" value="COLUMN_NAME" type="String" /><column name="SortFields(4)" property="SortOn" value="0" type="Double" /><column name="SortFields(4)" property="Order" value="1" type="Double" /><column name="SortFields(4)" property="DataOption" value="2" type="Double" /><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean" /><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean" /><column name="" property="ActiveWindow.Split" value="True" type="Boolean" /><column name="" property="ActiveWindow.SplitRow" value="0" type="Double" /><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double" /><column name="" property="PageSetup.Orientation" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double" /></columnFormats></table>');
INSERT INTO [gcrm].[formats] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_EXCEL_FORMAT_XML]) VALUES (N'gcrm', N'workbooks', N'<table name="gcrm.workbooks"><columnFormats><column name="" property="ListObjectName" value="workbooks" type="String" /><column name="" property="ShowTotals" value="False" type="Boolean" /><column name="" property="TableStyle.Name" value="TableStyleMedium15" type="String" /><column name="" property="ShowTableStyleColumnStripes" value="False" type="Boolean" /><column name="" property="ShowTableStyleFirstColumn" value="False" type="Boolean" /><column name="" property="ShowShowTableStyleLastColumn" value="False" type="Boolean" /><column name="" property="ShowTableStyleRowStripes" value="False" type="Boolean" /><column name="_RowNum" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="_RowNum" property="Address" value="$B$4" type="String" /><column name="_RowNum" property="NumberFormat" value="General" type="String" /><column name="_RowNum" property="VerticalAlignment" value="-4160" type="Double" /><column name="ID" property="EntireColumn.Hidden" value="True" type="Boolean" /><column name="ID" property="Address" value="$C$4" type="String" /><column name="ID" property="NumberFormat" value="General" type="String" /><column name="ID" property="VerticalAlignment" value="-4160" type="Double" /><column name="NAME" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="NAME" property="Address" value="$D$4" type="String" /><column name="NAME" property="ColumnWidth" value="30" type="Double" /><column name="NAME" property="NumberFormat" value="General" type="String" /><column name="NAME" property="VerticalAlignment" value="-4160" type="Double" /><column name="TEMPLATE" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TEMPLATE" property="Address" value="$E$4" type="String" /><column name="TEMPLATE" property="ColumnWidth" value="30" type="Double" /><column name="TEMPLATE" property="NumberFormat" value="General" type="String" /><column name="TEMPLATE" property="VerticalAlignment" value="-4160" type="Double" /><column name="DEFINITION" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="DEFINITION" property="Address" value="$F$4" type="String" /><column name="DEFINITION" property="ColumnWidth" value="70.71" type="Double" /><column name="DEFINITION" property="NumberFormat" value="General" type="String" /><column name="DEFINITION" property="VerticalAlignment" value="-4160" type="Double" /><column name="DEFINITION" property="WrapText" value="True" type="Boolean" /><column name="TABLE_SCHEMA" property="EntireColumn.Hidden" value="False" type="Boolean" /><column name="TABLE_SCHEMA" property="Address" value="$G$4" type="String" /><column name="TABLE_SCHEMA" property="ColumnWidth" value="16.57" type="Double" /><column name="TABLE_SCHEMA" property="NumberFormat" value="General" type="String" /><column name="TABLE_SCHEMA" property="VerticalAlignment" value="-4160" type="Double" /><column name="SortFields(1)" property="KeyfieldName" value="TABLE_SCHEMA" type="String" /><column name="SortFields(1)" property="SortOn" value="0" type="Double" /><column name="SortFields(1)" property="Order" value="1" type="Double" /><column name="SortFields(1)" property="DataOption" value="0" type="Double" /><column name="SortFields(2)" property="KeyfieldName" value="NAME" type="String" /><column name="SortFields(2)" property="SortOn" value="0" type="Double" /><column name="SortFields(2)" property="Order" value="1" type="Double" /><column name="SortFields(2)" property="DataOption" value="0" type="Double" /><column name="" property="ActiveWindow.DisplayGridlines" value="False" type="Boolean" /><column name="" property="ActiveWindow.FreezePanes" value="True" type="Boolean" /><column name="" property="ActiveWindow.Split" value="True" type="Boolean" /><column name="" property="ActiveWindow.SplitRow" value="0" type="Double" /><column name="" property="ActiveWindow.SplitColumn" value="-2" type="Double" /><column name="" property="PageSetup.Orientation" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesWide" value="1" type="Double" /><column name="" property="PageSetup.FitToPagesTall" value="1" type="Double" /></columnFormats></table>');
GO

INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'formats', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-formats.htm', N'', 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'handlers', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-handlers.htm', NULL, 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'objects', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-objects.htm', NULL, 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'queries', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-queries.htm', NULL, 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'translations', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-translations.htm', NULL, 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', NULL, N'Actions', N'gcrm', N'AppointmentItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', NULL, N'Actions', N'gcrm', N'ContactItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.contactitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', NULL, N'Actions', N'gcrm', N'JournalItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.journalitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', NULL, N'Actions', N'gcrm', N'MailItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', NULL, N'Actions', N'gcrm', N'MeetingItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', NULL, N'Actions', N'gcrm', N'PostItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.postitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', NULL, N'Actions', N'gcrm', N'NoteItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.noteitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'Actions', N'gcrm', N'TaskItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.taskitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'Actions', N'gcrm', N'TaskRequestAcceptItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.taskrequestacceptitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'Actions', N'gcrm', N'TaskRequestDeclineItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.taskrequestdeclineitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'Actions', N'gcrm', N'TaskRequestItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.taskrequestitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'Actions', N'gcrm', N'TaskRequestUpdateItem object', N'HTTP', N'https://docs.microsoft.com/en-us/office/vba/api/outlook.taskrequestupdateitem', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'workbooks', NULL, N'Actions', N'gcrm', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-workbooks.htm', NULL, 13, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', NULL, N'DoNotAddValidation', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'id', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'last_import_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'last_update_time', N'DoNotChange', NULL, NULL, NULL, NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'handlers', N'HANDLER_CODE', N'DoNotConvertFormulas', NULL, NULL, N'ATTRIBUTE', NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'gcrm', N'version', N'Information', NULL, NULL, N'ATTRIBUTE', N'2.0', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'handlers', N'EVENT_NAME', N'ValidationList', NULL, NULL, N'VALUES', N'Actions, AddHyperlinks, AddStateColumn, BitColumn, Change, ContextMenu, ConvertFormulas, DataTypeBit, DataTypeBoolean, DataTypeDate, DataTypeDateTime, DataTypeDateTimeOffset, DataTypeDouble, DataTypeInt, DataTypeGuid, DataTypeString, DataTypeTime, DefaultListObject, DefaultValue, DependsOn, DoNotAddChangeHandler, DoNotAddDependsOn, DoNotAddManyToMany, DoNotAddValidation, DoNotChange, DoNotConvertFormulas, DoNotKeepComments, DoNotKeepFormulas, DoNotSave, DoNotSelect, DoNotSort, DoNotTranslate, DoubleClick, DynamicColumns, Format, Formula, FormulaValue, Information, JsonForm, KeepFormulas, KeepComments, License, ManyToMany, ParameterValues, ProtectRows, RegEx, SelectionChange, SelectionList, SelectPeriod, SyncParameter, UpdateChangedCellsOnly, UpdateEntireRow, ValidationList', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'handlers', N'HANDLER_TYPE', N'ValidationList', NULL, NULL, N'VALUES', N'TABLE, VIEW, PROCEDURE, FUNCTION, CODE, HTTP, TEXT, MACRO, CMD, VALUES, RANGE, REFRESH, MENUSEPARATOR, PDF, REPORT, SHOWSHEETS, HIDESHEETS, SELECTSHEET, ATTRIBUTE', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'objects', N'PROCEDURE_TYPE', N'ValidationList', NULL, NULL, N'VALUES', N'TABLE, PROCEDURE, CODE, MERGE', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'objects', N'TABLE_TYPE', N'ValidationList', NULL, NULL, N'VALUES', N'TABLE, VIEW, PROCEDURE, FUNCTION, CODE, HTTP, TEXT, HIDDEN', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'BusyStatus', N'ValidationList', N'gcrm', N'usp_xl_list_busy_status_id', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_appointment_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 3', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_appointments', N'Importance', N'ValidationList', N'gcrm', N'usp_xl_list_importance_id', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_contacts', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_contact_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 1', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_journals', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_journal_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 5', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_mail_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 2', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_mails', N'Importance', N'ValidationList', N'gcrm', N'usp_xl_list_importance_id', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_notes', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_note_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 6', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'folder_id', N'ValidationList', N'gcrm', N'usp_xl_list_task_folder_id', N'CODE', N'EXEC gcrm.usp_xl_list_folder_id 4', NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'Importance', N'ValidationList', N'gcrm', N'usp_xl_list_importance_id', N'PROCEDURE', NULL, NULL, NULL, NULL);
INSERT INTO [gcrm].[handlers] ([TABLE_SCHEMA], [TABLE_NAME], [COLUMN_NAME], [EVENT_NAME], [HANDLER_SCHEMA], [HANDLER_NAME], [HANDLER_TYPE], [HANDLER_CODE], [TARGET_WORKSHEET], [MENU_ORDER], [EDIT_PARAMETERS]) VALUES (N'gcrm', N'view_tasks', N'Status', N'ValidationList', N'gcrm', N'usp_xl_list_task_status_id', N'PROCEDURE', NULL, NULL, NULL, NULL);
GO

INSERT INTO [gcrm].[importance_types] ([id], [name]) VALUES (2, N'High');
INSERT INTO [gcrm].[importance_types] ([id], [name]) VALUES (0, N'Low');
INSERT INTO [gcrm].[importance_types] ([id], [name]) VALUES (1, N'Normal');
GO

INSERT INTO [gcrm].[journal_recipient_types] ([id], [name]) VALUES (1, N'Associated contact');
GO

INSERT INTO [gcrm].[mail_recipient_types] ([id], [name]) VALUES (3, N'BCC');
INSERT INTO [gcrm].[mail_recipient_types] ([id], [name]) VALUES (2, N'CC');
INSERT INTO [gcrm].[mail_recipient_types] ([id], [name]) VALUES (0, N'Sender');
INSERT INTO [gcrm].[mail_recipient_types] ([id], [name]) VALUES (1, N'To');
GO

INSERT INTO [gcrm].[meeting_recipient_types] ([id], [name]) VALUES (2, N'Optional');
INSERT INTO [gcrm].[meeting_recipient_types] ([id], [name]) VALUES (0, N'Organizer');
INSERT INTO [gcrm].[meeting_recipient_types] ([id], [name]) VALUES (1, N'Required');
INSERT INTO [gcrm].[meeting_recipient_types] ([id], [name]) VALUES (3, N'Resource');
GO

INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_appointments', N'VIEW', N'', N'gcrm.usp_xl_appointments_insert', N'gcrm.usp_xl_appointments_update', N'gcrm.usp_xl_appointments_delete');
INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_contacts', N'VIEW', NULL, N'gcrm.usp_xl_contacts_insert', N'gcrm.usp_xl_contacts_update', N'gcrm.usp_xl_contacts_delete');
INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_journals', N'VIEW', NULL, N'gcrm.usp_xl_journals_insert', N'gcrm.usp_xl_journals_update', N'gcrm.usp_xl_journals_delete');
INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_mails', N'VIEW', NULL, N'gcrm.usp_xl_mails_insert', N'gcrm.usp_xl_mails_update', N'gcrm.usp_xl_mails_delete');
INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_notes', N'VIEW', NULL, N'gcrm.usp_xl_notes_insert', N'gcrm.usp_xl_notes_update', N'gcrm.usp_xl_notes_delete');
INSERT INTO [gcrm].[objects] ([TABLE_SCHEMA], [TABLE_NAME], [TABLE_TYPE], [TABLE_CODE], [INSERT_OBJECT], [UPDATE_OBJECT], [DELETE_OBJECT]) VALUES (N'gcrm', N'view_tasks', N'VIEW', NULL, N'gcrm.usp_xl_tasks_insert', N'gcrm.usp_xl_tasks_update', N'gcrm.usp_xl_tasks_delete');
GO

INSERT INTO [gcrm].[sensitivity_types] ([id], [name]) VALUES (3, N'Confidential');
INSERT INTO [gcrm].[sensitivity_types] ([id], [name]) VALUES (0, N'Normal');
INSERT INTO [gcrm].[sensitivity_types] ([id], [name]) VALUES (1, N'Personal');
INSERT INTO [gcrm].[sensitivity_types] ([id], [name]) VALUES (2, N'Private');
GO

INSERT INTO [gcrm].[task_recipient_types] ([id], [name]) VALUES (3, N'Final status');
INSERT INTO [gcrm].[task_recipient_types] ([id], [name]) VALUES (2, N'Update');
GO

INSERT INTO [gcrm].[task_statuses] ([id], [name]) VALUES (2, N'Complete');
INSERT INTO [gcrm].[task_statuses] ([id], [name]) VALUES (4, N'Deferred');
INSERT INTO [gcrm].[task_statuses] ([id], [name]) VALUES (1, N'InProgress');
INSERT INTO [gcrm].[task_statuses] ([id], [name]) VALUES (0, N'NotStarted');
INSERT INTO [gcrm].[task_statuses] ([id], [name]) VALUES (3, N'Waiting');
GO

INSERT INTO [gcrm].[workbooks] ([NAME], [TEMPLATE], [DEFINITION], [TABLE_SCHEMA]) VALUES (N'gcrm_configuration.xlsx', N'', N'objects=gcrm.objects,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":null,"TABLE_TYPE":null},"ListObjectName":"objects"}
handlers=gcrm.handlers,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":null,"EVENT_NAME":null,"HANDLER_TYPE":null},"ListObjectName":"handlers"}
translations=gcrm.translations,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":null,"LANGUAGE_NAME":null},"ListObjectName":"translations"}
workbooks=gcrm.workbooks,(Default),False,$B$3,,{"Parameters":{"TABLE_SCHEMA":null},"ListObjectName":"workbooks"}', N'gcrm');
GO

CREATE ROLE gcrm_users;
GO

GRANT SELECT ON gcrm.view_appointments  TO gcrm_users;
GO
GRANT SELECT ON gcrm.view_contacts      TO gcrm_users;
GO
GRANT SELECT ON gcrm.view_journals      TO gcrm_users;
GO
GRANT SELECT ON gcrm.view_mails         TO gcrm_users;
GO
GRANT SELECT ON gcrm.view_notes         TO gcrm_users;
GO
GRANT SELECT ON gcrm.view_tasks         TO gcrm_users;
GO

GRANT EXECUTE ON gcrm.usp_sync_appointments_select  TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_appointments_update  TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_contacts_select      TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_contacts_update      TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_journals_select      TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_journals_update      TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_mails_select         TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_mails_update         TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_notes_select         TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_notes_update         TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_tasks_select         TO gcrm_users;
GO
GRANT EXECUTE ON gcrm.usp_sync_tasks_update         TO gcrm_users;
GO

print 'Application installed';
