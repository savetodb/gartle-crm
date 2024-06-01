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

IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_appointments_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[appointments]'))
    ALTER TABLE [gcrm].[appointments] DROP CONSTRAINT [FK_appointments_folders];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_contacts_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[contacts]'))
    ALTER TABLE [gcrm].[contacts] DROP CONSTRAINT [FK_contacts_folders];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_folders_folder_types]') AND parent_object_id = OBJECT_ID(N'[gcrm].[folders]'))
    ALTER TABLE [gcrm].[folders] DROP CONSTRAINT [FK_folders_folder_types];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_folders_users]') AND parent_object_id = OBJECT_ID(N'[gcrm].[folders]'))
    ALTER TABLE [gcrm].[folders] DROP CONSTRAINT [FK_folders_users];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_journals_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[journals]'))
    ALTER TABLE [gcrm].[journals] DROP CONSTRAINT [FK_journals_folders];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_mails_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[mails]'))
    ALTER TABLE [gcrm].[mails] DROP CONSTRAINT [FK_mails_folders];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_notes_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[notes]'))
    ALTER TABLE [gcrm].[notes] DROP CONSTRAINT [FK_notes_folders];
GO
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[gcrm].[FK_tasks_folders]') AND parent_object_id = OBJECT_ID(N'[gcrm].[tasks]'))
    ALTER TABLE [gcrm].[tasks] DROP CONSTRAINT [FK_tasks_folders];
GO

IF OBJECT_ID('[gcrm].[usp_get_folder_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_get_folder_id];
GO
IF OBJECT_ID('[gcrm].[usp_sync_appointments_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_appointments_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_appointments_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_appointments_update];
GO
IF OBJECT_ID('[gcrm].[usp_sync_contacts_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_contacts_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_contacts_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_contacts_update];
GO
IF OBJECT_ID('[gcrm].[usp_sync_journals_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_journals_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_journals_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_journals_update];
GO
IF OBJECT_ID('[gcrm].[usp_sync_mails_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_mails_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_mails_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_mails_update];
GO
IF OBJECT_ID('[gcrm].[usp_sync_notes_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_notes_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_notes_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_notes_update];
GO
IF OBJECT_ID('[gcrm].[usp_sync_tasks_select]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_tasks_select];
GO
IF OBJECT_ID('[gcrm].[usp_sync_tasks_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_sync_tasks_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_appointments_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_appointments_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_appointments_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_appointments_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_appointments_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_appointments_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_contacts_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_contacts_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_contacts_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_contacts_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_contacts_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_contacts_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_journals_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_journals_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_journals_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_journals_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_journals_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_journals_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_busy_status_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_busy_status_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_folder_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_folder_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_importance_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_importance_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_journal_recipient_type_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_journal_recipient_type_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_mail_recipient_type_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_mail_recipient_type_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_meeting_recipient_type_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_meeting_recipient_type_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_sensitivity_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_sensitivity_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_task_recipient_type_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_task_recipient_type_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_list_task_status_id]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_list_task_status_id];
GO
IF OBJECT_ID('[gcrm].[usp_xl_mails_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_mails_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_mails_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_mails_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_mails_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_mails_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_notes_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_notes_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_notes_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_notes_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_notes_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_notes_update];
GO
IF OBJECT_ID('[gcrm].[usp_xl_tasks_delete]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_tasks_delete];
GO
IF OBJECT_ID('[gcrm].[usp_xl_tasks_insert]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_tasks_insert];
GO
IF OBJECT_ID('[gcrm].[usp_xl_tasks_update]', 'P') IS NOT NULL
DROP PROCEDURE [gcrm].[usp_xl_tasks_update];
GO

IF OBJECT_ID('[gcrm].[queries]', 'V') IS NOT NULL
DROP VIEW [gcrm].[queries];
GO
IF OBJECT_ID('[gcrm].[view_appointments]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_appointments];
GO
IF OBJECT_ID('[gcrm].[view_contacts]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_contacts];
GO
IF OBJECT_ID('[gcrm].[view_journals]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_journals];
GO
IF OBJECT_ID('[gcrm].[view_mails]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_mails];
GO
IF OBJECT_ID('[gcrm].[view_notes]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_notes];
GO
IF OBJECT_ID('[gcrm].[view_outlook_columns]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_outlook_columns];
GO
IF OBJECT_ID('[gcrm].[view_outlook_parameters]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_outlook_parameters];
GO
IF OBJECT_ID('[gcrm].[view_tasks]', 'V') IS NOT NULL
DROP VIEW [gcrm].[view_tasks];
GO

IF OBJECT_ID('[gcrm].[appointments]', 'U') IS NOT NULL
DROP TABLE [gcrm].[appointments];
GO
IF OBJECT_ID('[gcrm].[busy_statuses]', 'U') IS NOT NULL
DROP TABLE [gcrm].[busy_statuses];
GO
IF OBJECT_ID('[gcrm].[contacts]', 'U') IS NOT NULL
DROP TABLE [gcrm].[contacts];
GO
IF OBJECT_ID('[gcrm].[folder_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[folder_types];
GO
IF OBJECT_ID('[gcrm].[folders]', 'U') IS NOT NULL
DROP TABLE [gcrm].[folders];
GO
IF OBJECT_ID('[gcrm].[formats]', 'U') IS NOT NULL
DROP TABLE [gcrm].[formats];
GO
IF OBJECT_ID('[gcrm].[handlers]', 'U') IS NOT NULL
DROP TABLE [gcrm].[handlers];
GO
IF OBJECT_ID('[gcrm].[importance_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[importance_types];
GO
IF OBJECT_ID('[gcrm].[journal_recipient_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[journal_recipient_types];
GO
IF OBJECT_ID('[gcrm].[journals]', 'U') IS NOT NULL
DROP TABLE [gcrm].[journals];
GO
IF OBJECT_ID('[gcrm].[mail_recipient_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[mail_recipient_types];
GO
IF OBJECT_ID('[gcrm].[mails]', 'U') IS NOT NULL
DROP TABLE [gcrm].[mails];
GO
IF OBJECT_ID('[gcrm].[meeting_recipient_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[meeting_recipient_types];
GO
IF OBJECT_ID('[gcrm].[notes]', 'U') IS NOT NULL
DROP TABLE [gcrm].[notes];
GO
IF OBJECT_ID('[gcrm].[objects]', 'U') IS NOT NULL
DROP TABLE [gcrm].[objects];
GO
IF OBJECT_ID('[gcrm].[sensitivity_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[sensitivity_types];
GO
IF OBJECT_ID('[gcrm].[task_recipient_types]', 'U') IS NOT NULL
DROP TABLE [gcrm].[task_recipient_types];
GO
IF OBJECT_ID('[gcrm].[task_statuses]', 'U') IS NOT NULL
DROP TABLE [gcrm].[task_statuses];
GO
IF OBJECT_ID('[gcrm].[tasks]', 'U') IS NOT NULL
DROP TABLE [gcrm].[tasks];
GO
IF OBJECT_ID('[gcrm].[translations]', 'U') IS NOT NULL
DROP TABLE [gcrm].[translations];
GO
IF OBJECT_ID('[gcrm].[users]', 'U') IS NOT NULL
DROP TABLE [gcrm].[users];
GO
IF OBJECT_ID('[gcrm].[workbooks]', 'U') IS NOT NULL
DROP TABLE [gcrm].[workbooks];
GO


DECLARE @sql nvarchar(max) = ''

SELECT
    @sql = @sql + 'ALTER ROLE ' + QUOTENAME(r.name) + ' DROP MEMBER ' + QUOTENAME(m.name) + ';' + CHAR(13) + CHAR(10)
FROM
    sys.database_role_members rm
    INNER JOIN sys.database_principals r ON r.principal_id = rm.role_principal_id
    INNER JOIN sys.database_principals m ON m.principal_id = rm.member_principal_id
WHERE
    r.name IN ('gcrm_users')

IF LEN(@sql) > 1
    BEGIN
    EXEC (@sql);
    PRINT @sql
    END
GO

IF DATABASE_PRINCIPAL_ID('gcrm_users') IS NOT NULL
DROP ROLE [gcrm_users];
GO

IF SCHEMA_ID('gcrm') IS NOT NULL
DROP SCHEMA [gcrm];
GO


print 'Application removed';
