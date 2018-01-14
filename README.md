# Remove-DuplicateItems

## Getting Started

This script will scan each folder of a given primary mailbox and personal archive (when
configured, Exchange 2010 and later) and removes duplicate items per folder. You can specify
how the items should be deleted and what items to process, e.g. mail items or appointments.
Sample scenarios are misbehaving 3rd party synchronization tool creates duplicate items or
(accidental) import of PST file with duplicate items. 

Script will process mailbox and archive if configured, unless MailboxOnly or ArchiveOnly 
is specified. For Exchange 2007, you need to specify -MailboxOnly.

### Requirements

* PowerShell 3.0 or later
* EWS Managed API 1.2 or later

### Usage

Syntax:
Remove-DuplicateItems.ps1 [[-Identity] <String>] [[-Type] <String>] [-Retain <String>] [-Server <String>] [-Impersonation] [-DeleteMode <String>] [-Credentials <PSCredential>] [-Mode <String>] [-MailboxOnly] [-ArchiveOnly] [-IncludeFolders <String[]>] [-ExcludeFolders <String[]>] [-PriorityFolders <String[]>] [-MailboxWide] [-NoSize] [-NoProgressBar] [Report] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]

Examples:
```
$Credentials= Get-Credential
.\Remove-DuplicateItems.ps1 -Mailbox olrik@office365tenant.com -Credentials $Credentials
```
Sets $Credentials variable. Then, check olrik@office365tenant.com's mailbox for duplicate items in each folder, using
Credentials provided earlier.

```
$Credentials= Get-Credential
.\Remove-DuplicateItems.ps1 -Mailbox olrik@office365tenant.com -Server outlook.office365.com -Credentials $Credentials -IncludeFolders '#Inbox#\*','\Projects\*' -ExcludeFolders 'Keep Out' -PriorityFolders '*Important*' -MailboxWide
```
Remove duplicate items from specified mailbox in Office365 using fixed Server FQDN - bypassing AutoDiscover, limiting
operation against the Inbox, and top Projects folder, and all of their subfolders, but excluding any folder named Keep Out.
Duplicates are checked over all folders, but priority is given to folders containing the word Important, causing items in
those folders to be kept over items in other folders when duplicates are found.

### About

For more information on this script, as well as usage and examples, see
the related blog article, [Removing Duplicate Items from a Mailbox](http://eightwone.com/2013/06/21/removing-duplicate-items-from-a-mailbox/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 