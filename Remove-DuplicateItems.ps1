<#
    .SYNOPSIS
    Remove-DuplicateItems

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 2.21, May 31st, 2022

    Acknowledgements: Rob Gray (PF support)

    .DESCRIPTION
    This script will scan each folder of a given primary mailbox and personal archive (when
    configured, Exchange 2010 and later) and removes duplicate items per folder. You can specify
    how the items should be deleted and what items to process, e.g. mail items or appointments.
    Sample scenarios are misbehaving 3rd party synchronization tool creates duplicate items or
    (accidental) import of PST file with duplicate items. Script will process
    mailbox and archive if configured, unless MailboxOnly or ArchiveOnly is specified.

    Note that usage of the Verbose, Confirm and WhatIf parameters is supported.
    When using Confirm, you will be prompted per batch.

    .LINK
    http://eightwone.com

    .NOTES
    EWS.WebServices.Managed.Api 2.2 or up is required (see https://eightwone.com/2020/10/05/ews-webservices-managed-api)
    For OAuth, Microsoft Authentication are required.
    
    Search order for DLL's is script Folder then installed packages.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
    1.1     Fixed issue with PowerShell v3 (System.Collections.Generic.List`1)
            Specified mailbox will also match object using mail attribute
    1.2     Added requested Retain option (default Newest, was undetermined)
    1.21    Switched Retain from using DateTimeReceived to LastModifiedTime
    1.3     Changed parameter Mailbox, you can now use an e-mail address as well
            Added parameter Credentials
            Added item class and size for certain  duplication checks
            Changed item removal process. Remove items after, not while processing
            folder. Avoids asynchronous deletion issues.
            Works against Office 365
    1.4     Added personal archive support
    1.41    Fixed typo preventing script from working on Ex2007
    1.5     Prevents sending cancellation notices when removing calendar items
    1.6     Added IncludeFolder parameter
            Added ExcludeFolder parameter
            Added MD5 hashing of keys to lower memory usage
            Added MailboxWide switch (CAUTION)
    1.61    Fixed impersonation logic issue
    1.62    Fixed using 2+ Exclude folders
    1.63    Identity parameter replaces Mailbox
            Made "can't access information store" more verbose.
            Fixed bug in non-wildcard matching
    1.7     Changed IncludeFolders and ExcludeFolders to add path matching
            Added PriorityFolders
            Added #JunkEmail# and #DeletedItems# to IncludeFolders/ExcludeFolders
            Added NoSize switch
    1.8     Added EWS throttling handling
            Added progress bar
            Added NoProgressBar switch
            Fixed Well-Known Folder processing to use current mailbox folder name
            Added some statistics, e.g. items/minute summary
            Optimalizations when running against multiple mailboxes
            Changed folder /* matching to include folder and subfolders, not subfolders only
            Some code rewriting
    1.81    Fixed partial folder name matching
            Fixed progress bare issue when no. items to remove exceeds removal batch size
    1.82    Fixed progress bar cleanup
            Added notice of mailbox/archive processing
    1.83    Added Report switch
            Added EWS Managed API DLL version reporting (Verbose)
    1.84    Added X-AnchorMailbox for impersonation requests
    1.85    Added Body option for Mode
    1.86    Fixed issue with processing delegate mailboxes using Full Access permissions
    1.87    Fixed Examples
    1.88    Fixed bug in folder selection process
    1.89    Added code to leverage installed package EWS.WebServices.Managed.Api 
    2.00    Added OAuth authentication options
            Changed DLL loading routing (EWS Managed API + MSAL)
            Not trusting self-signed certs by default; added -TrustAll switch to trust all certs
            Added pipeline proper processing with begin/process/end
            Replaced all strings with var-subsitution with -f 
            Added certificate authentication example
            Small performance tweaks here and there
    2.01    Fixed verification of loading Microsoft.Identity.Client
    2.02    Determine DeletedItems once per mailbox, not for every folder to process
    2.03    Fixed accepting multiple Identity entries
            Added CleanupMode parameter
            Removed MailboxWide switch
    2.04    Fixed loading of module when using installed NuGet packages
    2.05    Changed PropertySet constructors to prevent possible initialization issues
    2.06    Fixed parenthesis omission when running Verbose
    2.07    Fixed handling MoveToDelete for archive mailbox
    2.10    Added UseDefaultCredentials for usage on-premises (using current security context)
    2.11    Changed class to check proper loading of Microsoft.Identity.Client module
    2.12    Changed class to check proper loading of Microsoft.Identity.Client module in PS7 with latest module
    2.20    Added Public Folder support
            Refactoring to accomodate PF support
            Requires PowerShell 3 and up (removed <PF3 compatibility code)
            Removed Exchange Server 2007 support
    2.21    Fixed display 'Mailbox' when processing Public Folders

    .PARAMETER Identity
    Identity of the Mailbox. Can be CN/SAMAccountName (for on-premises) or e-mail format (on-prem & Office 365)

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited,
    script will attempt to use Autodiscover.

    .PARAMETER Impersonation
    When specified, uses impersonation when accessing the mailbox, otherwise account specified with Credentials is
    used. When using OAuth authentication with a registered app, you don't need to specify Impersonation.
    For details on how to configure impersonation access for Exchange 2010 using RBAC, see this article:
    https://eightwone.com/2014/08/13/application-impersonation-to-be-or-pretend-to-be/

    .PARAMETER Retain
    Determines which matching items are kept, per folder (based on Last Modification Time):
    - Oldest:             Oldest received item is kept, newest item(s) are deleted
    - Newest:             Newest received item is kept, oldest item(s) are deleted (default)

    .PARAMETER DeleteMode
    Determines how to remove messages. Options are:
    - HardDelete:         Items will be permanently deleted.
    - SoftDelete:         Items will be moved to the dumpster (default).
    - MoveToDeletedItems: Items will be moved to the Deleted Items folder.

    When using MoveToDeletedItems, the Deleted Items folder will not be processed.

    .PARAMETER Type
    Determines what kind of folders to check for duplicates.
    Options: Mail, Calendar, Contacts, Tasks, Notes or All (Default).

    .PARAMETER Mode
    Determines how items are matched. Options are:
    - Quick:  Removes duplicate items with matching PidTagSearchKey
              attribute; This is the default mode.
    - Full:   Removes duplicate items with predefined matching criteria,
              depending on item class:
              - Contacts: File As, First Name, Last Name, Company Name,
                Business Phone, Mobile Phone, Home Phone, Size
              - Distribution List: FileAs, Number of Members, Size
              - Calendar: Subject, Location, Start & End Date, Size
              - Task: Subject, Start Date, Due Date, Status, Size
              - Note: Contents, Color, Size
              - Mail: Subject, Internet Message ID, DateTimeSent,
                DateTimeReceived, Sender, Size
              - Other: Subject, DateTimeReceived, Size
    - Body:   Removes duplicate items with matching Body attribute.

    When NoSize is used in Full mode, Size is not used as criteria.

    Note that when Quick mode is used and PidTagSearchKey is missing or
    inaccessible, search will fall back to Full mode. For more info on
    PidTagSearchKey: http://msdn.microsoft.com/en-us/library/cc815908.aspx

    .PARAMETER MailboxOnly
    Only process primary mailbox of specified users. You als need to use this parameter when
    running against mailboxes on Exchange Server 2007.

    .PARAMETER ArchiveOnly
    Only process personal archives of specified users.

    .PARAMETER PublicFolders
    Switch to indicate that (modern) Public Folders need to be processed instead
    of mailboxes or archives.

    .PARAMETER IncludeFolders
    Specify one or more names of folder(s) to include, e.g. 'Projects'. You can use wildcards
    around or at the end to include folders containing or starting with this string, e.g.
    'Projects*' or '*Project*'. To match folders and subfolders, add a trailing \*,
    e.g. Projects\*. This will include folders named Projects and all subfolders.
    To match from the top of the structure, prepend using '\'. Matching is case-insensitive.

    Some examples, using the following folder structure:

    + TopFolderA
        + FolderA
            + SubFolderA
            + SubFolderB
        + FolderB
    + TopFolderB
        + FolderA

    Filter              Match(es)
    --------------------------------------------------------------------------------------------------------------------
    FolderA             \TopFolderA\FolderA, \TopFolderB\FolderA
    Folder*             \TopFolderA\FolderA, \TopFolderA\FolderB, \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB
    FolderA\*Folder*    \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB
    \*FolderA\*         \TopFolderA, \TopFolderA\FolderA, \TopFolderA\FolderB, \TopFolderA\FolderA\SubFolderA, \TopFolderA\FolderA\SubFolderB, \TopFolderB\FolderA
    \*\FolderA          \TopFolderA\FolderA, \TopFolderB\FolderA

    For mailbox processing, you can also use well-known folders, by using this format: #WellKnownFolderName#, 
    e.g. #Inbox#. Supported are #Calendar#, #Contacts#, #Inbox#, #Notes#, #SentItems#, #Tasks#, #JunkEmail# 
    and #DeletedItems#. The script uses the currently configured Well-Known Folder of the mailbox to be processed.

    .PARAMETER ExcludeFolders
    Specify one or more folder(s) to exclude. Usage of wildcards and well-known folders identical to IncludeFolders.
    Note that ExcludeFolders criteria overrule IncludeFolders when matching folders.

    .PARAMETER Force
    Force removal of items without prompting.

    .PARAMETER CleanupMode
    Options are:
    - Folder (default) - performs duplicate cleanup per-folder comparison of mailboxes/public folders.
    - Mailbox - performs duplicate cleanup against whole mailbox or public folders, instead of per folder.
      By default, the first unique item encountered will be retained. When an item is found in Folder A and
      in Folder B, it is undetermined which item will be kept, unless PriorityFolders is used.
    - MultiMailbox - When passing multiple identities, performs duplicate cleanup over multiple mailboxes. Items 
      are evaluated sequentially, e.g. items found in the first mailbox are considered duplicate when they are located
      in the second or later mailboxes. 

    .PARAMETER PriorityFolders
    Determines which folders have priority over other folders, identifying items in these folders first when
    using MailboxWide mode. Usage of wildcards and well-known folders is identical to IncludeFolders.

    .PARAMETER NoSize
    Don't use size to match items in Full mode.

    .PARAMETER Report
    Reports individual items detected as duplicate. Can be used together with WhatIf to perform pre-analysis.

    .PARAMETER NoProgressBar
    Use this switch to prevent displaying a progress bar as folders and items are being processed.

    .PARAMETER TrustAll
    Specifies if all certificates should be accepted, including self-signed certificates.

    .PARAMETER TenantId
    Specifies the identity of the Tenant.

    .PARAMETER ClientId
    Specifies the identity of the application configured in Azure Active Directory.

    .PARAMETER UseDefaultCredentials
    Instruct script to use current security context, for example, for usage against Exchange on-premises.

    .PARAMETER Credentials
    Specify credentials to use with Basic Authentication. Credentials can be set using $Credentials= Get-Credential
    This parameter is mutually exclusive with CertificateFile, CertificateThumbprint and Secret. 

    .PARAMETER CertificateThumbprint
    Specify the thumbprint of the certificate to use with OAuth authentication. The certificate needs
    to reside in the personal store. When using OAuth, providing TenantId and ClientId is mandatory.
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificateFile
    Specify the .pfx file containing the certificate to use with OAuth authentication. When a password is required,
    you will be prompted or you can provide it using CertificatePassword.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificatePassword
    Sets the password to use with the specified .pfx file. The provided password needs to be a secure string, 
    eg. -CertificatePassword (ConvertToSecureString -String 'P@ssword' -Force -AsPlainText)

    .PARAMETER Secret
    Specifies the client secret to use with OAuth authentication. The secret needs to be provided as a secure string.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and CertificateThumbprint. 

    .EXAMPLE
    .\Remove-DuplicateItems.ps1 -Identity Francis -Type All -Impersonation -DeleteMode SoftDelete -Mode Quick -Verbose

    Check Francis' mailbox for duplicate items in each folder, soft deleting
    duplicates, matching on PidTagSearchKey and using impersonation.

    .EXAMPLE
    .\Remove-DuplicateItems.ps1 -Identity Philip -Retain Oldest -Type Mail -Impersonation -DeleteMode MoveToDeletedItems -Mode Full -Verbose

    Check Philip's mailbox for duplicate task items in each folder and moves
    duplicates to the Deleted Items folder, using preset matching criteria
    and impersonation. When duplicates are found, the oldest is retained.

    .EXAMPLE
    $Credentials= Get-Credential
    .\Remove-DuplicateItems.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials

    Sets $Credentials variable. Then, check olrik@office365tenant.com's mailbox for duplicate items in each folder, using
    Credentials provided earlier for Basic Authentication.

    .EXAMPLE
    $Secret= Read-Host 'Secret' -AsSecureString
    Import-Csv Users.Csv | .\Remove-DuplicateItems.ps1 -Server outlook.office365.com -IncludeFolders '#Inbox#\*','\Projects\*' -ExcludeFolders 'Keep Out' -PriorityFolders '*Important*' -CleanupMode Mailbox --TenantId '1ab81a53-2c16-4f28-98f3-fd251f0459f3' -ClientId 'ea76025c-592d-43f1-91f4-2dec7161cc59' -Secret $Secret

    Remove duplicate items from mailboxes identified by CSV file in Office365 bypassing AutoDiscover, limiting operation against the Inbox, and 
    top Projects folder, and all of their subfolders, but excluding any folder named Keep Out. Duplicates are checked over all folders, but priority is
    given to folders containing the word Important, causing items in those folders to be kept over items in other folders when duplicates are found.
    OAuth authentication is performed against indicated tenant <TenantID> using registered App <ClientID> and App secret entered.
#>
#Requires -Version 3
[cmdletbinding(
    DefaultParameterSetName = 'DefaultAuth',
    SupportsShouldProcess= $true,
    ConfirmImpact= 'High'
)]
param(
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuth')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuth')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuthPublicFolders')] 
    [alias('Mailbox')]
    [string[]]$Identity,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [ValidateSet( 'Mail', 'Calendar', 'Contacts', 'Tasks', 'Notes', 'All')]
    [string]$Type= 'All',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [ValidateSet( 'Oldest', 'Newest')]
    [string]$Retain= 'Newest',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [string]$Server,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$Impersonation,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [ValidateSet( 'HardDelete', 'SoftDelete', 'MoveToDeletedItems')]
    [string]$DeleteMode= 'SoftDelete',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [ValidateSet( 'Quick', 'Full', 'Body')]
    [string]$Mode= 'Quick',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [switch]$MailboxOnly,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [switch]$ArchiveOnly,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$PublicFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')]
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [string[]]$IncludeFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [string[]]$ExcludeFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')]
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [ValidateSet( 'Folder', 'Mailbox', 'MultiMailbox')]
    [string]$CleanupMode= 'Folder',
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [string[]]$PriorityFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$NoSize,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$Force,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$NoProgressBar,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$Report,
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [Switch]$UseDefaultCredentials,
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuthPublicFolders')] 
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [System.Security.SecureString]$Secret,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [String]$CertificateThumbprint,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$CertificateFile,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [System.Security.SecureString]$CertificatePassword,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [string]$TenantId,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [string]$ClientId,
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'DefaultAuthPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthMailboxOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFileArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthArchiveOnly')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumbPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFilePublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecretPublicFolders')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuthPublicFolders')] 
    [switch]$TrustAll
)
#Requires -Version 3.0

begin {

    # Process folders these batches
    $MaxFolderBatchSize= 100
    # Process items in these page sizes
    $MaxItemBatchSize= 100
    # Max of concurrent item deletes
    $MaxDeleteBatchSize= 100

    # Initial sleep timer (ms) and treshold before lowering
    $script:SleepTimerMax= 300000               # Maximum delay (5min)
    $script:SleepTimerMin= 100                  # Minimum delay
    $script:SleepAdjustmentFactor= 2.0          # When tuning, use this factor
    $script:SleepTimer= $script:SleepTimerMin   # Initial sleep timer value

    # Error codes
    $ERR_DLLNOTFOUND= 1000
    $ERR_DLLLOADING= 1001
    $ERR_MAILBOXNOTFOUND= 1002
    $ERR_AUTODISCOVERFAILED= 1003
    $ERR_CANTACCESSMAILBOXSTORE= 1004
    $ERR_PROCESSINGMAILBOX= 1005
    $ERR_PROCESSINGARCHIVE= 1006
    $ERR_INVALIDCREDENTIALS= 1007
    $ERR_PROBLEMIMPORTINGCERT= 1008
    $ERR_CERTNOTFOUND= 1009
    $ERR_PROCESSINGPUBLICFOLDERS= 1010
    $ERR_CANTACCESSPUBLICFOLDERS= 1011

    # Initialize list to keep track of unique items
    $global:UniqueList= [System.Collections.ArrayList]@()

    Function Import-ModuleDLL {
        param(
            [string]$Name,
            [string]$FileName,
            [string]$Package
        )

        $AbsoluteFileName= Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If ( Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
            If( $Package) {
                If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If( Get-Package -Name $Package -ErrorAction SilentlyContinue) {
                        $AbsoluteFileName= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package -Name $Package | Sort-Object -Property Version -Descending | Select-Object -First 1).Source) -Filter $FileName -Recurse).FullName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
            If( $ModLoaded) {
                Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            }
            Else {
                Write-Verbose ('Loading module {0}' -f $absoluteFileName)
                try {
                    Import-Module -Name $absoluteFileName -Global -Force
                    Start-Sleep 1
                }
                catch {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
                $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
                If( $ModLoaded) {
                    Write-Verbose ('Module {0} v{1} loaded' -f $ModLoaded.Name, $ModLoaded.Version)
                }
                If(!( Get-Module -Name $Name -ErrorAction SilentlyContinue)) {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $_.Exception.Message)
                    Exit $ERR_DLLLOADING
                }
            }
        }
        Else {
            Write-Verbose ('Required module {0} could not be located' -f $FileName)
            Exit $ERR_DLLNOTFOUND
        }
    }

    Function Set-SSLVerification {
        param(
            [switch]$Enable,
            [switch]$Disable
        )

        Add-Type -TypeDefinition  @"
            using System.Net.Security;
            using System.Security.Cryptography.X509Certificates;
            public static class TrustEverything
            {
                private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors sslPolicyErrors) { return true; }
                public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= ValidationCallback; }
                public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= null; }
        }
"@
        If($Enable) {
            Write-Verbose ('Enabling SSL certificate verification')
            [TrustEverything]::UnsetCallback()
        }
        Else {
            Write-Verbose ('Disabling SSL certificate verification')
            [TrustEverything]::SetCallback()
        }
    }

    Function iif( $eval, $tv= '', $fv= '') {
        If ( $eval) { return $tv } else { return $fv}
    }

    Function Get-EmailAddress {
        param(
            [string]$Identity
        )
        $address= [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if ( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as SamAccountName
            $ADSearch= New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter= "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result= $ADSearch.FindOne()
            If ( $Result) {
                $objUser= $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function Construct-FolderFilter {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string[]]$Folders,
            [string]$emailAddress
        )
        If ( $Folders) {
            $FolderFilterSet= [System.Collections.ArrayList]@()
            ForEach ( $Folder in $Folders) {
                # Convert simple filter to (simple) regexp
                $Parts= $Folder -match '^(?<root>\\)?(?<keywords>.*?)?(?<sub>\\\*)?$'
                If ( !$Parts) {
                    Write-Error ('Invalid regular expression matching against {0}' -f $Folder)
                }
                Else {
                    $Keywords= Search-ReplaceWellKnownFolderNames -EwsService $EwsService -Criteria ($Matches.keywords) -EmailAddress $emailAddress
                    $EscKeywords= [Regex]::Escape( $Keywords) -replace '\\\*', '.*'
                    $Pattern= iif -eval $Matches.Root -tv '^\\' -fv '^\\(.*\\)*'
                    $Pattern += iif -eval $EscKeywords -tv $EscKeywords -fv ''
                    $Pattern += iif -eval $Matches.sub -tv '(\\.*)?$' -fv '$'
                    $Obj= [pscustomobject]@{
                        'Pattern'    = [string]$Pattern
                        'IncludeSubs'= [bool]$Matches.Sub
                        'OrigFilter' = [string]$Folder
                    }
                    $FolderFilterSet.Add( $Obj) | Out-Null
                    Write-Debug ($Obj -join ',')
                }
            }
        }
        Else {
            $FolderFilterSet= $null
        }
        return $FolderFilterSet
    }

    Function Search-ReplaceWellKnownFolderNames {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$criteria= '',
            [string]$emailAddress
        )
        $AllowedWKF= 'Inbox', 'Calendar', 'Contacts', 'Notes', 'SentItems', 'Tasks', 'JunkEmail', 'DeletedItems'
        # Construct regexp to see if allowed WKF is part of criteria string
        ForEach ( $ThisWKF in $AllowedWKF) {
            If ( $criteria -match '#{0}#') {
                $criteria= $criteria -replace ('#{0}#' -f $ThisWKF), (myEWSBind-WellKnownFolder $EwsService $ThisWKF $emailAddress).DisplayName
            }
        }
        return $criteria
    }
    Function Tune-SleepTimer {
        param(
            [bool]$previousResultSuccess= $false
        )
        if ( $previousResultSuccess) {
            If ( $script:SleepTimer -gt $script:SleepTimerMin) {
                $script:SleepTimer= [int]([math]::Max( [int]($script:SleepTimer / $script:SleepAdjustmentFactor), $script:SleepTimerMin))
                Write-Warning ('Previous EWS operation successful, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
            }
        }
        Else {
            $script:SleepTimer= [int]([math]::Min( ($script:SleepTimer * $script:SleepAdjustmentFactor) + 100, $script:SleepTimerMax))
            If ( $script:SleepTimer -eq 0) {
                $script:SleepTimer= 5000
            }
            Write-Warning ('Previous EWS operation failed, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
        }
        Start-Sleep -Milliseconds $script:SleepTimer
    }

    Function myEWSFind-Folders {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
            [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$FolderSearchCollection,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $FolderId, $FolderSearchCollection, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders with Search options in {0}. Error: {1}' -f $FolderId.FolderName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-FoldersNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $FolderId, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders without Search options in {0}. Error: {1}' -f $FolderId.FolderName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-Items {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$ItemSearchFilterCollection,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $Folder.FindItems( $ItemSearchFilterCollection, $ItemView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindItems with Search options in {0}. Error: {1}' -f $Folder.DisplayName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-ItemsNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $Folder.FindItems( $ItemView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindItems without Search options in {0}. Error {1}' -f $Folder.DisplayName, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSRemove-Items {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            $ItemIds,
            [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteMode,
            [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]$SendCancellationsMode,
            [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]$AffectedTaskOccurrences,
            [bool]$SuppressReadReceipt
        )
        $OpSuccess= $false
        $critErr= $false
        Do {
            Try {
                If ( @([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013, [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1) -contains $EwsService.RequestedServerVersion) {
                    $res= $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences, $SuppressReadReceipt)
                }
                Else {
                    $res= $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences)
                }
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation RemoveItems with {0}. Error: {1}' -f $RemoveItems, $Error[0])
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSBind-WellKnownFolder {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$WellKnownFolderName,
            [string]$emailAddress
        )
        $OpSuccess= $false
        $critErr= $false
        Do {
            Try {
                If( $emailAddress) {
                    $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $emailAddress)  
                }
                Else {
                    $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName)  
                }
                $res= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $explicitFolder)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Cannot bind to {0}: {1}' -f $WellKnownFolderName, $_.Exception.Message)
            }
            finally {
                If ( !$critErr) { Tune-SleepTimer $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function Get-Hash {
        param(
            [string]$string
        )
        $md5= New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
        $data= ([system.Text.Encoding]::UTF8).GetBytes( $string)
        return ([System.BitConverter]::ToString( $md5.ComputeHash( $data)) -replace ' - ', '')
    }

    Function Get-FolderPriority {
        param(
            $FolderPath,
            $PriorityFilter
        )
        $prio= 0
        If ( $PriorityFilter) {
            $num= 0
            ForEach ( $Filter in $PriorityFilter) {
                $num++
                If ( $FolderPath -match $Filter.Pattern) {
                    $prio= $num
                }
            }
        }
        return $prio
    }

    Function Get-SubFolders {
        param(
            $Folder,
            $CurrentPath,
            $IncludeFilter,
            $ExcludeFilter,
            $PriorityFilter,
            $EwsService
        )
        $FoldersToProcess= [System.Collections.ArrayList]@()
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( $MaxFolderBatchSize)
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $FolderView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
        $FolderView.PropertySet.Add( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass)
        $FolderSearchCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection( [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        If ( $Type -ne 'All') {
            $FolderSearchClass= (@{Mail= 'IPF.Note'; Calendar= 'IPF.Appointment'; Contacts= 'IPF.Contact'; Tasks= 'IPF.Task'; Notes= 'IPF.StickyNotes'})[$Type]
            $FolderSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, $FolderSearchClass)
            $FolderSearchCollection.Add( $FolderSearchFilter)
        }
        Do {
            If ( $FolderSearchCollection.Count -ge 1) {
                $FolderSearchResults= myEWSFind-Folders $EwsService $Folder.Id $FolderSearchCollection $FolderView
            }
            Else {
                $FolderSearchResults= myEWSFind-FoldersNoSearch $EwsService $Folder.Id $FolderView
            }
            ForEach ( $FolderItem in $FolderSearchResults) {
                $FolderPath= '{0}\{1}' -f $CurrentPath, $FolderItem.DisplayName
                If ( $IncludeFilter) {
                    $Add= $false
                    # Defaults to true, unless include does not specifically include subfolders
                    $Subs= $true
                    ForEach ( $Filter in $IncludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add= $true
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs= $Filter.IncludeSubs
                        }
                    }
                }
                Else {
                    # If no includeFolders specified, include all (unless excluded)
                    $Add= $true
                    $Subs= $true
                }
                If ( $ExcludeFilter) {
                    # Excludes can overrule includes
                    ForEach ( $Filter in $ExcludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add= $false
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs= $Filter.IncludeSubs
                        }
                    }
                }
                If ( $Add) {
                    $Prio= Get-FolderPriority $FolderPath -PriorityFilter $PriorityFilter
                    Write-Verbose ( 'Adding folder {0} (priority {1})' -f $FolderPath, $Prio)

                    $Obj= New-Object -TypeName PSObject -Property @{
                        'Name'    = $FolderPath;
                        'Priority'= $Prio;
                        'Folder'  = $FolderItem
                    }
                    $FoldersToProcess.Add( $Obj) | Out-Null
                }
                If ( $Subs) {
                    # Could be that specific folder is to be excluded, but subfolders needs evaluation
                    ForEach ( $AddFolder in (Get-SubFolders -Folder $FolderItem -CurrentPath $FolderPath -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService)) {
                        $FoldersToProcess.Add( $AddFolder) | Out-Null
                    }
                }
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)
        Write-Output -NoEnumerate $FoldersToProcess
    }

    Function Process-Mailbox {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param(
            $Folder,
            $Desc,
            $IncludeFilter,
            $ExcludeFilter,
            $PriorityFilter,
            $EwsService,
            $emailAddress,
            $DeletedItemsFolder= $null
        )

        $ProcessingOK= $True
        $ThisMailboxMode= $Mode
        $temp= $null
        $TotalMatch= 0
        $TotalRemoved= 0
        $FoldersFound= 0
        $FoldersProcessed= 0
        $TimeProcessingStart= Get-Date
        $PidTagSearchKey= New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition( 0x300B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

        # Build list of folders to process
        Write-Verbose ('Collecting folders to process..')
        $FoldersToProcess= Get-SubFolders -Folder $Folder -CurrentPath '' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService

        $FoldersFound= $FoldersToProcess.Count
        Write-Verbose ('Found {0} folders that match search criteria' -f $FoldersFound)

        # Sort complete set of folders on Priority
        $FoldersToProcess= $FoldersToProcess | Sort-Object Priority -Descending

        ForEach ( $SubFolder in $FoldersToProcess) {

            If (!$NoProgressBar) {
                Write-Progress -Id 1 -Activity ('Processing {0} ({1})' -f $Identity, $Desc) -Status ('Processed folder {0} of {1}' -f $FoldersProcessed, $FoldersFound) -PercentComplete ( $FoldersProcessed / $FoldersFound * 100)
            }
            If ( ! ( $DeleteMode -eq 'MoveToDeletedItems' -and $SubFolder.Folder.Id -eq $DeletedItemsFolder.Id)) {
                If ( $Report.IsPresent) {
                    Write-Host ('Processing folder {0}' -f $SubFolder.Name)
                }
                Else {
                    Write-Verbose ('Processing folder {0}' -f $SubFolder.Name)
                }
                $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( $MaxItemBatchSize, 0, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                $ItemView.Traversal= [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                If ( $Retain -eq 'Oldest') {
                    $ItemView.OrderBy.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, [Microsoft.Exchange.WebServices.Data.SortDirection]::Ascending)
                }
                Else {
                    $ItemView.OrderBy.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)
                }
                $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet( [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $ItemView.PropertySet.Add( $PidTagSearchKey)

                $DuplicateList= [System.Collections.ArrayList]@()
                $TotalDuplicates= 0
                $TotalFolder= 0
                $type= ("System.Collections.Generic.List" + '`' + "1") -as 'Type'
                $type= $type.MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId] -as 'Type')
                $ItemIds= [Activator]::CreateInstance($type)

                Do {
                    $SendCancellationsMode= $null
                    $AffectedTaskOccurrences= [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::AllOccurrences
                    $ItemSearchResults= MyEWSFind-ItemsNoSearch $SubFolder.Folder $ItemView
                    Write-Debug ('Checking {0} items in {1}' -f $ItemSearchResults.Items.Count, $SubFolder.Name)
                    If (!$NoProgressBar) {
                        Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.Name) -Status ('Finding duplicate items: Checked {0}, found {1}' -f $TotalFolder, $TotalDuplicates)
                    }
                    If ( $ItemSearchResults.Items.Count -gt 0) {
                        If( $ThisMailboxMode -ne 'Quick') {
                            # Fetch properties for found items to conduct matching
                            $EwsService.LoadPropertiesForItems( $ItemSearchResults.Items, $ItemView.PropertySet)  
                        }

                        ForEach ( $Item in $ItemSearchResults.Items) {
                            Write-Debug ('Inspecting item {0} of {1}, modified {2}' -f $Item.Subject, $Item.DateTimeReceived, $Item.LastModifiedTime)
                            $TotalFolder++
                            $TotalMatch++
                            if ($ThisMailboxMode -eq 'Body'){
                                # Use PR_MESSAGE_BODY for matching duplicates
                                $key= $Item.Body
                            }
                            if ($ThisMailboxMode -eq 'Quick') {
                                # Use PidTagSearchKey for matching duplicates
                                $PropVal= $null
                                if ( $Item.TryGetProperty( $PidTagSearchKey, [ref]$PropVal)) {
                                    $key= [System.BitConverter]::ToString($PropVal).Replace("-", "")
                                }
                                Else {
                                    Write-Debug 'Cannot access or missing PidTagSearchKey property, falling back to property mode (Full)'
                                    $ThisMailboxMode= 'Full'
                                }
                            }
                            If ( $ThisMailboxMode -eq 'Full') {
                                # Use predefined criteria for matching duplicates depending on ItemClass
                                $key= $Item.ItemClass
                                switch ($Item.ItemClass) {
                                    'IPM.Note' {
                                        if ($Item.DateTimeReceived) { $key += $Item.DateTimeReceived.ToString()}
                                        if ($Item.Subject) { $key += $Item.Subject}
                                        if ($Item.InternetMessageId) { $key += $Item.InternetMessageId}
                                        if ($Item.DateTimeSent) { $key += $Item.DateTimeSent.ToString()}
                                        if ($Item.Sender) { $key += $Item.Sender}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                    'IPM.Appointment' {
                                        if ($Item.Subject) { $key += $Item.Subject}
                                        if ($Item.Location) { $key += $Item.Location}
                                        if ($Item.Start) { $key += $Item.Start.ToString()}
                                        if ($Item.End) { $key += $Item.End.ToString()}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                    'IPM.Contact' {
                                        if ($Item.FileAs) { $key += $Item.FileAs}
                                        if ($Item.GivenName) { $key += $Item.GivenName}
                                        if ($Item.Surname) { $key += $Item.Surname}
                                        if ($Item.CompanyName) { $key += $Item.CompanyName}
                                        if ($Item.PhoneNUmbers.TryGetValue('BusinessPhone', [ref]$temp)) { $key += $temp}
                                        if ($Item.PhoneNUmbers.TryGetValue('HomePhone', [ref]$temp)) { $key += $temp}
                                        if ($Item.PhoneNUmbers.TryGetValue('MobilePhone', [ref]$temp)) { $key += $temp}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                    'IPM.DistList' {
                                        if ($Item.FileAs) { $key += $Item.FileAs}
                                        if ($Item.Members) { $key += $Item.Members.Count.ToString()}
                                    }
                                    'IPM.Task' {
                                        if ($Item.Subject) { $key += $Item.Subject}
                                        if ($Item.StartDate) { $key += $Item.StartDate.ToString()}
                                        if ($Item.DueDate) { $key += $Item.DueDate.ToString()}
                                        if ($Item.Status) { $key += $Item.Status}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                    'IPM.Post' {
                                        if ($Item.Subject) { $key += $Item.Subject}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                    Default {
                                        if ($Item.DateTimeReceived) { $key += $Item.DateTimeReceived.ToString()}
                                        if ($Item.Subject) { $key += $Item.Subject}
                                        If (!$NoSize) {if ($Item.Size) { $key += $Item.Size.ToString()}}
                                    }
                                }
                            }
                            If ( $null -ne $key) {
                                $hash= Get-Hash $key
                                If ( $global:UniqueList.contains( $hash)) {
                                    If ( $Report.IsPresent) {
                                        Write-Host ('Item: {0} of {1} ({2})' -f $Item.Subject, $Item.DateTimeReceived, $Item.ItemClass)
                                    }
                                    Write-Debug "Duplicate: $hash ($key)"
                                    $DuplicateList.Add( $Item.Id) | Out-Null
                                    $TotalDuplicates++
                                }
                                Else {
                                    Write-Debug "Unique: $($Item.id), $hash ($key)"
                                    $global:UniqueList.Add( $hash) | Out-Null
                                }
                            }
                            Else {
                                # Couldn't determine key, skip
                            }
                        }
                        $ItemView.Offset += $ItemSearchResults.Items.Count
                    }
                    Else {
                        # No items found

                    }
                } While ( $ItemSearchResults.MoreAvailable -and $ProcessingOK)
            }
            Else {
                Write-Debug 'Skipping Deleted Items folder'
            }

            $TotalMatch += $ItemSearchResults.TotalCount

            If ( ($DuplicateList.Count -gt 0) -and ($Force -or $PSCmdlet.ShouldProcess( ('Remove {0} items from {1}' -f $DuplicateList.Count, $SubFolder.Name)))) {
                try {
                    Write-Verbose ('Removing {0} items from {1}' -f $TotalDuplicates, $SubFolder.Name)

                    $SendCancellationsMode= [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone
                    $AffectedTaskOccurrences= [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly
                    $SuppressReadReceipt= $true # Only works using EWS with Exchange2013+ mode

                    $ItemsRemoved= 0
                    $ItemsRemaining= $DuplicateList.Count

                    # Remove ItemIDs in batches
                    ForEach ( $ItemID in $DuplicateList) {
                        $ItemIds.Add( $ItemID)
                        If ( $ItemIds.Count -eq $MaxDeleteBatchSize) {
                            $ItemsRemoved += $ItemIds.Count
                            $ItemsRemaining -= $ItemIds.Count
                            If (!$NoProgressBar) {
                                Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.DisplayName) -Status ('Items removed {0} - remaining {1}' -f $ItemsRemoved, $ItemsRemaining) -PercentComplete ( $ItemsRemoved / $DuplicateList.Count * 100)
                            }
                            myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt | Out-Null
                            $ItemIds.Clear()
                        }
                    }
                    # .. also remove last ItemIDs
                    If ( $ItemIds.Count -gt 0) {
                        $ItemsRemoved += $ItemIds.Count
                        $ItemsRemaining= 0
                        myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt | Out-Null
                        $ItemIds.Clear()
                    }
                    $TotalRemoved += $DuplicateList.Count
                }
                catch {
                    Write-Error ('Problem removing items: {0}' -f $_.Exception.Message)
                    $ProcessingOK= $False
                }
            }
            Else {
                Write-Debug 'No duplicates found in this folder'
            }
            $FoldersProcessed++

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.DisplayName) -Status 'Finished processing.' -Completed
            }

            # If not operating against whole mailbox, clear unique list after processing every folder
            If ( $CleanupMode -eq 'Folder') {
                Write-Verbose ('Cleaning unique list (finished folder)')
                $global:UniqueList= [System.Collections.ArrayList]@()
            }

        } # ForEach SubFolder

        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status 'Finished processing.' -Completed
        }

        # Not MultiMailbox (per mailbox), track MD5 hashes per mailbox
        If ( $CleanupMode -ne 'MultiMailbox' ) {
            Write-Verbose ('Cleaning unique list')
            $global:UniqueList= [System.Collections.ArrayList]@()
        }
        
        If ( $ProcessingOK) {
            $TimeProcessingDiff= (Get-Date) - $TimeProcessingStart
            $Speed= [int]( $TotalMatch / $TimeProcessingDiff.TotalSeconds * 60)
            Write-Host ('{0} items processed and {1} removed in {2:hh}:{2:mm}:{2:ss} - average {3} items/min' -f $TotalMatch, $TotalRemoved, $TimeProcessingDiff, $Speed)
        }
        Return $ProcessingOK
    }

    Import-ModuleDLL -Name 'Microsoft.Exchange.WebServices' -FileName 'Microsoft.Exchange.WebServices.dll' -Package 'Exchange.WebServices.Managed.Api'
    Import-ModuleDLL -Name 'Microsoft.Identity.Client' -FileName 'Microsoft.Identity.Client.dll' -Package 'Microsoft.Identity.Client'

    If ( $MailboxOnly) {
        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
    }
    Else {
        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    }
    $EwsService= [Microsoft.Exchange.WebServices.Data.ExchangeService]::new( $ExchangeVersion)

    If( $Credentials -or $UseDefaultCredentials) {
        If( $Credentials) {
            try {
                Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
                $EwsService.Credentials= [System.Net.NetworkCredential]::new( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
            }
            catch {
                Write-Error ('Invalid credentials provided: {0}' -f $_.Exception.Message)
                Exit $ERR_INVALIDCREDENTIALS
            }
        }
        Else {
            Write-Verbose ('Using Default Credentials')
            $EwsService.UseDefaultCredentials = $true
        }
    }
    Else {

        # Use OAuth (and impersonation/X-AnchorMailbox always set)
        $Impersonation= $true

        If( $CertificateThumbprint -or $CertificateFile) {
            If( $CertificateFile) {
                
                # Use certificate from file using absolute path to authenticate
                $CertificateFile= (Resolve-Path -Path $CertificateFile).Path
                
                Try {
                    If( $CertificatePassword) {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $CertificatePassword)))
                    }
                    Else {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile)
                    }
                }
                Catch {
                    Write-Error ('Problem importing PFX: {0}' -f $_.Exception.Message)
                    Exit $ERR_PROBLEMIMPORTINGCERT
                }
            }
            Else {
                # Use provided certificateThumbprint to retrieve certificate from My store, and authenticate with that
                $CertStore= [System.Security.Cryptography.X509Certificates.X509Store]::new( [Security.Cryptography.X509Certificates.StoreName]::My, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
                $CertStore.Open( [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly )
                $X509Certificate2= $CertStore.Certificates.Find( [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $CertificateThumbprint, $False) | Select-Object -First 1
                If(!( $X509Certificate2)) {
                    Write-Error ('Problem locating certificate in My store: {0}' -f $error[0])
                    Exit $ERR_CERTNOTFOUND
                }
            }
            Write-Verbose ('Will use certificate {0}, issued by {1} and expiring {2}' -f $X509Certificate2.Thumbprint, $X509Certificate2.Issuer, $X509Certificate2.NotAfter)
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithCertificate( $X509Certificate2).withTenantId( $TenantId).Build()
               
        }
        Else {
            # Use provided secret to authenticate
            Write-Verbose ('Will use provided secret to authenticate')
            $PlainSecret= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Secret))
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithClientSecret( $PlainSecret).withTenantId( $TenantId).Build()
        }
        $Scopes= New-Object System.Collections.Generic.List[string]
        $Scopes.Add( 'https://outlook.office365.com/.default')
        Try {
            $Response=$App.AcquireTokenForClient( $Scopes).executeAsync()
            $Token= $Response.Result
            $EwsService.Credentials= [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$Token.AccessToken
            Write-Verbose ('Authentication token acquired')
        }
        Catch {
            Write-Error ('Problem acquiring token: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }

    Write-Verbose ('Cleanup Mode: {0}' -f $CleanupMode)

    If( $TrustAll) {
        Set-SSLVerification -Disable
    }
}

Process {

    ForEach ( $CurrentIdentity in $Identity) {

        $EmailAddress= get-EmailAddress -Identity $CurrentIdentity
        If ( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $EmailAddress)
            Exit $ERR_MAILBOXNOTFOUND
        }

        If( $PublicFolders) {
            Write-Host ('Processing Public Folders as {0} ({1})' -f $EmailAddress, $CurrentIdentity)
        }
        Else {
            Write-Host ('Processing mailbox {0} ({1})' -f $EmailAddress, $CurrentIdentity)
        }

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId= [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new( [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Clear()
            $EwsService.HttpHeaders.Add( 'X-AnchorMailbox', $EmailAddress)
        }
            
        If ($Server) {
            $EwsUrl= 'https://{0}/EWS/Exchange.asmx' -f $Server
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url= $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose ('Using EWS endpoint {0}' -f $EwsService.Url)
        } 

        # Construct search filters
        Write-Verbose 'Constructing folder matching rules'
        $IncludeFilter= Construct-FolderFilter $EwsService $IncludeFolders $EmailAddress
        $ExcludeFilter= Construct-FolderFilter $EwsService $ExcludeFolders $EmailAddress
        $PriorityFilter= Construct-FolderFilter $EwsService $PriorityFolders $EmailAddress

        If ( $PublicFolders.IsPresent) {
            try {
                $RootFolder= myEWSBind-WellKnownFolder $EwsService 'PublicFoldersRoot' 
                If ($null -ne $RootFolder) {
                    Write-Verbose ('Processing Public Folders')
                    If (! ( Process-Mailbox -Folder $RootFolder -Desc 'Public Folders' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService -emailAddress $emailAddress)) {
                        Write-Error ('Problem processing Public Folders as {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                        Exit $ERR_PROCESSINGPUBLICFOLDERS
                    }
                }
            }
            catch {
                Write-Error ('Cannot access public folders as {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                Exit $ERR_CANTACCESSPUBLICFOLDERS
            }
            Write-Verbose ('Processing Public Folders finished')
        }
        Else {
            If ( -not $ArchiveOnly.IsPresent) {
                try {
                    $RootFolder= myEWSBind-WellKnownFolder $EwsService 'MsgFolderRoot' $EmailAddress
                    If ($null -ne $RootFolder) {
                        Write-Verbose ('Processing primary mailbox {0}' -f $EmailAddress)
                        $DeletedItemsFolder= myEWSBind-WellKnownFolder $EwsService 'DeletedItems' $emailAddress
                        If (! ( Process-Mailbox -Folder $RootFolder -Desc 'Mailbox' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService -emailAddress $emailAddress -DeletedItemsFolder $DeletedItemsFolder)) {
                            Write-Error ('Problem processing primary mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                            Exit $ERR_PROCESSINGMAILBOX
                        }
                    }
                }
                catch {
                    Write-Error ('Cannot access mailbox information store for {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                    Exit $ERR_CANTACCESSMAILBOXSTORE
                }
            }

            If ( -not $MailboxOnly.IsPresent) {
                try {
                    $ArchiveRootFolder= myEWSBind-WellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $EmailAddress
                    If ($null -ne $ArchiveRootFolder) {
                        Write-Verbose ('Processing archive mailbox {0}' -f $EmailAddress)
                        $DeletedItemsFolder= myEWSBind-WellKnownFolder $EwsService 'ArchiveDeletedItems' $emailAddress
                        If (! ( Process-Mailbox -Folder $ArchiveRootFolder -Desc 'Archive' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService -emailAddress $emailAddress -DeletedItemsFolder $DeletedItemsFolder)) {
                            Write-Error ('Problem processing archive mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                            Exit $ERR_PROCESSINGARCHIVE
                        }
                    }
                }
                catch {
                    Write-Debug 'No archive configured or cannot access archive'
                }
                Write-Verbose ('Processing {0} finished' -f $EmailAddress)
            }
        }
    }
}

end {
    If( $TrustAll) {
        Set-SSLVerification -Enable
    }
}
