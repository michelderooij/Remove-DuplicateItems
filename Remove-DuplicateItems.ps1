<#
    .SYNOPSIS
    Remove-DuplicateItems

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 1.88, February 9th, 2020

    .DESCRIPTION
    This script will scan each folder of a given primary mailbox and personal archive (when
    configured, Exchange 2010 and later) and removes duplicate items per folder. You can specify
    how the items should be deleted and what items to process, e.g. mail items or appointments.
    Sample scenarios are misbehaving 3rd party synchronization tool creates duplicate items or
    (accidental) import of PST file with duplicate items. Script will process
    mailbox and archive if configured, unless MailboxOnly or ArchiveOnly is specified. For
    Exchange 2007, you need to specify -MailboxOnly.

    Note that usage of the Verbose, Confirm and WhatIf parameters is supported.
    When using Confirm, you will be prompted per batch.

    .LINK
    http://eightwone.com

    .NOTES
    Microsoft Exchange Web Services (EWS) Managed API 1.2 or up is required.

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

    .PARAMETER Identity
    Identity of the Mailbox. Can be CN/SAMAccountName (for on-premises) or e-mail format (on-prem & Office 365)

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited,
    script will attempt to use Autodiscover.

    .PARAMETER Credentials
    Specify credentials to use. When not specified, current credentials are used.
    Credentials can be set using $Credentials= Get-Credential

    .PARAMETER Impersonation
    When specified, uses impersonation for mailbox access, otherwise current
    logged on user is used. For details on how to configure impersonation
    access for Exchange 2010 using RBAC, see this article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095(v=exchg.140).aspx
    For details on how to configure impersonation for Exchange 2007, see KB article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095%28v=exchg.80%29.aspx

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

    You can also use well-known folders, by using this format: #WellKnownFolderName#, e.g. #Inbox#.
    Supported are #Calendar#, #Contacts#, #Inbox#, #Notes#, #SentItems#, #Tasks#, #JunkEmail# and #DeletedItems#.
    The script uses the currently configured Well-Known Folder of the mailbox to be processed.

    .PARAMETER ExcludeFolders
    Specify one or more folder(s) to exclude. Usage of wildcards and well-known folders identical to IncludeFolders.
    Note that ExcludeFolders criteria overrule IncludeFolders when matching folders.

    .PARAMETER Force
    Force removal of items without prompting.

    .PARAMETER MailboxWide
    Performs duplicate cleanup against whole mailbox, instead of per folder.
    By default, the first unique item encountered will be retained. When an item is found in Folder A and
    in Folder B, it is undetermined which item will be kept, unless PriorityFolders is used.

    .PARAMETER PriorityFolders
    Determines which folders have priority over other folders, identifying items in these folders first when
    using MailboxWide mode. Usage of wildcards and well-known folders is identical to IncludeFolders.

    .PARAMETER NoSize
    Don't use size to match items in Full mode.

    .PARAMETER Report
    Reports individual items detected as duplicate. Can be used together with WhatIf to perform pre-analysis.

    .PARAMETER NoProgressBar
    Use this switch to prevent displaying a progress bar as folders and items are being processed.

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
    Credentials provided earlier.

    .EXAMPLE
    $Credentials= Get-Credential
    .\Remove-DuplicateItems.ps1 -Identity olrik@office365tenant.com -Server outlook.office365.com -Credentials $Credentials -IncludeFolders '#Inbox#\*','\Projects\*' -ExcludeFolders 'Keep Out' -PriorityFolders '*Important*' -MailboxWide

    Remove duplicate items from specified mailbox in Office365 using fixed Server FQDN - bypassing AutoDiscover, limiting
    operation against the Inbox, and top Projects folder, and all of their subfolders, but excluding any folder named Keep Out.
    Duplicates are checked over all folders, but priority is given to folders containing the word Important, causing items in
    those folders to be kept over items in other folders when duplicates are found.
#>

[cmdletbinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = "High"
)]
param(
    [parameter( Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = "All")]
    [alias('Mailbox')]
    [string]$Identity,
    [parameter( Position = 1, Mandatory = $false, ParameterSetName = "All")]
    [ValidateSet("Mail", "Calendar", "Contacts", "Tasks", "Notes", "Groups", "All")]
    [string]$Type = "All",
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [ValidateSet("Oldest", "Newest")]
    [string]$Retain = "Newest",
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string]$Server,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$Impersonation,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [ValidateSet("HardDelete", "SoftDelete", "MoveToDeletedItems")]
    [string]$DeleteMode = 'SoftDelete',
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [ValidateSet("Quick", "Full", "Body")]
    [string]$Mode = 'Quick',
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [parameter( Mandatory = $false, ParameterSetName = "MailboxOnly")]
    [switch]$MailboxOnly,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [parameter( Mandatory = $false, ParameterSetName = "ArchiveOnly")]
    [switch]$ArchiveOnly,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string[]]$IncludeFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string[]]$ExcludeFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$MailboxWide,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [string[]]$PriorityFolders,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$NoSize,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$Force,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$NoProgressBar,
    [parameter( Mandatory = $false, ParameterSetName = "All")]
    [switch]$Report
)

process {

    # Process folders these batches
    $MaxFolderBatchSize = 100
    # Process items in these page sizes
    $MaxItemBatchSize = 100
    # Max of concurrent item deletes
    $MaxDeleteBatchSize = 100

    # Initial sleep timer (ms) and treshold before lowering
    $script:SleepTimerMax = 300000               # Maximum delay (5min)
    $script:SleepTimerMin = 100                  # Minimum delay
    $script:SleepAdjustmentFactor = 2.0          # When tuning, use this factor
    $script:SleepTimer = $script:SleepTimerMin   # Initial sleep timer value

    # Errors
    $ERR_EWSDLLNOTFOUND = 1000
    $ERR_EWSLOADING = 1001
    $ERR_MAILBOXNOTFOUND = 1002
    $ERR_AUTODISCOVERFAILED = 1003
    $ERR_CANTACCESSMAILBOXSTORE = 1004
    $ERR_PROCESSINGMAILBOX = 1005
    $ERR_PROCESSINGARCHIVE = 1006
    $ERR_INVALIDCREDENTIALS = 1007

    Function Get-EmailAddress( $Identity) {
        $address = [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if ( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as SamAccountName
            $ADSearch = New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter = "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result = $ADSearch.FindOne()
            If ( $Result) {
                $objUser = $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function Load-EWSManagedAPIDLL {
        $EWSDLL = "Microsoft.Exchange.WebServices.dll"
        If ( Test-Path "$pwd\$EWSDLL") {
            $EWSDLLPath = "$pwd"
        }
        Else {
            $EWSDLLPath = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
            if (!( Test-Path "$EWSDLLPath\$EWSDLL")) {
                Write-Error "This script requires EWS Managed API 1.2 or later to be installed, or the Microsoft.Exchange.WebServices.DLL in the current folder."
                Write-Error "You can download and install EWS Managed API from http://go.microsoft.com/fwlink/?LinkId=255472"
                Exit $ERR_EWSDLLNOTFOUND
            }
        }

        Write-Verbose "Loading $EWSDLLPath\$EWSDLL"
        try {
            # EX2010
            If (!( Get-Module Microsoft.Exchange.WebServices)) {
                Import-Module "$EWSDLLPATH\$EWSDLL"
            }
        }
        catch {
            #<= EX2010
            [void][Reflection.Assembly]::LoadFile( "$EWSDLLPath\$EWSDLL")
        }
        try {
            $Temp = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
        }
        catch {
            Write-Error "Problem loading $EWSDLL"
            Exit $ERR_EWSLOADING
        }
        $DLLObj = Get-ChildItem -Path "$EWSDLLPATH\$EWSDLL" -ErrorAction SilentlyContinue
        If ( $DLLObj) {
            Write-Verbose ('Loaded EWS Managed API v{0}' -f $DLLObj.VersionInfo.FileVersion)
        }
    }

    # After calling this any SSL Warning issues caused by Self Signed Certificates will be ignored
    # Source: http://poshcode.org/624
    Function set-TrustAllWeb() {
        Write-Verbose "Set to trust all certificates"
        $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler = $Provider.CreateCompiler()
        $Params = New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable = $False
        $Params.GenerateInMemory = $True
        $Params.IncludeDebugInformation = $False
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

        $TASource = @'
            namespace Local.ToolkitExtensions.Net.CertificatePolicy {
                public class TrustAll : System.Net.ICertificatePolicy {
                    public TrustAll() {
                    }
                    public bool CheckValidationResult(System.Net.ServicePoint sp, System.Security.Cryptography.X509Certificates.X509Certificate cert,   System.Net.WebRequest req, int problem) {
                        return true;
                    }
                }
            }
'@
        $TAResults = $Provider.CompileAssemblyFromSource($Params, $TASource)
        $TAAssembly = $TAResults.CompiledAssembly
        $TrustAll = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
    }

    Function iif( $eval, $tv = '', $fv = '') {
        If ( $eval) { return $tv } else { return $fv}
    }

    Function Construct-FolderFilter {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string[]]$Folders,
            [string]$emailAddress
        )
        If ( $Folders) {
            $FolderFilterSet = @()
            ForEach ( $Folder in $Folders) {
                # Convert simple filter to (simple) regexp
                $Parts = $Folder -match '^(?<root>\\)?(?<keywords>.*?)?(?<sub>\\\*)?$'
                If ( !$Parts) {
                    Write-Error ('Invalid regular expression matching against {0}' -f $Folder)
                }
                Else {
                    $Keywords = Search-ReplaceWellKnownFolderNames $EwsService ($Matches.keywords) $emailAddress
                    $EscKeywords = [Regex]::Escape( $Keywords) -replace '\\\*', '.*'
                    $Pattern = iif -eval $Matches.Root -tv '^\\' -fv '^\\(.*\\)*'
                    $Pattern += iif -eval $EscKeywords -tv $EscKeywords -fv ''
                    $Pattern += iif -eval $Matches.sub -tv '(\\.*)?$' -fv '$'
                    $Obj = New-Object -TypeName PSObject -Prop @{
                        'Pattern'     = $Pattern;
                        'IncludeSubs' = -not [string]::IsNullOrEmpty( $Matches.Sub)
                        'OrigFilter'  = $Folder
                    }
                    $FolderFilterSet += $Obj
                    Write-Debug ($Obj -join ',')
                }
            }
        }
        Else {
            $FolderFilterSet = $null
        }
        return $FolderFilterSet
    }

    Function Search-ReplaceWellKnownFolderNames {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$criteria = '',
            [string]$emailAddress
        )
        $AllowedWKF = 'Inbox', 'Calendar', 'Contacts', 'Notes', 'SentItems', 'Tasks', 'JunkEmail', 'DeletedItems'
        # Construct regexp to see if allowed WKF is part of criteria string
        ForEach ( $ThisWKF in $AllowedWKF) {
            If ( $criteria -match '#{0}#') {
                $criteria = $criteria -replace ('#{0}#' -f $ThisWKF), (myEWSBind-WellKnownFolder $EwsService $ThisWKF $emailAddress).DisplayName
            }
        }
        return $criteria
    }
    Function Tune-SleepTimer {
        param(
            [bool]$previousResultSuccess = $false
        )
        if ( $previousResultSuccess) {
            If ( $script:SleepTimer -gt $script:SleepTimerMin) {
                $script:SleepTimer = [int]([math]::Max( [int]($script:SleepTimer / $script:SleepAdjustmentFactor), $script:SleepTimerMin))
                Write-Warning ('Previous EWS operation successful, adjusted sleep timer to {0}ms' -f $script:SleepTimer)
            }
        }
        Else {
            $script:SleepTimer = [int]([math]::Min( ($script:SleepTimer * $script:SleepAdjustmentFactor) + 100, $script:SleepTimerMax))
            If ( $script:SleepTimer -eq 0) {
                $script:SleepTimer = 5000
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $EwsService.FindFolders( $FolderId, $FolderSearchCollection, $FolderView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $EwsService.FindFolders( $FolderId, $FolderView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $Folder.FindItems( $ItemSearchFilterCollection, $ItemView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $CritErr = $false
        Do {
            Try {
                $res = $Folder.FindItems( $ItemView)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $critErr = $false
        Do {
            Try {
                If ( @([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013, [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1) -contains $EwsService.RequestedServerVersion) {
                    $res = $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences, $SuppressReadReceipt)
                }
                Else {
                    $res = $EwsService.DeleteItems( $ItemIds, $DeleteMode, $SendCancellationsMode, $AffectedTaskOccurrences)
                }
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
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
        $OpSuccess = $false
        $critErr = $false
        Do {
            Try {
                $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $emailAddress)  
                $res = [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $explicitFolder)
                $OpSuccess = $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess = $false
                Write-Warning 'EWS operation failed, server busy - will retry later'
            }
            catch {
                $OpSuccess = $false
                $critErr = $true
                Write-Warning ('Cannot bind to {0} - skipping. Error: {1}' -f $WellKnownFolderName, $Error[0])
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
        $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
        $data = ([system.Text.Encoding]::UTF8).GetBytes( $string)
        return ([System.BitConverter]::ToString( $md5.ComputeHash( $data)) -replace ' - ', '')
    }

    Function Get-FolderPriority {
        param(
            $FolderPath,
            $PriorityFilter
        )
        $prio = 0
        If ( $PriorityFilter) {
            $num = 0
            ForEach ( $Filter in $PriorityFilter) {
                $num++
                If ( $FolderPath -match $Filter.Pattern) {
                    $prio = $num
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
            $PriorityFilter
        )
        $FoldersToProcess = [System.Collections.ArrayList]@()
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView( $MaxFolderBatchSize)
        $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass)
        $FolderSearchCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection( [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        If ( $Type -ne "All") {
            $FolderSearchClass = (@{"Mail" = "IPF.Note"; "Calendar" = "IPF.Appointment"; "Contacts" = "IPF.Contact"; "Tasks" = "IPF.Task"; "Notes" = "IPF.StickyNotes"})[$Type]
            $FolderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, $FolderSearchClass)
            $FolderSearchCollection.Add( $FolderSearchFilter)
        }
        Do {
            If ( $FolderSearchCollection.Count -ge 1) {
                $FolderSearchResults = myEWSFind-Folders $EwsService $Folder.Id $FolderSearchCollection $FolderView
            }
            Else {
                $FolderSearchResults = myEWSFind-FoldersNoSearch $EwsService $Folder.Id $FolderView
            }
            ForEach ( $FolderItem in $FolderSearchResults) {
                $FolderPath = '{0}\{1}' -f $CurrentPath, $FolderItem.DisplayName
                If ( $IncludeFilter) {
                    $Add = $false
                    # Defaults to true, unless include does not specifically include subfolders
                    $Subs = $true
                    ForEach ( $Filter in $IncludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add = $true
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs = $Filter.IncludeSubs
                        }
                    }
                }
                Else {
                    # If no includeFolders specified, include all (unless excluded)
                    $Add = $true
                    $Subs = $true
                }
                If ( $ExcludeFilter) {
                    # Excludes can overrule includes
                    ForEach ( $Filter in $ExcludeFilter) {
                        If ( $FolderPath -match $Filter.Pattern) {
                            $Add = $false
                            # When multiple criteria match, one with and one without subfolder processing, subfolders will be processed.
                            $Subs = $Filter.IncludeSubs
                        }
                    }
                }
                If ( $Add) {
                    $Prio = Get-FolderPriority $FolderPath -PriorityFilter $PriorityFilter
                    Write-Verbose ( 'Adding folder {0} (priority {1})' -f $FolderPath, $Prio)

                    $Obj = New-Object -TypeName PSObject -Property @{
                        'Name'     = $FolderPath;
                        'Priority' = $Prio;
                        'Folder'   = $FolderItem
                    }
                    $FoldersToProcess.Add( $Obj) | Out-Null
                }
                If ( $Subs) {
                    # Could be that specific folder is to be excluded, but subfolders needs evaluation
                    ForEach ( $AddFolder in (Get-SubFolders -Folder $FolderItem -CurrentPath $FolderPath -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter)) {
                        $FoldersToProcess.Add( $AddFolder)  | Out-Null
                    }
                }
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)
        Write-Output -NoEnumerate $FoldersToProcess
    }

    Function Process-Mailbox {
        param(
            $Folder,
            $Desc,
            $IncludeFilter,
            $ExcludeFilter,
            $PriorityFilter,
            $EwsService,
            $emailAddress
        )

        $ProcessingOK = $True
        $ThisMailboxMode = $Mode
        $temp = $null
        $TotalMatch = 0
        $TotalRemoved = 0
        $FoldersFound = 0
        $FoldersProcessed = 0
        $TimeProcessingStart = Get-Date
        $DeletedItemsFolder = myEWSBind-WellKnownFolder $EwsService 'DeletedItems' $emailAddress
        $PidTagSearchKey = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition( 0x300B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

        # Build list of folders to process
        Write-Verbose ('Collecting folders to process, type {0}' -f $Type)
        $FoldersToProcess = Get-SubFolders -Folder $Folder -CurrentPath '' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter

        $FoldersFound = $FoldersToProcess.Count
        Write-Verbose ('Found {0} folders that match search criteria' -f $FoldersFound)

        # Sort complete set of folders on Priority
        $FoldersToProcess = $FoldersToProcess | Sort Priority -Descending

        # Initialize list to keep track of unique items
        $UniqueList = [System.Collections.ArrayList]@()

        ForEach ( $SubFolder in $FoldersToProcess) {
            If (!$NoProgressBar) {
                Write-Progress -Id 1 -Activity ('Processing {0} ({1})' -f $Identity, $Desc) -Status "Processed folder $FoldersProcessed of $FoldersFound" -PercentComplete ( $FoldersProcessed / $FoldersFound * 100)
            }
            If ( ! ( $DeleteMode -eq "MoveToDeletedItems" -and $SubFolder.Folder.Id -eq $DeletedItemsFolder.Id)) {
                If ( $Report.IsPresent) {
                    Write-Host ('Processing folder {0}' -f $SubFolder.Name)
                }
                Else {
                    Write-Verbose ('Processing folder {0}' -f $SubFolder.Name)
                }
                $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView( $MaxItemBatchSize, 0, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
                $ItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                If ( $Retain -eq 'Oldest') {
                    $ItemView.OrderBy.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, [Microsoft.Exchange.WebServices.Data.SortDirection]::Ascending)
                }
                Else {
                    $ItemView.OrderBy.Add( [Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)
                }
                $ItemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet( [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $ItemView.PropertySet.Add( $PidTagSearchKey)

                # Not MailboxWide (per folder), track MD5 hashes per folder
                If ( -not $MailboxWide) {
                    $UniqueList = [System.Collections.ArrayList]@()
                }
                $DuplicateList = [System.Collections.ArrayList]@()
                $TotalDuplicates = 0
                $TotalFolder = 0
                If ( $psversiontable.psversion.major -lt 3) {
                    $ItemIds = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
                }
                Else {
                    $type = ("System.Collections.Generic.List" + '`' + "1") -as 'Type'
                    $type = $type.MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId] -as 'Type')
                    $ItemIds = [Activator]::CreateInstance($type)
                }
                Do {
                    $SendCancellationsMode = $null
                    $AffectedTaskOccurrences = [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::AllOccurrences
                    $ItemSearchResults = MyEWSFind-ItemsNoSearch $SubFolder.Folder $ItemView
                    Write-Debug "Checking $($ItemSearchResults.Items.Count) items in $($SubFolder.Name)"
                    If (!$NoProgressBar) {
                        Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.Name)" -Status "Finding duplicate items, checked $TotalFolder, found $TotalDuplicates"
                    }
                    If ( $ItemSearchResults.Items.Count -gt 0) {
			If( $ThisMailboxMode -ne 'Quick') {
                            # Fetch properties for found items to conduct matching
                            $EwsService.LoadPropertiesForItems( $ItemSearchResults.Items, $ItemView.PropertySet)  
                        }

                        ForEach ( $Item in $ItemSearchResults.Items) {
                            Write-Debug "Inspecting item $($Item.Subject) of $($Item.DateTimeReceived), modified $($Item.LastModifiedTime)"
                            $TotalFolder++
                            $TotalMatch++
                            if ($ThisMailboxMode -eq 'Body'){
                                # Use PR_MESSAGE_BODY for matching duplicates
                                $key = $Item.Body
#$key
                            }
                            if ($ThisMailboxMode -eq 'Quick') {
                                # Use PidTagSearchKey for matching duplicates
                                $PropVal = $null
                                if ( $Item.TryGetProperty( $PidTagSearchKey, [ref]$PropVal)) {
                                    $key = [System.BitConverter]::ToString($PropVal).Replace("-", "")
                                }
                                Else {
                                    Write-Debug 'Cannot access or missing PidTagSearchKey property, falling back to property mode (Full)'
                                    $ThisMailboxMode = 'Full'
                                }
                            }
                            If ( $ThisMailboxMode -eq 'Full') {
                                # Use predefined criteria for matching duplicates depending on ItemClass
                                $key = $Item.ItemClass
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
                            If ( $key -ne $null) {
                                $hash = Get-Hash $key
                                If ( $UniqueList.contains( $hash)) {
                                    If ( $Report.IsPresent) {
                                        Write-Host ('Item: {0} of {1} ({2})' -f $Item.Subject, $Item.DateTimeReceived, $Item.ItemClass)
                                    }
                                    Write-Debug "Duplicate: $hash ($key)"
                                    $tmp = $DuplicateList.Add( $Item.Id)
                                    $TotalDuplicates++
                                }
                                Else {
                                    Write-Debug "Unique: $($Item.id), $hash ($key)"
                                    $tmp = $UniqueList.Add( $hash)
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
            If ( ($DuplicateList.Count -gt 0) -and ($Force -or $PSCmdlet.ShouldProcess( "Remove $($DuplicateList.Count) items from $($SubFolder.Name)"))) {
                try {
                    Write-Verbose "Removing $TotalDuplicates items from $($SubFolder.Name)"

                    $SendCancellationsMode = [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone
                    $AffectedTaskOccurrences = [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::SpecifiedOccurrenceOnly
                    $SuppressReadReceipt = $true # Only works using EWS with Exchange2013+ mode

                    $ItemsRemoved = 0
                    $ItemsRemaining = $DuplicateList.Count

                    # Remove ItemIDs in batches
                    ForEach ( $ItemID in $DuplicateList) {
                        $ItemIds.Add( $ItemID)
                        If ( $ItemIds.Count -eq $MaxDeleteBatchSize) {
                            $ItemsRemoved += $ItemIds.Count
                            $ItemsRemaining -= $ItemIds.Count
                            If (!$NoProgressBar) {
                                Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.DisplayName)" -Status "Items removed $ItemsRemoved - remaining $ItemsRemaining" -PercentComplete ( $ItemsRemoved / $DuplicateList.Count * 100)
                            }
                            $res = myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                            $ItemIds.Clear()
                        }
                    }
                    # .. also remove last ItemIDs
                    If ( $ItemIds.Count -gt 0) {
                        $ItemsRemoved += $ItemIds.Count
                        $ItemsRemaining = 0
                        $res = myEWSRemove-Items $EwsService $ItemIds $DeleteMode $SendCancellationsMode $AffectedTaskOccurrences $SuppressReadReceipt
                        $ItemIds.Clear()
                    }
                    $TotalRemoved += $DuplicateList.Count
                }
                catch {
                    Write-Error "Problem removing items: $($error[0])"
                    $ProcessingOK = $False
                }
            }
            Else {
                Write-Debug 'No duplicates found in this folder'
            }
            $FoldersProcessed++

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity "Processing folder $($SubFolder.DisplayName)" -Status 'Finished processing.' -Completed
            }

            # If not operating against whole mailbox, clear unique list after processing every folder
            If ( !$MailboxWide) {
                $UniqueList = [System.Collections.ArrayList]@()
            }

        } # ForEach SubFolder

        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity "Processing $Identity" -Status "Finished processing." -Completed
        }
        If ( $ProcessingOK) {
            $TimeProcessingDiff = (Get-Date) - $TimeProcessingStart
            $Speed = [int]( $TotalMatch / $TimeProcessingDiff.TotalSeconds * 60)
            Write-Host ('{0} items processed and {1} removed in {2:hh}:{2:mm}:{2:ss} - average {3} items/min' -f $TotalMatch, $TotalRemoved, $TimeProcessingDiff, $Speed)
        }
        Return $ProcessingOK
    }

    ##################################################
    # Main
    ##################################################
    #Requires -Version 3

    Load-EWSManagedAPIDLL
    set-TrustAllWeb

    If ( $MailboxOnly) {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
    }
    Else {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    }

    $EwsService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService( $ExchangeVersion)
    If ( $Credentials) {
        try {
            Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
            $EwsService.Credentials = New-Object System.Net.NetworkCredential( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
        }
        catch {
            Write-Error ('Invalid credentials provided, error: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }
    Else {
        $EwsService.UseDefaultCredentials = $true
    }

    Write-Verbose ('Processing items of type {0}, delete mode is {1}' -f $Type, $DeleteMode)

    ForEach ( $CurrentIdentity in $Identity) {

        $EmailAddress = get-EmailAddress $CurrentIdentity
        If ( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $CurrentIdentity)
            Exit $ERR_MAILBOXNOTFOUND
        }

        Write-Host ('Processing mailbox {0} ({1})' -f $CurrentIdentity, $EmailAddress)

        If ( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)
        }

        If ($Server) {
            $EwsUrl = ('https://{0}/EWS/Exchange.asmx' -f $Server)
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url = $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference = 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed, error: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference = 'Continue'
            Write-Verbose ('Using EWS on CAS {0}' -f $EwsService.Url)
        }

        # Construct search filters
        Write-Verbose 'Constructing folder matching rules'
        $IncludeFilter = Construct-FolderFilter $EwsService $IncludeFolders $EmailAddress
        $ExcludeFilter = Construct-FolderFilter $EwsService $ExcludeFolders $EmailAddress
        $PriorityFilter = Construct-FolderFilter $EwsService $PriorityFolders $EmailAddress

        If ( -not $ArchiveOnly.IsPresent) {
            try {
                $RootFolder = myEWSBind-WellKnownFolder $EwsService 'MsgFolderRoot' $EmailAddress
                If ( $RootFolder) {
                    Write-Verbose ('Processing primary mailbox {0}' -f $Identity)
                    If (! ( Process-Mailbox -Folder $RootFolder -Desc 'Mailbox' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService -emailAddress $emailAddress)) {
                        Write-Error ('Problem processing primary mailbox of {0} ({1})' -f $CurrentIdentity, $EmailAddress)
                        Exit $ERR_PROCESSINGMAILBOX
                    }
                }
            }
            catch {
                Write-Error ('Cannot access mailbox information store, error: {0}' -f $Error[0])
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
        }

        If ( -not $MailboxOnly.IsPresent) {
            try {
                $ArchiveRootFolder = myEWSBind-WellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $EmailAddress
                If ( $ArchiveRootFolder) {
                    Write-Verbose ('Processing archive mailbox {0}' -f $Identity)
                    If (! ( Process-Mailbox -Folder $ArchiveRootFolder -Desc 'Archive' -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter -PriorityFilter $PriorityFilter -EwsService $EwsService -emailAddress $emailAddress)) {
                        Write-Error ('Problem processing archive mailbox of {0} ({1})' -f $CurrentIdentity, $EmailAddress)
                        Exit $ERR_PROCESSINGARCHIVE
                    }
                }
            }
            catch {
                Write-Debug 'No archive configured or cannot access archive'
            }
        }
        Write-Verbose ('Processing {0} finished' -f $CurrentIdentity)
    }
}
