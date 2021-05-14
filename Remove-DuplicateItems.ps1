<#
# controller for Remove Duplicate public folder script
# 
# updated 18.Apr.2019 to save PS transaction to log file
#>
# 
# Remove-DuplicateItems.ps1 
# calls modifed version of Remove-DuplicateItems by  Michel de Rooij,  michel@eightwone.com
#
# current version & logfile
$RDup = './Remove-DuplicateItems203rg1.ps1'
$logf = "./RemDup_"+ (get-date -format "yyyyMMdd_HHmm")+".log"



if ($Host.Name -ne "Windows PowerShell ISE Host")
{
Start-Transcript -Path $logf
}


Write-Output "### Start-- $(get-date -Format "ddd dd MMM yyyy HHmmss") -----"
if(!$Credentials) 
{
	$Credentials= Get-Credential;
}

# & $RDup -Identity boss -PublicFolders -PFStart "_Projects\2004" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2005" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2006" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2007" -Credentials $Credentials -Impersonation -Report -Force

& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2008" -Credentials $Credentials -Impersonation -Report -Force
# & $RDup -Mailbox boss -PublicFolders -PFStart "_Projects\2009" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2010" -Credentials $Credentials -Impersonation -Report -Force

& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2011" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2012" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2013" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2014" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2015" -Credentials $Credentials -Impersonation -Report -Force

& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2016" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2017" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2018" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2019" -Credentials $Credentials -Impersonation -Report -Force
& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2020" -Credentials $Credentials -Impersonation -Report -Force

& $RDup -Identity boss -PublicFolders -PFStart "_Projects\2021" -Credentials $Credentials -Impersonation -Report -Force


Write-Output " "
Write-Output "z######## END-- $(get-date -Format "ddd dd MMM yyyy HHmmss") ######"
if ($Host.Name -ne "Windows PowerShell ISE Host")
{
Stop-Transcript
}
