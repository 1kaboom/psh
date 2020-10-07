#################################################################################
#
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages.
#
################################################################################# 
#
[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$UserPrincipalName
)

$FormatEnumerationLimit=-1
$DateTime = Get-Date -Format yyyyMMdd_hhmmss
$OutputPath = "c:\temp\Get-UserData"
If(!(test-path $OutputPath)) {New-Item -ItemType Directory -Force -Path $OutputPath}
Start-Transcript -path "$OutputPath\Get-UserData-$UserPrincipalName-$DateTime.txt" -IncludeInvocationHeader
Write-Host "ORIGINATING SCRIPT: $ScriptName - START."
Write-Host '----------------------------------------------------'
$Script = (Get-Content $MyInvocation.MyCommand.Path)
Foreach ($line in $Script) {Write-Host ">> $Line"}
Write-Host '----------------------------------------------------'
Write-Host "ORIGINATING SCRIPT:  $ScriptName - END."
Write-Host ''
Write-Host '***SCRIPT OUTPUT BEGINS***'
Write-Host ''
$List = @("Get-PSSession | FL", `
    "Get-User $UserPrincipalName | FL", `
    "Get-MsolUser -UserPrincipalName $UserPrincipalName | FL", `
    "Get-Recipient $UserPrincipalName | FL", `
    "Get-MailUser $UserPrincipalName | FL", `
    "Get-Mailbox $UserPrincipalName | FL", `
    "Get-Mailbox $UserPrincipalName -SoftDeletedMailbox | FL", `
    "Get-Mailbox $UserPrincipalName -InactiveMailboxOnly | FL", `
    "Get-Mailbox $UserPrincipalName -Archive | FL", `
    "Get-CASMailbox $UserPrincipalName | FL", `
    "Get-MailboxStatistics $UserPrincipalName | FL", `
    "Get-MailboxStatistics $UserPrincipalName -Archive | FL", `
    "Get-MailboxFolderStatistics $UserPrincipalName | FL", `
    "Get-MailboxFolderStatistics $UserPrincipalName -Archive | FL", `
    "Get-MailboxPermission $UserPrincipalName | FL", `
    "Get-MailboxFolderPermission $UserPrincipalName | FL", `
    "Get-MailboxFolderPermission $UserPrincipalName`:\Inbox | FL", `
    "Get-MailboxFolderPermission $UserPrincipalName`:\Calendar | FL", `
    "Get-CalendarProcessing $UserPrincipalName | FL", `
	"Get-OwaMailboxPolicy | FL")
Foreach ($Item in $List) {
    Write-Host $Item
    Write-Host '**********************************************'
    Invoke-Expression $Item
}
$UserRetentionPolicy = (Get-Mailbox $UserPrincipalName).RetentionPolicy
Write-Host "$UserPrincipalName is using the retention policy: $UserRetentionPolicy"
Write-Host '**********************************************'
Get-RetentionPolicy $UserRetentionPolicy | FL
$RetentionPolicyTagList = (Get-RetentionPolicy $UserRetentionPolicy).RetentionPolicyTagLinks
Foreach ($RetentionPolicyTag in $RetentionPolicyTagList) {
    Write-Host "Retention tag: $RetentionPolicyTag"
    Write-Host '**********************************************'
    Get-RetentionPolicyTag $RetentionPolicyTag | FL
}
Stop-Transcript
Invoke-Expression "$OutputPath\Get-UserData-$UserPrincipalName-$DateTime.txt"