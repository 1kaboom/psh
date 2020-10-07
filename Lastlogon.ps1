$users = get-mailbox -organizationalunit "attribute here" -resultsize unlimited
$userarray = @()
foreach ($user in $users)
{
$MailUser = $user.UserPrincipalName
$stats= Get-MailboxStatistics $MailUser
$datetime = $stats.LastLogonTime
$date, $time = $datetime -split(' ')
$Maildetails = New-Object -TypeName PSObject -Property @{
DisplayName = $stats.DisplayName
ItemCount = $stats.ItemCount
MailboxSize = $stats.TotalItemSize
LastLogonDate = $date
LastLogonTime = $time
Email = $MailUser

}

$userarray += $Maildetails
} 

$userarray | Export-Csv -Path C:\Users1.csv