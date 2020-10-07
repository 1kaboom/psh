#############################################################################
# Author:    Jude Perera
# Date:      06/12/2019
# Description: This tool will provide you with a comprehensive detail html of your Exchange environment and provide a summary in an HTML file. 
# 
#NOTE
# The script is custom written and the functionality should be tested prior to ensure validity and is provided “AS IS” without warranty of any kind, 
# either expressed or implied.  
#############################################################################
 

$Header = @"
<style>
table {font-family: "verdana";border-collapse: collapse;BORDER: 0.3px solid white;FONT-SIZE:  100%; padding: 1px;}
th {padding-top: 12px;padding-bottom: 12px;font-family: "calibri";text-align: left;background-color: #FF9F2C;color: white;BORDER: 0.3px solid white;FONT-SIZE:  100%; padding: 5px;}
td {white-space: nowrap;text-align: left;font-family: "calibri";BORDER: 0.3px  solid white;BORDER-COLOR: white;FONT-SIZE:  90%;padding: 5px; }
tr:nth-child(even){background-color: #EAEAEA} 
tr:nth-child(odd){background-color: #fafafa}
</style>
"@



$filenamedate = [datetime]::now.ToString('_ddMMyyyy_HHmmss')
$Attachment = "$PSScriptRoot\ExchangeComprehensiveReport$filenamedate.htm"
$ErrorActionPreference= 'silentlycontinue'
$FormatEnumerationLimit = '-1'

$start = @"
<html><div class="CSSTableGenerator">
"@

$ReportHD =@"
<h1 style="font-family:Calibri;text-align:center">Exchange Environment Comprehensive Report</h1>
<h3 style="font-family:Calibri;text-align:center">$(Get-Date)</h3>
"@


##START OF SUMMARY
$ServerHead =@"
<table>
                    <h3 style="font-family:verdana;">Server Summary</h3>
                    <tr><th>No.Exchange Servers</td>
                        <th>No.Databases</td>
                    </tr>
                    <tr><td>$((Get-ExchangeServer).count)</td>
                        <td>$((Get-MailboxDatabase).count)</td>
                    </tr>
</table> 


<br>


<table>
                    <tr><th>Name</th>
                        <th>Edition</th>
                        <th>Version</th>
                        <th>Server Role</th>
                        <th>Site</th>
                        <th>Operating System</th>
                        <th>Database Count</th>
                        <th>Mailbox Count</th>
                    </tr>

"@


$ServerBody = 
foreach ($server in ((Get-ExchangeServer)))
             {"
                    <tr>
                        <td>$($server.name)</td>
                        <td>$($server.edition)</td>
                        <td>$([string]($server.AdminDisplayVersion))</td>
                        <td>$($server.serverrole)</td>
                        <td>$($server.site.name)</td>
                        <td>
$($windows2012above = ((Get-WmiObject -ComputerName $server.name -class Win32_OperatingSystem -ErrorAction SilentlyContinue) | Where-Object{($_.Version -like "6.*") -or ($_.version -like "10.*") -and ($_.Version -notlike "6.1.*") -and ($_.version -notlike "6.0.*")}).version.count
if ($windows2012above -eq 1)
{(Get-CimInstance -ComputerName $server.name Win32_OperatingSystem -ErrorAction SilentlyContinue).caption})</td>
                        <td>$((Get-mailboxdatabase -server $server).count)</td>
                        <td>$((Get-mailbox -server $server -resultsize Unlimited).count)</td>
                    </tr>
             "}
             

$ServerEnd = 
@"
</table>
<br>
<br>
"@



##START OF DATABASE DETAILS
$DatabaseHead =@"
<table>
                    <h3 style="font-family:verdana;">Database Summary</h3>
                    <tr><th>Name</th>
                        <th>Database Size</th>
                        <th>Mounted On</th>
                        <th>Warning Quota (GB)</th>
                        <th>Prohibit Send Quota (GB)</th>
                        <th>Prohibit Send/Recv Quota (GB)</th>
                        <th>Database Copies</th>
                        <th>Last Full backup</th>
                        <th>RPC CAS FQDN</th>
                        <th>Circular Logging</th>
                        <th>EDB Path</th>
                        <th>Log Path</th>
                    </tr>

"@


$DatabaseBody = 
foreach ($database in ((Get-MailboxDatabase -status -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($database.name)</td>
                        <td>$("{0:N2}" -f ($database.databasesize.Tobytes() / 1GB))  </td>
                        <td>$($database.server.name)</td>
                        <td>$("{0:N2}" -f ($database.IssueWarningQuota.value.Tobytes() / 1GB))  </td>
                        <td>$("{0:N2}" -f ($database.ProhibitSendQuota.value.Tobytes() / 1GB))  </td>
                        <td>$("{0:N2}" -f ($database.ProhibitSendReceiveQuota.value.Tobytes() / 1GB))  </td>
                        <td>$([string]::join("<br>",($database.servers)))</td>
                        <td>$($database.LastFullBackup)</td>
                        <td>$($database.RpcClientAccessServer)</td>
                        <td>$($database.CircularLoggingEnabled)</td>
                        <td>$($database.EdbFilePath)</td>
                        <td>$($database.LogFolderPath)</td>
                        <td>
$($windows2012above = ((Get-WmiObject -ComputerName $server.name -class Win32_OperatingSystem -ErrorAction SilentlyContinue) | Where-Object{($_.Version -like "6.*") -or ($_.version -like "10.*") -and ($_.Version -notlike "6.1.*") -and ($_.version -notlike "6.0.*")}).version.count
if ($windows2012above -eq 1)
{(Get-CimInstance -ComputerName $server.name Win32_OperatingSystem -ErrorAction SilentlyContinue).caption})</td>
                        
                    </tr>
             "}
             
             #<td>$((Get-mailboxdatabase -server $server).count)</td>
                        #<td>$((Get-mailbox -server $server).count)</td>

$DatabaseEnd = 
@"
</table>
<br>
<br>
"@


##START OF DATABASE COPY DETAILS
$DBCopyHead =@"
<table>
                    <h3 style="font-family:verdana;">Database Copy Summary</h3>
                    <tr><th>Database Name</th>
                        <th>Server</th>
                        <th>Status</th>
                        <th>Content Index Status</th>
                        <th>Copy Queue Length</th>
                        <th>Replay Queue Length</th>
                    </tr>

"@


$DBCopyBody = 
foreach ($dbcopy in ((Get-MailboxDatabaseCopyStatus * -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($dbcopy.databasename)</td>
                        <td>$($dbcopy.MailboxServer)  </td>
                        <td>$($dbcopy.Status)</td>
                        <td>$($dbcopy.ContentIndexState)</td>
                        <td>$($dbcopy.CopyQueueLength)</td>
                        <td>$($dbcopy.ReplayQueueLength)</td>                       
                    </tr>
             "}
             

$DBCopyEnd = 
@"
</table>
<br>
<br>
"@
  
 
##START OF DATABASE AVAILABILITY GROUP DETAILS


$DAGHead =@"
<table>
                    <h3 style="font-family:verdana;">Database Availability Group Summary</h3>
                    <tr><th>DAG Name</th>
                        <th>DAG Members</th>
                        <th>Witness In Use</th>
                        <th>Witness Server</th>
                        <th>Alternate Witness Server</th>
                        <th>DAG IP</th>
                        <th>Primary Active Manager</th>
                        <th>Operational Servers</th>
                        <th>Exchange Version</th>
                        <th>Witness Directory</th>
                        <th>Alternate Witness Directory</th>
                   </tr>

"@


$DAGBody = 
foreach ($dag in ((Get-DatabaseAvailabilityGroup -status * -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($dag.Name)</td>
                        <td>$([string]::join("<br>",($dag.Servers)))</td>
                        <td>$($dag.WitnessShareInUse)</td>
                        <td>$($dag.WitnessServer)</td>
                        <td>$($dag.AlternateWitnessServer)</td>
                        <td>$($dag.DatabaseAvailabilityGroupIpv4Addresses)</td>
                        <td>$($dag.PrimaryActiveManager)</td>
                        <td>$([string]::join("<br>",($dag.OperationalServers)))</td>
                        <td>$($dag.ExchangeVersion)</td>
                        <td>$($dag.WitnessDirectory)</td>
                        <td>$($dag.AlternateWitnessDirectory)</td>
                    </tr>
             "}
             

$DAGEnd = 
@"
</table>
<br>
<br>
"@ 
  

  
######################################################
##START OF DATABASE AVAILABILITY GROUP NETWORK DETAILS
######################################################
$DAGNWHead =@"
<table>
                    <h3 style="font-family:verdana;">Database Availability Group Network Summary</h3>
                    <tr><th>Network Name</th>
                        <th>Subnets</th>
                        <th>Interfaces</th>
                        <th>MAPI Enabled</th>
                        <th>Replication Enabled</th>
                   </tr>

"@

$DAGNWBody = 
foreach ($dagnw in ((Get-DatabaseAvailabilityGroupNetwork -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($dagnw.Name)</td>
                        <td>$([string]::join("<br>",($dagnw.subnets)))</td>
            			<td>$([string]::join("<br>",($dagnw.Interfaces)))</td>
                        <td>$($dagnw.MapiAccessEnabled)</td>
                        <td>$($dagnw.ReplicationEnabled)</td>
                    </tr>
             "}
             
$DAGNWEnd = 
@"
</table>
<br>
<br>
"@ 



  
######################################################
##START OF OWA VIRTUAL DIRECTORY DETAILS
######################################################
$OWAVDHead =@"
<table>
                    <h3 style="font-family:verdana;">OWA Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                        <th>Forms Auth</th>
                        <th>Digest Auth</th>
                   </tr>

"@

$OWAVDBody = 
foreach ($owavd in ((Get-OwaVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($owavd.Server)</td>
                        <td>$($owavd.Name)</td>
                        <td>$($owavd.InternalURL)</td>
                        <td>$($owavd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($owavd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($owavd.ExternalAuthenticationMethods)))</td>
                        <td>$($owavd.BasicAuthentication)</td>
                        <td>$($owavd.WindowsAuthentication)</td>
                        <td>$($owavd.FormsAuthentication)</td>
                        <td>$($owavd.DigestAuthentication)</td>
                    </tr>
             "}
             
$OWAVDEnd = 
@"
</table>
<br>
<br>
"@ 
 
 
   
######################################################
##START OF ECP VIRTUAL DIRECTORY DETAILS
######################################################
$ECPVDHead =@"
<table>
                    <h3 style="font-family:verdana;">ECP Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                        <th>Forms Auth</th>
                        <th>Digest Auth</th>
                   </tr>

"@

$ECPVDBody = 
foreach ($ecpvd in ((Get-ECPVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($ecpvd.Server)</td>
                        <td>$($ecpvd.Name)</td>
                        <td>$($ecpvd.InternalURL)</td>
                        <td>$($ecpvd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($ecpvd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($ecpvd.ExternalAuthenticationMethods)))</td>
                        <td>$($ecpvd.BasicAuthentication)</td>
                        <td>$($ecpvd.WindowsAuthentication)</td>
                        <td>$($ecpvd.FormsAuthentication)</td>
                        <td>$($ecpvd.DigestAuthentication)</td>
                    </tr>
             "}
                          
$ECPVDEnd = 
@"
</table>
<br>
<br>
"@   


######################################################
##START OF OAB VIRTUAL DIRECTORY DETAILS
######################################################
$OABVDHead =@"
<table>
                    <h3 style="font-family:verdana;">OAB Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                   </tr>

"@

$OABVDBody = 
foreach ($oabvd in ((Get-OABVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($oabvd.Server)</td>
                        <td>$($oabvd.Name)</td>
                        <td>$($oabvd.InternalURL)</td>
                        <td>$($oabvd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($oabvd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($oabvd.ExternalAuthenticationMethods)))</td>
                        <td>$($oabvd.BasicAuthentication)</td>
                        <td>$($oabvd.WindowsAuthentication)</td>
                    </tr>
             "}
                          
$OABVDEnd = 
@"
</table>
<br>
<br>
"@ 
 
 
######################################################
##START OF EAS VIRTUAL DIRECTORY DETAILS
######################################################
$EASVDHead =@"
<table>
                    <h3 style="font-family:verdana;">ActiveSync Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                   </tr>

"@

$EASVDBody = 
foreach ($easvd in ((Get-ActiveSyncVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($easvd.Server)</td>
                        <td>$($easvd.Name)</td>
                        <td>$($easvd.InternalURL)</td>
                        <td>$($easvd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($easvd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($easvd.ExternalAuthenticationMethods)))</td>
                        <td>$($easvd.BasicAuthentication)</td>
                        <td>$($easvd.WindowsAuthentication)</td>
                    </tr>
             "}
                          
$EASVDEnd = 
@"
</table>
<br>
<br>
"@ 



######################################################
##START OF EWS VIRTUAL DIRECTORY DETAILS
######################################################
$EWSVDHead =@"
<table>
                    <h3 style="font-family:verdana;">EWS Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                   </tr>

"@

$EWSVDBody = 
foreach ($ewsvd in ((Get-WebServicesVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($ewsvd.Server)</td>
                        <td>$($ewsvd.Name)</td>
                        <td>$($ewsvd.InternalURL)</td>
                        <td>$($ewsvd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($ewsvd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($ewsvd.ExternalAuthenticationMethods)))</td>
                        <td>$($ewsvd.BasicAuthentication)</td>
                        <td>$($ewsvd.WindowsAuthentication)</td>
                    </tr>
             "}
                          
$EWSVDEnd = 
@"
</table>
<br>
<br>
"@ 



######################################################
##START OF MAPI VIRTUAL DIRECTORY DETAILS
######################################################
$MAPIVDHead =@"
<table>
                    <h3 style="font-family:verdana;">MAPI Virtual Directory Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>IIS Auth Methods</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                        <th>Basic Auth</th>
                        <th>Windows Auth</th>
                   </tr>

"@

$MAPIVDBody = 
foreach ($mapivd in ((Get-MapiVirtualDirectory -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($mapivd.Server)</td>
                        <td>$($mapivd.Name)</td>
                        <td>$($mapivd.InternalURL)</td>
                        <td>$($mapivd.ExternalURL)</td>
                        <td>$([string]::join("<br>",($mapivd.IISAuthenticationMethods)))</td>
                        <td>$([string]::join("<br>",($mapivd.InternalAuthenticationMethods)))</td>
            			<td>$([string]::join("<br>",($mapivd.ExternalAuthenticationMethods)))</td>
                        <td>$($mapivd.BasicAuthentication)</td>
                        <td>$($mapivd.WindowsAuthentication)</td>
                    </tr>
             "}
                          
$MAPIVDEnd = 
@"
</table>
<br>
<br>
"@ 


 
######################################################
##START OF OUTLOOK ANYWHERE DETAILS
######################################################
$OAHead =@"
<table>
                    <h3 style="font-family:verdana;">Outlook Anywhere Summary</h3>
                    <tr><th>Server</th>
                        <th>Name</th>
                        <th>Internal URL</th>
                        <th>External URL</th>
                        <th>IIS Auth Methods</th>
                        <th>Internal Auth Methods</th>
                        <th>External Auth Methods</th>
                   </tr>

"@

$OABody = 
foreach ($oa in ((Get-OutlookAnywhere -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($oa.ServerName)</td>
                        <td>$($oa.Name)</td>
                        <td>$($oa.InternalHostname)</td>
                        <td>$($oa.ExternalHostname)</td>
                        <td>$([string]::join("<br>",($oa.IISAuthenticationMethods)))</td>
                        <td>$([string]::join("<br>",($oa.InternalClientAuthenticationMethod)))</td>
            			<td>$([string]::join("<br>",($oa.ExternalClientAuthenticationMethod)))</td>
                    </tr>
             "}
                          
$OAEnd = 
@"
</table>
<br>
<br>
"@  


######################################################
##START OF CLIENT ACCESS SERVER DETAILS
######################################################
$CASHead =@"
<table>
                    <h3 style="font-family:verdana;">Client Access Server Autodiscover URL</h3>
                    <tr><th>Name</th>
                        <th>AutoDiscover URL</th>
                   </tr>

"@

$CASBody = 
foreach ($CASsvr in ((Get-ClientAccessService -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($cassvr.Name)</td>
                        <td>$($cassvr.AutoDiscoverServiceInternalUri)</td>
                     </tr>
             "}
                          
$CASEnd = 
@"
</table>
<br>
<br>
"@  



######################################################
##START OF RECEIVE CONNECTOR DETAILS
######################################################
$RCONHead =@"
<table>
                    <h3 style="font-family:verdana;">Receive Connector Summary</h3>
                   <tr>
                        <th>Server</th>
                        <th>Name</th>
                        <th>Enabled</th>
                        <th>Max Message Size (MB)</th>
                        <th>FQDN</th>
                        <th>Transport Role</th>
                        <th>Require TLS</th>
                        <th>Permissions</th>
                        <th>Authentication</th>
                        <th>Protocol Logging Enabled</th>
                        <th>Bindings</th>
                        <th>Remote IPs</th>
                   </tr>
"@

##<th>Open Relay Connectors</th>

$RCONBody = 
foreach ($RCON in ((Get-ReceiveConnector -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($RCON.Server.name)</td>
                        <td>$($RCON.Name)</td>
                        <td>$($RCON.Enabled)</td>
                        <td>$("{0:N2}" -f ($RCON.MaxMessageSize.Tobytes() / 1MB))  </td>
                        <td>$($RCON.FQDN)</td>
                        <td>$($RCON.TransportRole)</td>
                        <td>$($RCON.RequireTLS)</td>
                        <td>$($RCON.PermissionGroups -replace ",","<br>")</td>
                        <td>$($RCON.AuthMechanism -replace ",","<br>")</td>
                        <td>$($RCON.ProtocolLoggingLevel)</td>
                        <td>$([string]::join("<br>",($RCON.Bindings)))</td>
                        <td>$([string]::join("<br>",($RCON.RemoteIPRanges)))</td>
                    </tr>
             "}

##<td>$(Get-ReceiveConnector -id $RCON | Get-ADPermission | where {$_.identity -notlike "*Default*" -and $_.identity -notlike "*Client*" -and $_.user -like "NT AUTHORITY\*" -and $_.ExtendedRights -like "MS-Exch-SMTP-Accept-Any-Recipient"}.ExtendedRights)</td>

                          
$RCONEnd = 
@"
</table>
<br>
<br>
"@  



######################################################
##START OF SEND CONNECTOR DETAILS
######################################################
$SCONHead =@"
<table>
                    <h3 style="font-family:verdana;">Send Connector Summary</h3>
                   <tr>
                        <th>Name</th>
                        <th>Enabled</th>
                        <th>Max Message Size (MB)</th>
                        <th>FQDN</th>
                        <th>DNS Routing Enabled</th>
                        <th>Frontend Proxy Enabled</th>
                        <th>Ignore STARTTLS</th>
                        <th>Require TLS</th>
                        <th>Logging</th>
                        <th>Smarthosts</th>
                        <th>Source Servers</th>
                   </tr>

"@

$SCONBody = 
foreach ($SCON in ((Get-SendConnector -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($SCON.Name)</td>
                        <td>$($SCON.Enabled)</td>
                        <td>$($SCON.MaxMessageSize)</td>
                        <td>$($SCON.AddressSpaces)</td>
                        <td>$($SCON.DNSRoutingEnabled)</td>
                        <td>$($SCON.FrontendProxyEnabled)</td>
                        <td>$($SCON.IgnoreSTARTTLS)</td>
                        <td>$($SCON.RequireTLS)</td>
                        <td>$($SCON.ProtocolLoggingLevel)</td>
                        <td>$($SCON.Smarthosts)</td>
                        <td>$([string]::join("<br>",($SCON.SourceTransportServers.name)))</td>
                    </tr>
             "}
                          
$SCONEnd = 
@"
</table>
<br>
<br>
"@  



######################################################
##START OF TRANSPORT CONFIG DETAILS
######################################################
$TransportHead =@"
<table>
                    <h3 style="font-family:verdana;">Transport Configuration Summary</h3>
                   <tr>
                        <th>Max Receive Size</th>
                        <th>Max Send Size (MB)</th>
                   </tr>

"@

$TransportBody = 
foreach ($tx in ((get-TransportConfig -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($tx.MaxReceiveSize.value.tomb())</td>
                        <td>$($tx.MaxSendSize.value.tomb())</td>
                    </tr>
             "}
                          
$TransportEnd = 
@"
</table>
<br>
<br>
"@  



######################################################
### START OF OAB DETAILS
######################################################
$oabHead =@"
<table>
                    <h3 style="font-family:verdana;">Outlook Address Book Summary</h3>
                   <tr>
                        <th>Name</th>
                        <th>Version</th>
                        <th>Public Folder Distribution</th>
                        <th>Web Distribution</th>
                        <th>Address Lists</th>
                   </tr>

"@

$oabBody = 
foreach ($oab in ((Get-OfflineAddressBook -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($oab.name)</td>
                        <td>$($oab.Version)</td>
                        <td>$($oab.PublicFolderDistributionEnabled)</td>
                        <td>$($oab.WebDistributionEnabled)</td>
                        <td>$($oab.AddressLists)</td>
                    </tr>
             "}
                          
$oabEnd = 
@"
</table>
<br>
<br>
"@  


######################################################
### START OF ACCEPTED DOMAIN DETAILS
######################################################
$accdomHead =@"
<table>
                    <h3 style="font-family:verdana;">Accepted Domain Summary</h3>
                   <tr>
                        <th>Name</th>
                        <th>Domain Name</th>
                        <th>Domain Type</th>
                        <th>Default Domain</th>
                   </tr>

"@

$accdomBody = 
foreach ($accdom in ((Get-AcceptedDomain -ErrorAction SilentlyContinue)))
             {
             "
                    <tr>
                        <td>$($accdom.name)</td>
                        <td>$($accdom.DomainName)</td>
                        <td>$($accdom.DomainType)</td>
                        <td>$($accdom.Default)</td>
                    </tr>
             "}
                          
$accdomEnd = 
@"
</table>
<br>
<br>
"@  



  
          
#############################################################################################################################################################################################
#############################################################################################################################################################################################
#Format all into HTML
#############################################################################################################################################################################################
#############################################################################################################################################################################################
ConvertTo-HTML -Body "
$ReportHD $ServerHead $ServerBody $ServerEnd 
$DatabaseHead $DatabaseBody $DatabaseEnd 
$DBCopyHead $DBCopyBody $DBCopyEnd 
$DAGHead $DAGBody $DAGEnd 
$DAGNWHead $DAGNWBody $DAGNWEnd
$OWAVDHead $OWAVDBody $OWAVDEnd
$ECPVDHead $ECPVDBody $ECPVDEnd
$OABVDHead $OABVDBody $OABVDEnd
$EASVDHead $EASVDBody $EASVDEnd
$EWSVDHead $EWSVDBody $EWSVDEnd
$MAPIVDHead $MAPIVDBody $MAPIVDEnd
$OAHead $OABody $OAEnd
$CASHead $CASBody $CASEnd
$RCONHead $RCONBody $RCONEnd
$SCONHead $SCONBody $SCONEnd
$TransportHead $TransportBody $TransportEnd
$oabHead $oabBody $oabEnd
$accdomHead $accdomBody $accdomEnd
" -Title "Exchange Server Comprehensive Report" -Head $Header | Out-File $Attachment


