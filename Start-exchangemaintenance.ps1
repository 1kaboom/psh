#These functions were originally created by Adam C Brown. If you use them in scripts that are published on the internet,
#please give proper credit. 
#https://acbrownit.com

<#
    .SYNOPSIS
    This cmdlet will initiate maintenance status of an Exchange Server. 
    .DESCRIPTION
    Initiates maintenance of an Exchange server for Exchange 2013 and 2016 (2010 has a different process). 
    Script will make the following changes:
    1. Block database activation and move all databases to a different server. 
    2. Configure Hub transport services to drain email (prevents server from accepting email and completes existing transactions)
    3. Restart transport services on the server (must pass valid admin creds)
    4. Verify that all databases have been moved to a new server. If databases are still mounted, it will force a move
        with the option to skip checks on active database limits. 
    5. Set the server components to "ServerWideOffline." This disables Exchange functionality while keeping services active.
    .PARAMETER server
    Pass Exchange Server name of system being taken offline
    .PARAMETER credential
    Pass the credentials used to restart transport services. 
    .EXAMPLE
    start-exchangemaint -server server1 -credential (get-credential)
#>
function start-exchangemaint{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        $server,
        $credential)
        #Validates server name to determine if the 
        try{
            get-exchangeserver $server | Out-Null
            $isvalid = $true
        }
        catch{
            $output = $server + " is not a valid Exchange Server"
            $isvalid = $false
        }
        if ($isvalid = $true){
            set-mailboxserver $server -DatabaseCopyActivationDisabledAndMoveNow $true
            #Sleep script for 30 seconds to allow failover to succeed
            start-sleep -s 30
            Set-ServerComponentState $server -Component Hubtransport -state draining -Requester Maintenance
            "Restarting transport services to apply state change"
            Invoke-Command -ComputerName $server -ScriptBlock {get-service | where {$_.displayname -like "*transport*"} | Restart-Service} -Credential $credential
            #Sleep the script execution for 1.5 minutes each time a check against the submission queue shows messages in queue
            $queue = $server + "\submission"
            while ((get-queue -Identity $queue).messagecount -ne 0){start-sleep -m 1}
            if((get-mailboxserver $server).DatabaseCopyAutoActivationPolicy -notlike "blocked"){set-mailboxserver $server -DatabaseCopyAutoActivationPolicy blocked}
            start-sleep -s 5
            #Checks for any databases that were not successfully moved and moves them with a "manual" switchover
            $dbs = get-mailboxdatabasecopystatus -server $server|where{$_.status -like "mounted"}
            foreach ($dbname in $dbs){
                $db = $dbname.name.split("\")[0]
                $preference = (get-mailboxdatabase $db).activationpreference[1].split("[,")[1]
                Move-ActiveMailboxDatabase $db -ActivateOnServer $preference -SkipMaximumActiveDatabasesChecks -Confirm:$false
            }
            #Sets "ServerWideOffline" component to Inactive. This disables all server function without stopping services, which 
            #allows administrators to work with Exchange components and resolve problems with the server.
            Set-ServerComponentState $server -Component serverwideoffline -state inactive -Requester maintenance
            #Notification of success
            $output = $server + "has been placed in maintenance mode"
        }
}
<#
    .SYNOPSIS
    Takes a server out of maintenance mode
    .DETAILS
    This cmdlet will end maintenance mode set by start-exchangemaint cmdlet by:
    1. Disabling ServerWideOffline state
    2. Reactivating HubTransport (And restarting transport services
    3. Resetting server to allow database activation
    Databases that have been dismounted on the server should re-mount if AutoReseed/AutoDAG is enabled. 
    .PARAMETER server
    Passes the server name to take out of maintenance mode
    .PARAMETER credential
    Passes credentials used to restart transport services
    .EXAMPLE
    end-exchangemaint -server server1 -credential (get-credential)
#>

function end-exchangemaint{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        $server,
        [parameter(mandatory=$true)]
        $credential)
    #Checks the entered server name. If it is a valid mailbox server, it will return a $true result. If not, returns $false and notifies user.
    try{
            get-mailboxserver $server
            $isvalid = $true
       }
    catch{
            $isvalid = $false
            $output = $server + " is not a valid exchange server."
        }
    #Only ends maintenance mode if the server validity check returns $true. 
    if($isvalid = $true){
        #Sets the ServerWideOffline component state to Active. This sets all components to Active, enabling all server service components.  
        "Setting ServerWideOffline Component State"
        Set-ServerComponentState $server -Component serverwideoffline -state active -Requester maintenance
        #Wait 15 seconds for status change to apply
        start-sleep -s 15
        #Reactivates HubTransport status, which enables mailflow for the server. 
        "Implementing HubTransport Activation"
        Set-ServerComponentState $server -Component hubtransport -state active -requester maintenance
        #wait 15 seconds for status change to apply
        start-sleep -s 15
        #Restarts transport service (Required for the component state to apply, which will re-enable the transport service's active state)
        "Restarting Transport Services to apply state change"
        Invoke-Command -ComputerName $server -ScriptBlock {get-service | where {$_.displayname -like "*transport*"} | Restart-Service} -Credential $credential
        #Sets the database activation policy for the server to allow failback of database copies. 
        "Resetting Activation policy"
        set-mailboxserver $server -DatabaseCopyAutoActivationPolicy unrestricted -DatabaseCopyActivationDisabledAndMoveNow $false
    }
}
<#
    .SYNOPSIS
    This function is used to connect to remote PowerShell on an Exchange server. This is necessary if you
    want to manage Exchange (in any situation)
    .DESCRIPTION
    Microsoft no longer supports directly loading the Exchange PowerShell cmdlets to manage Exchange. 
    You must now connect to Exchange via remote PowerShell to use the Exchange Management Shell.
    This rule applies to local connections initiated on an Exchange server as well as remote sessions. 
    .PARAMETER server
    Enter the FQDN of a server that you would like to connect to. This usually needs to be the name
    of the server plus its domain name. For example, server1.domain.internal will connect to 
    the Server1 server on the domain.internal domain. This has to match the name of the server
    as it appears in Active Directory and on the Server itself under Computer Name. Otherwise, this
    connection will fail.
    .EXAMPLE
    connect-exchange -server server1.domain.local

#>

function connect-exchange {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        $server
        )
        #Test attribute - Remove comment character to test function against a specific server value, add comment character to function normally
        $uri = "http://" + $server + "/powershell/"
           
    try{get-mailboxserver | Out-Null}
    catch{
        try{
            $UserCredential = Get-Credential
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication kerberos -Credential $UserCredential
            Import-PSSession $Session
            }
        catch [System.Management.Automation.ParameterBindingException]{write-host "The URL used is not accessible"}
        catch [System.Management.Automation.Remoting.PSRemotingTransportException]{write-host "The authentication attempt failed - try a different user/password"}
        }
}

#Comment out the function you don't want to run. If you need to start maintenance on a server, replace the 
#Existing server name and place a # in front of the end-exchangemaint line, then run the script. 
connect-exchange -server <server>
start-exchangemaint -server <server> -credential (get-credential)
end-exchangemaint -server <server> -credential (get-credential)