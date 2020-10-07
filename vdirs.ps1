#Edit URL to match your environment's URL
$url = "https://mail.company.com"
#Leave this as if you are running this from Exchange Powershell console or change it to have a list of servers to apply changes to. 
$servers = get-exchangeserver
#This loop will go through each server and apply the URL defined in $url object to each virtual directory. Note, this does not include Autodiscover
#that needs to be set independant of this script. 
foreach ($server in $servers)
{
    #This is a hashtable of arrays that controls which VDirs are modified and the urls that are added to the root URL to create the correct VDIR
    $vdirs = @{
            cmd = @("owa","webservices","mapi","powershell","oab","activesync")
            url = @("owa","ews/Exchange.asmx","mapi","powershell","oab","Microsoft-Server-ActiveSync")
            }
    $i=0
    #Loops through the command and selects each variable in the hash table to generate a command that modifies the VDIR's URL 
    while($i -lt 6){
        #This populats a variable that contains the command required to change the URL for each attribute pair in the hashtable. 
        $newurl = "get-" + $vdirs.cmd[$i] + "virtualdirectory -server " + $server + " | set-" + $vdirs.cmd[$i] +"virtualdirectory -externalurl " + $url + $vdirs.url[$i] + " -internalurl " + $url + $vdirs.url[$i]
        #This runs the above generated command. Variables won't run as commands directly, so it's necessary to invoke the variable as a command. 
        Invoke-expression $newurl
        $i++
    }
}