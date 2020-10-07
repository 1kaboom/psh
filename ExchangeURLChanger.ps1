# Script:    ExchangeURLChanger.ps1 
# Purpose:   The tool will take the internal and external domain URL's and change 
#            all the virtual directory URL's in the environment 
# Author:    Jude Perera  
# Date:      September 2019  
# The script is custom written and the functionality should be tested prior to ensure validity and is provided “AS IS” without warranty of any kind, either expressed or implied. 
  
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"

<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Exchange Virtual Directory URL Changer" Height="262" Width="525">
    <Grid Margin="0,0,0,-1">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="70*"/>
            <ColumnDefinition Width="15*"/>
            <ColumnDefinition Width="432*"/>
        </Grid.ColumnDefinitions>
        <Label Name="label" Content="Input your public domain (contoso.com):" HorizontalAlignment="Left" Margin="38,26,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>
        <Label Name="label1" Content="Input your internal FQDN (mail.contoso.com):" HorizontalAlignment="Left" Margin="38,57,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>
        <Label Name="label1_Copy" Content="Input your external FQDN (mail.contoso.com):" HorizontalAlignment="Left" Margin="38,88,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>
        <TextBox Name="textBox" HorizontalAlignment="Left" Height="23" Margin="207,29,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="198" Grid.Column="2"/>
        <TextBox Name="textBox1" HorizontalAlignment="Left" Height="23" Margin="207,59,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="198" Grid.Column="2"/>
        <TextBox Name="textBox2" HorizontalAlignment="Left" Height="23" Margin="207,88,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="198" Grid.Column="2"/>
        <Button Name="change" Content="Change URLs" HorizontalAlignment="Left" Margin="207,176,0,0" VerticalAlignment="Top" Width="84" Grid.Column="2" Height="29" FontWeight="Bold"/>
        <Label Name="label3" Content="Select External Client Authentication Method for &#xD;&#xA;Outlook Anywhere" HorizontalAlignment="Left" Margin="38,119,0,0" VerticalAlignment="Top" Grid.ColumnSpan="3" Width="254" FontSize="11"/>
        <Button Name="button1" Content="Exit" Grid.Column="2" HorizontalAlignment="Left" Margin="321,176,0,0" VerticalAlignment="Top" Width="84" Height="29" FontWeight="Bold"/>
        <ComboBox Name="Combobox" Grid.Column="2" HorizontalAlignment="Left" Margin="207,123,0,0" VerticalAlignment="Top" Width="84">
            <ComboBoxItem>Basic</ComboBoxItem>
            <ComboBoxItem IsSelected="True">NTLM</ComboBoxItem>
            <ComboBoxItem>Negotiate</ComboBoxItem>
        </ComboBox>
        <Label Name="version" Content="v1.0" HorizontalAlignment="Left" Margin="4,208,0,0" VerticalAlignment="Top" FontSize="7"/>
    </Grid>
</Window>
"@
$wshell = New-Object -ComObject Wscript.Shell
$i = 1
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader"; exit}

# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#Assign event
$change.Add_Click(
    {
        
        $Domain = $textBox.Text;
        $IntDomain = $textBox1.Text;
        $ExtDomain = $textBox2.Text;
        $AuthType = $ComboBox.Text;

        #Connect to Exchange PowerShell remoting
        #$UserCredential = Get-Credential
        #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$env:COMPUTERNAME/PowerShell/ -Authentication Kerberos -Credential $UserCredential 
        #Import-PSSession $Session -DisableNameChecking -AllowClobber
    
        #Set OWA URLs
        Get-OwaVirtualDirectory | Set-OwaVirtualDirectory -InternalUrl "https://$IntDomain/owa" -ExternalUrl "https://$ExtDomain/owa"  -WarningAction silentlyContinue
        $i = 12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Outlook Web App URLs configured" -PercentComplete $i;

        #Set ECP URLs
        Get-EcpVirtualDirectory | Set-EcpVirtualDirectory -InternalUrl "https://$IntDomain/ecp" -ExternalUrl "https://$ExtDomain/ecp"-WarningAction silentlyContinue
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Exchange Control Panel URLs configured" -PercentComplete $i;

        #Set OAB URLs
        Get-OabVirtualDirectory | Set-OabVirtualDirectory -InternalUrl "https://$IntDomain/oab" -ExternalUrl "https://$ExtDomain/oab"
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Outlook Address Book URLs configured"  -PercentComplete $i;

        #Set Web Services URLs
        Get-WebServicesVirtualDirectory | Set-WebServicesVirtualDirectory -InternalUrl "https://$IntDomain/ews/exchange.asmx" –ExternalUrl "https://$ExtDomain/ews/exchange.asmx"
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Exchange Web Services URLs configured"  -PercentComplete $i;

        #Set MAPI URLs
        Get-MapiVirtualDirectory | Set-MapiVirtualDirectory -InternalUrl "https://$IntDomain/mapi" -ExternalUrl "https://$ExtDomain/mapi"
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Mapi URLs configured"  -PercentComplete $i;

        #Set ActiveSync URLs
        Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -InternalUrl "https://$IntDomain/Microsoft-Server-ActiveSync" -ExternalUrl "https://$ExtDomain/Microsoft-Server-ActiveSync"
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "ActiveSync URLs configured"  -PercentComplete $i;

        #Set AutoDiscover Service URLs
        Get-ClientAccessService | Set-ClientAccessService -AutoDiscoverServiceInternalUri "https://autodiscover.$Domain/Autodiscover/Autodiscover.xml" 
        $i = $i+12.5;
        Write-Progress -Activity "URL Change in Progress" -status "Autodiscover Service Internal Uri configured"  -PercentComplete $i;

        #Set Outlook Anyware URLs
        Get-OutlookAnywhere | Set-OutlookAnywhere -InternalHostname "$IntDomain" -ExternalHostname "$ExtDomain" -ExternalClientsRequireSsl:$false -InternalClientsRequireSsl:$false -SSLOffloading:$true -ExternalClientAuthenticationMethod $AuthType
        $i = $i+12.5;
        Write-Progress -Activity "URL Change is completed" -status "100% completed"  -PercentComplete $i;

        Write-Host "URL change completed, click exit to close" -ForegroundColor Green    

    }
)

$button1.Add_Click({$form.Close()})


#Show Form
$Form.ShowDialog() | out-null