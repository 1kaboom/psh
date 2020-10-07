#####################################################################
#Author-SANDEEP DHONDE
#Version 1.0
#Created - 02/03/2019
#Ver-1.1 - Removed "Temp" folder prerequisite
#DESCRIPTION
#This function helps import bulk users in Distribution List. Displays the current count of total members in DL and then final count post importing. Errors encountered while adding members are logged to filename created as "DLName Date.txt" to "C:\Temp" folder.
#####################################################################

$ErrorActionPreference = 'Stop'
Function CheckDL{
                 [CmdletBinding()]
                  param(
                        [String]$DistributionGroup
                   )
                   Try {Get-DistributionGroup $DistributionGroup
                   } Catch{ Write-Host $Error[0].ToString()
                      }
          }

Function Import-DLMembers{
<#
.Notes
#####################################################################
Author-SANDEEP DHONDE
Version 1.0
Created - 02/03/2019
#####################################################################
Prerequisites:
SMTP relay allowed from the machine where script is being executed, if you want logs on email
#####################################################################
.SYNOPSIS
Import bulk users in Distribution Group or List

.DESCRIPTION
This function helps import bulk users in Distribution List. Displays the current count of total members in DL and then final count post importing. Errors encountered while adding members are logged to filename created as "DLName Date.txt" to "C:\Temp" folder.

 .EXAMPLE

 Import-DLMembers -DLName "DL LAB USERS" -FilePath C:\Temp\Users.txt

Current DL Members:998
*****************************************
Final DL Members!:1650
*****************************************
Check Error Logfile at C:\Temp\DL LAB USERS 2019-02-19.txt
*****************************************

The above example imports users specified in "Users.txt" in "DL LAB USERS" distribution list and displays the path of error logfile logged during the import.

 .EXAMPLE

 Import-DLMembers -DLName "DL LAB USERS" -FilePath C:\Temp\Users.txt -SMTPServerName smtp.lab.com -SenderEmailID Test1@lab.com -RecipientEmailID Test2@lab.com

Current DL Members:998
*****************************************
Final DL Members!:1650
*****************************************
Check Error Logfile at C:\Temp\DL LAB USERS 2019-02-19.txt
*****************************************
Sending Error Logs on Test2@lab.com
*****************************************

The above example imports users in Distribution List and sends the error logfile on emailaddress specified.
#>
                  [CmdletBinding()]
                  param(
                         [Parameter(Mandatory=$True,
                         HelpMessage='E.g: DL India User')]
                         [String]$DLName,

                         [Parameter(Mandatory=$True,
                         ValueFromPipeline=$True,
                         HelpMessage='E.g: C:\Temp\User.txt')]
                         [String]$FilePath,

                         [Parameter(Mandatory=$False,
                         HelpMessage='E.g: smtp.company.com')]
                         [String]$SMTPServerName,

                         [Parameter(Mandatory=$False)]
                         [String]$SenderEmailID,

                         [Parameter(Mandatory=$False)]
                         [String]$RecipientEmailID
                   )
                   
                                                BEGIN{
                                                      if (!(Test-path C:\Temp)){
                                                          mkdir C:\Temp | Out-Null
                                                       }
                                                }
                                                
                                                PROCESS{ If (CheckDL $DLName){
                                                                              Try{$Users = Get-Content $FilePath} Catch{ Write-Host $Error[0].ToString(); break}
                                                                              $DLMemCount = (Get-DistributionGroupMember $DLName).count
                                                                              Write-Host "Current DL Members: $DLMemCount" -ForegroundColor Green
                                                                              $FileName = "$DLName" + " " + (Get-Date -Format "yyyy-MM-dd") + ".txt"
                                                                              $i = 0
                                                                              foreach($User in $Users){
                                                                              $i++
                                                                              $Percent = [math]::Round($i/$Users.count*100)
                                                                              Write-Progress -Activity "Adding $User in $DLName" -Status "$Percent % Complete" -PercentComplete $Percent
                                                                              Try{
                                                                                   Add-DistributionGroupMember -Identity $DLName -Member $User -ErrorAction stop
                                                                                } Catch{(Get-Date).ToString("dd/MM/yyyy %H:mm:ss") +" "+ $Error[0].ToString() | Out-File c:\temp\$FileName -Append
                                                                                   }
                                                                               }
                                                                             } Else {Break}


                                                }
                                                END { Write-Progress -Activity "Adding $User in $DLName" -Status "Completed" -Completed
                                                      $DLPostCount = (Get-DistributionGroupMember $DLName).count
                                                      Write-Host "*******************************************************************************************" -ForegroundColor Cyan
                                                      Write-Host "Final DL Members!: $DLPostCount" -ForegroundColor Green
                                                      Write-Host "*******************************************************************************************" -ForegroundColor Cyan
                                                      Write-Host "Check Error Logfile at C:\Temp\$FileName" -ForegroundColor Green
                                                      If($SMTPServerName){
                                                                           Write-Host "*******************************************************************************************" -ForegroundColor Cyan
                                                                           Write-Host "Sending Error Logs on $RecipientEmailID" -ForegroundColor Green
                                                                           Write-Host "*******************************************************************************************" -ForegroundColor Cyan
                                                                           $Attachment = "C:\temp\$FileName"
                                                                         Try{ Send-MailMessage -To $RecipientEmailID -From $SenderEmailID -Subject "Bulk DLImport Error Logs" -SmtpServer $SMTPServerName -Attachments $Attachment
                                                                         } Catch{$Err =$Error[0].ToString()
                                                                                 Switch -Regex ($Err)
                                                                                 {
                                                                                 "Unable to connect to the remote server" {Write-host "Unable to connect to $SMTPServerName on port 25. Logs sending failed!"}
                                                                                 "The remote name could not be resolved"  {Write-Host "Unable to resolve $SMTPServerName, check if the smtpname is resolvable. Logs sending failed!"}
                                                                                 "Default"                                {Write-Host $Err}
                                                                                 }
                                                                            }
                                                           }Else {Break}
                                                }
         }
