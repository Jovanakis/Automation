################################################################################################
#This sample script is not supported under any Microsoft standard support program or service.  #
#This sample script is provided AS IS without warranty of any kind.                            #  
#Microsoft further disclaims all implied warranties including, without limitation, any implied #
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising #
#out of the use or performance of the sample script and documentation remains with you. In no  #
#event shall Microsoft, its authors, or anyone else involved in the creation, production, or   #
#delivery of the scripts be liable for any damages whatsoever (including, without limitation,  #
#damages for loss of business profits, business interruption, loss of business information,    #
#or other pecuniary loss) arising out of the use of or inability to use the sample script or   #
#documentation, even if Microsoft has been advised of the possibility of such damages.         #
################################################################################################

<#
.SYNOPSIS
Get-MBXSFolderPermission.ps1  

.DESCRIPTION 
The script gather permissions from folder mailboxes report 

.EXAMPLE
.\Get-MBXSFolderPermission.ps1

.NOTES
AUTHOR
Joanna Vathis, jovath@microsoft.com

COPYRIGHT
(c) 2021 Microsoft, all rights reserved

CHANGE LOG
v1.0, 2021 - Initial version.

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

#>

#===============================================================================
# Function: Create-EXOSession
# Connect to Exchange Online
#===============================================================================

 Function Create-EXOSession (){

 Param (
 [String]$EXO_Module = (Get-PSSession))

  # Check if Pssesion Open and if its open, remove it
    Get-PSSession | Remove-PSSession
 
 # Check if the O365 Username is not empty as viriable cls

   If(($EXO_Module -eq $Null) -or ($EXO_Module -eq "")){
    Connect-ExchangeOnline -ShowBanner:$false -WarningAction:SilentlyContinue -ErrorVariable:ConnectErrors -PSSessionOption $RPSProxySetting | Out-Null
    Write-Host "$(Get-Date) Connecting to Exchange Online..." -ForegroundColor Green
      
 }else{ 

    If(($EXO_Module -ne $Null) -or ($EXO_Module -ne "")){
    Write-Host "$(Get-Date) Connection to Exchange is Active..." -ForegroundColor Green
    
  }
 }
}

# Set Variables
  $logname = Get-Date -Format "MM.dd.yyyy_HH.mm"
  $logpath = ".\O365Logs_$logname.log"
  $Error.Clear()

# Start Transcript
  Start-Transcript -Path $logpath -Append

# START OF SCRIPT
   $ScriptProcessStartTime = Get-Date -Format "dd-MMM-yyyy HH:mm"
    Write-Host "***********************************************************"
    Write-Host " Script started processing at:" $ScriptProcessStartTime -ForegroundColor Yellow
    $ScriptTimer = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "***********************************************************"
    Write-Host ""

# Create PSSession to Exchange
   $session = Get-PSSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    If(($session -ne $null) -and ($session.State -eq 'Opened') -and ($session.Availability -eq 'Available') -and
       ($session.ConfigurationName -eq 'Microsoft.Exchange') -and ($session.ComputerName -eq "outlook.office365.com")){
        Write-Host "[$(Get-Date)]: Reusing opened PSSession to Exchange Online...... " -ForegroundColor Cyan
     
  

} elseif (($session -eq $null) -or ($session.State -eq 'Broken')){

# Create a new Session to Office 365
  Write-Host "Create a new PSSession to Exchange Online......." -ForegroundColor Yellow
  Write-Host ""
  Create-EXOSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null     
}


# Set Parameters
  [Array]$folderpermissions = @()
  [Array]$alias = @()
  [string]$MBXSFolderPermission = [string]::Format(".\Report_Mailbox_Folders_Statistics_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss"))

 # Set Variables
   $path = $null

  # Collect all the UserMailboxes 
    Write-Host "[$(Get-Date)]: Gathering Information of All mailboxes in the organization...." 
    $mailboxes = @(Get-Mailbox -ResultSize Unlimited -Filter "((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'EquipmentMailbox') -or
             (RecipientTypeDetails -eq 'RoomMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox') -and (Name -notlike 'DiscoverySearchMailbox*'))")

  # Start count mailboxes 
    $i = 1
    
    # Show message for gathering folder info
        Write-Host "[$(Get-Date)]: Collecting mailbox info please wait..." 
	    Write-Host "[$(Get-Date)]: Processing ......." 
       
            
    # Start the loop and gathering the statistics 
       If ($mailboxes –ne $null){
        Foreach($mbx in $mailboxes){
     
    # Show message for gathering folder info
       Write-Host "$($i). Gathering folder heirarchy information for mailbox : $($mbx.UserPrincipalName)    " -ForegroundColor Yellow
       $folders = Get-MailboxFolderStatistics -Identity $mbx.UserPrincipalName
       $folderpermissions += Get-MailboxFolderpermission -Identity $mbx.UserPrincipalName | select Identity,FolderName,User,AccessRights
       
       # Start the loop and gathering the statistics              
          Foreach($folder in $folders){
               
       # Gathering folder permissions information for the folder                    
         $folderpath = (($folder.Identity).ToString()).Split("\")
         $folderpath[0] = $mbx.UserPrincipalName +":"
         $path = [string]::Join("\",$folderpath)
         $folderpermissions += Get-MailboxFolderPermission $path -ErrorAction SilentlyContinue | Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}`
        | select Identity,FolderName,User,AccessRights
  }
     # Progress Bar
       $i = $i+1
}
    # Export Mbx Folder statistics 
       $folderpermissions | Out-GridView
       $folderpermissions | Export-csv $MBXSFolderPermission -NoTypeInformation -Encoding UTF8
       
   Write-Host ("")
   Write-Host "Output saved to:  $($MBXSFolderPermission)" -ErrorAction SilentlyContinue 
   
      
}

# END OF SCRIPT
   $ScriptProcessEndTime = Get-Date -Format "dd-MMM-yyyy HH:mm"
   Write-Host "*************************************************************"
    write-Host "Script started processing at:" $ScriptProcessStartTime -ForegroundColor Yellow 
    write-Host "Script completed at:" $ScriptProcessEndTime -ForegroundColor Yellow
    write-host "Total Processing Time: $($ScriptTimer.Elapsed.ToString())" -BackgroundColor White -ForegroundColor DarkBlue
   Write-Host "*************************************************************"
   Write-Host " "

# Stop the Transcript
If(-not $Error){
   Stop-Transcript -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
   
} else {
	Write-Host $error
	Stop-Transcript -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}# End of the Code
