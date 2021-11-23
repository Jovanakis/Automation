############################################################################
#This sample script is not supported under any Microsoft standard support program or service.
#This sample script is provided AS IS without warranty of any kind.
#Microsoft further disclaims all implied warranties including, without limitation, any implied
#warranties of merchantability or of fitness for a particular purpose. The entire risk arising
#out of the use or performance of the sample script and documentation remains with you. In no
#event shall Microsoft, its authors, or anyone else involved in the creation, production, or
#delivery of the scripts be liable for any damages whatsoever (including, without limitation,
#damages for loss of business profits, business interruption, loss of business information,
#or other pecuniary loss) arising out of the use of or inability to use the sample script or
#documentation, even if Microsoft has been advised of the possibility of such damages.
############################################################################

<#
.SYNOPSIS
Disable_Legacy_Protocols.ps1 disable legacy protocols (POP & IMAP,SMTP) for all users in Exchange Online 

.DESCRIPTION 
The script reads mailbox information going to a loop and disable legacy protocols (POP & IMAP & SMTP AUTH) for all users in Exchange Online 

.EXAMPLE
.\Disable_Legacy_Protocols.ps1

.NOTES
AUTHOR
Joanna Vathis, jovath@microsoft.com

COPYRIGHT
(c) 2021 Microsoft, all rights reserved

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

#>


[CmdletBinding()]
Param (
[string]$OutputDir = ".\OutputFiles",
[string]$logpath = ".\Logs",
[string]$ResultsPOPIMAP = [string]::Format(".\OutputFiles\IMAP_POP_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$ResultsSMTP = [string]::Format(".\OutputFiles\SmtpClientAuthenticationDisabled_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$SMTPAUTH = [string]::Format(".\OutputFiles\SmtpAuth_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$Results = [string]::Format(".\OutputFiles\SmtpAuthEnable_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$ResultsPOP = [string]::Format(".\OutputFiles\POP_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$ResultsIMAP = [string]::Format(".\OutputFiles\IMAP_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")),
[string]$ResultsLP = [string]::Format(".\OutputFiles\LegacyProtocols_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")))

# Version
  $Version = "2.0"

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

# =============================================================================
# Function: Check-FileExists
# Checks that file path exists
# =============================================================================

 Function Check-FileExists(){
    Param ([string]$filePath)
    
    If ((Test-Path -Path $filePath -PathType leaf) -ne $True){
        Write-Warning "Invalid file path: $($filePath)" 
        Exit
    }
}

# =============================================================================
# Function: Check-Input
# Checks that the input has the properties specified in the expected properties
# array
# =============================================================================

 Function Check-Input(){
  Param ($inputData, $expectedProperties)
  
  $props = Get-Member -MemberType NoteProperty -InputObject $inputData[0] | Select-Object -ExpandProperty Name  
   Foreach ($ep in $expectedProperties){
      If ($props -notcontains $ep){
        
        Write-Warning "Missing property $ep.
The input is not it the expected format.
Please check that the input is a CSV file that contains the following headers (case and order is ignored):
$expectedProperties"
        Exit
      }
  }  
}

# =============================================================================
# Function: Get-ErrorInfo
# =============================================================================
 
 Function Get-ErrorInfo(){
  $errStrings = @()
  if ($error.Count -ne 0)
  {
      foreach ($err in $error)
      {
         $errStrings += $err.ToString()
      }
      return [string]::join('; ', $errStrings)        
  }   
  #else
  return 'Ok'
}

#===============================================================================
# Function: DisablePOPIMAP
# Disable POP & IMAP for all mailboxes
#===============================================================================

 Function DisablePOPIMAP (){

[Array]$mbx = @()

  # Collect all Mailboxes 
    $mbx = Get-Mailbox -ResultSize Unlimited

  # Start count  
    Write-Host " "
    Write-Host "=====================================================" -ForegroundColor White
    Write-Host "       Disable POP & IMAP for all mailboxes          " -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor White
    $i = 1


  # Gathering mailbox info
     Get-CASMailbox | Where-Object {($_.PopEnabled -eq $true) -or ($_.ImapEnabled -eq $true)}
     
  # Loop and disable IMAP and POP protocol
     Foreach ($m in $mbx){

  # Set values variables
     $alias = $m.Name
    
  
  # Show message for gathering mailbox info
     Write-Host "[$(Get-Date)]: $i. Start progress disable POP & IMAP Protocol for $($alias).... " -ForegroundColor Yellow       
     Set-CASMailbox -Identity $m.alias -PopEnabled:$false -ImapEnabled:$false 
 
 
 # Progress Bar
   $i = $i+1
   Write-Host " "
 
 } 
   # Collect and export the results after disable the IMAP and POP protocols
     $output = Get-CASMailbox | Where-Object {($_.PopEnabled -eq $false) -and ($_.ImapEnabled -eq $false)} 
     $output |select Name,PopEnabled,ImapEnabled | Export-Csv -Path $ResultsPOPIMAP -NoTypeInformation -Encoding UTF8
     Write-Host ""
     Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
     Write-Host "[$(Get-Date)]: A file is exported to: $($ResultsPOPIMAP).... " -ForegroundColor Green
     Write-Host ""
}

#===============================================================================
# Function: SmtpClientAuthenticationDisabled
# Disable SMTP AUTH for all mailboxes
#===============================================================================

 Function SmtpClientAuthenticationDisabled (){

  [Array]$mbx = @()

  # Collect all Mailboxes 
    $mbx = Get-Mailbox -ResultSize Unlimited 

  # Start count  
    Write-Host " "
    Write-Host "=====================================================" -ForegroundColor White
    Write-Host "       Disable SMTP AUTH for all mailboxes           " -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor White
    $i = 1

 # Gathering mailbox info
    Get-CASMailbox | Where-Object {($_.SmtpClientAuthenticationDisabled -eq $false) -or ($_.SmtpClientAuthenticationDisabled -eq $null)}
 
 # Loop and disable SMTP protocol
    Foreach ($m in $mbx){
 
 # Set values variables
     $alias = $m.Name

 # Show message for gathering folder info
    Write-Host "[$(Get-Date)]: $i. Start progress disable Smtp Client Authentication for $($alias).... "

 # Progress Bar
   $i = $i+1
   
 }

# Collect and export the results after disable SMTP AUTH for all mailboxes
   $output = Get-CASMailbox | Where-Object {($_.SmtpClientAuthenticationDisabled -eq $true)} 
   $output |select Name,SmtpClientAuthenticationDisabled | Export-Csv -Path $Results -NoTypeInformation -Encoding UTF8
   Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
   Write-Host "[$(Get-Date)]: A file is exported to: $($Results).... " -ForegroundColor Green
}

#===============================================================================
# Function: Disable SMTP AUTH in your organization
# Disable SMTP AUTH globally in your organization
#===============================================================================

 Function SMTPAUTHGlobally (){

  # Check SMTP AUTH if its already disabled globally in your organization
    $SMTPAUTH = Get-TransportConfig 
    If($SMTPAUTH.SmtpClientAuthenticationDisabled -eq $true){
    Write-Host "[$(Get-Date)]: SMTP AUTH is already Disabled for your organization..." -ForegroundColor Yellow
    
    }else{

  # Disable SMTP AUTH globally in your organization
    If($SMTPAUTH.SmtpClientAuthenticationDisabled -eq $false){ 
    Set-TransportConfig -SmtpClientAuthenticationDisabled $true
    Write-Host "[$(Get-Date)]: SMTP AUTH is now Disabled for your organization..." -ForegroundColor Green}
 }
}

#===============================================================================
# Function: ExportSMTPAUTHList
# Export SMTP AUTH for all mailboxes
#===============================================================================

 Function ExportSMTPAUTHList (){

  # Check SMTP AUTH if its already disabled globally in your organization
    $smtp = (Get-CASMailbox -ResultSize Unlimited)
    If($smtp.SmtpClientAuthenticationDisabled -eq $null -or $smtp.SmtpClientAuthenticationDisabled -eq $true){
    Write-Host "[$(Get-Date)]: Getting Mailbox list with Smtp Client Authentication status disabled..."
    $smtp | select Name,PrimarySmtpAddress,SmtpClientAuthenticationDisabled | Export-Csv -Path $SMTPAUTH -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
    Write-Host "[$(Get-Date)]: A file is exported to: $($SMTPAUTH).... " -ForegroundColor Green
    
    }else{

  # Check SMTP AUTH for your mailboxes if its already enabled 
    If($smtp.SmtpClientAuthenticationDisabled -eq $false){ 
    Write-Host "[$(Get-Date)]: Smtp Client Authentication is already Disabled for your mailboxes..." -ForegroundColor Green
    Write-Host ""
   }
  }
 }
 
#===============================================================================
# Function: EnableSMTPAUTH
# Enable SMTP AUTH on multiple mailboxes
#===============================================================================

 Function EnableSMTPAUTH (){

  Param (
  [Array]$SMTPUser = @(),
  [Array]$ActiveUser = @())

  # Import CSV
    $ConfigFile = Import-Csv -Path ".\LegacyProtocols.csv"
 
  # Check if the file is null
    If ($ConfigFile -ne $null){

  # Enable SMTP AUTH on multiple mailboxes from CSV
     Foreach ($user in $ConfigFile) {

  # Set values variables
     $alias = $user.Name

  # Gather all the mailbox and export results
     Write-Host "[$(Get-Date)]: Start progress Enable SMTP AUTH for $($alias).... " -ForegroundColor Yellow       
     Set-CASMailbox -Identity $user.PrimarySmtpAddress -SmtpClientAuthenticationDisabled $false
 
  # Gather all the mailbox and export results
    $ActiveUser = Get-CASMailbox -Identity $user.PrimarySmtpAddress | Where-Object {($_.SmtpClientAuthenticationDisabled -eq $false)}
    $SMTPUser += $ActiveUser 

 } 
 
  # Export Results
    $SMTPUser |select Name,PrimarySmtpAddress,SmtpClientAuthenticationDisabled | Export-Csv -Path $Results -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
    Write-Host "[$(Get-Date)]: A file is exported to: $($Results).... " -ForegroundColor Green      
 }
}

#===============================================================================
# Function: DisablePOP
# Disable POP for multiple mailboxes
#===============================================================================

 Function DisablePOP (){

  Param (
  [Array]$POPUser = @(),
  [Array]$DisablePOP = @())

  # Import CSV
    $ConfigFile = Import-Csv -Path ".\LegacyProtocols.csv"

  # Start count  
    Write-Host " "
    Write-Host "==============================================================" -ForegroundColor White
    Write-Host "       Disable POP for the Mailbox from the CSV file         " -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor White
    $i = 1

 
  # Check if the file is null
    If ($ConfigFile -ne $null){

  # Enable SMTP AUTH on multiple mailboxes from CSV
     Foreach ($p in $ConfigFile) {
      If ($p.PrimarySmtpAddress -ne ""){

  # Set values variables
     $alias = $p.Name

  # Show message for gathering folder info
     Write-Host "[$(Get-Date)]: $i. Start progress disable POP Protocol for $($alias).... " -ForegroundColor Yellow       
     Set-CASMailbox -Identity $p.PrimarySmtpAddress -PopEnabled $false

  # Progress Bar
    $i = $i+1
    Write-Host " "
}
  # Gather all the mailbox and export results
    $POPUser = Get-CASMailbox -Identity $p.PrimarySmtpAddress | Where-Object {($_.PopEnabled -eq $false)}
    $DisablePOP += $POPUser 

 } 
 
  # Export Results
    $DisablePOP |select Name,PrimarySmtpAddress,PopEnabled | Export-Csv -Path $ResultsPOP -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
    Write-Host "[$(Get-Date)]: A file is exported to: $($ResultsPOP).... " -ForegroundColor Green}
}

#===============================================================================
# Function: DisableIMAP
# Disable IMAP for multiple mailboxes
#===============================================================================

 Function DisableIMAP (){

  Param (
  [Array]$IMAPUser = @(),
  [Array]$DisableIMAP = @())

  # Import CSV
    $ConfigFile = Import-Csv -Path ".\LegacyProtocols.csv"

  # Start count  
    Write-Host " "
    Write-Host "==============================================================" -ForegroundColor White
    Write-Host "       Disable IMAP for the Mailbox from the CSV file         " -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor White
    $i = 1

 
  # Check if the file is null
    If ($ConfigFile -ne $null){
    
  # Enable SMTP AUTH on multiple mailboxes from CSV
     Foreach ($im in $ConfigFile){
      If ($im.PrimarySmtpAddress -ne ""){

  # Set values variables
     $alias = $im.Name

  # Show message for gathering folder info
     Write-Host "[$(Get-Date)]: $i. Start applying the settings disable IMAP Protocol for $($alias).... " -ForegroundColor Yellow       
     Write-Host " "
     Set-CASMailbox -Identity $im.PrimarySmtpAddress -ImapEnabled $false

  # Progress Bar
    $i = $i+1
    Write-Host " "
}
 
  # Gather all the mailbox and export results
    $IMAPUser = Get-CASMailbox -Identity $im.PrimarySmtpAddress | Where-Object {($_.ImapEnabled -eq $false)}
    $DisableIMAP += $IMAPUser 

 } 
 
  # Export Results
    $DisableIMAP |select Name,PrimarySmtpAddress,ImapEnabled | Export-Csv -Path $ResultsIMAP -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
    Write-Host "[$(Get-Date)]: A file is exported to: $($ResultsIMAP).... " -ForegroundColor Green
  }
}

#===============================================================================
# Function: DisableIMAP
# Disable IMAP for multiple mailboxes
#===============================================================================

 Function LegacyProtocols (){

[Array]$mbx = @()
[Array]$output = @()

# Collect all Mailboxes 
   $mbx = Get-Mailbox -ResultSize Unlimited

# Export Legacy Protocols Status
  $output = Get-CASMailbox | Where-Object {($_.PopEnabled -eq $true) -and ($_.ImapEnabled -eq $true) -and ($_.SmtpClientAuthenticationDisabled -eq $true)} 
  $output |select Name,PopEnabled,ImapEnabled | Export-Csv -Path $ResultsLP -NoTypeInformation -Encoding UTF8
  Write-Host ""
  Write-Host "[$(Get-Date)]: Exporting the file results... " -ForegroundColor Yellow
  Write-Host "[$(Get-Date)]: A file is exported to: $($ResultsLP).... " -ForegroundColor Green
}



#===============================================================================
# String: Menu
# Create a menu with options
#===============================================================================

[string] $menu = @'
   ===========================================================================
        Disable Legacy Protocols in Office 365 - Version 2.0
   ===========================================================================
     
     1. Disable POP & IMAP Protocol for Bulk Mailboxes              
     2. Disable SMTP AUTH for all Mailboxes  
     3. Disable SMTP AUTH globally in your organization
     4. Export SMTP AUTH for all mailboxes
     5. Enable SMTP AUTH for multiple mailboxes (CSV)
     6. Disable POP for multiple mailboxes (CSV)
     7. Disable IMAP for multiple mailboxes (CSV)
     8. Export Legacy Protocols (POP,IMAP,SMTP) information's
    10. Exit Menu and logout from Menu 
                                       
   ===========================================================================
   Select an option from the menu...(1-10)
'@  



# ===============================================================================================================================================================
#                                                             SCRIPT BODY - MAIN CODE                                                                             
# ===============================================================================================================================================================

 # Check if Log folder exist, if not create OutputFiles folder
Try {
    # Check if output directory exists and create it if doesn't
	If (!(Test-Path -Path $OutputDir)){ 
      New-Item $OutputDir -Type directory | Out-Null}
}
catch [System.Exception]{
	  Write-Error $_.Exception.Message
	  Exit	
}

# Check if Log folder exist, if not create Log folder
Try {
    # Check if output directory exists and create it if doesn't
	If (!(Test-Path -Path $logpath)){ 
      New-Item $logpath -Type directory | Out-Null}
}
catch [System.Exception]{
	  Write-Error $_.Exception.Message
	  Exit	
}

# Set Variables
  $logname = Get-Date -Format "MM.dd.yyyy_HH.mm"
  $logpath = ".\Logs\O365Logs_$logname.log"
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

Do {       
 If ($opt -ne $null) {
 Write-Host ""
 Write-Host "Last Command:" $opt -ForegroundColor Cyan}
 Write-Host ""
 $opt = Read-Host $menu

 Switch ($opt){
           # Disable POP & IMAP Protocol for Bulk Mailboxes  
        1    {DisablePOPIMAP}
         
           # Disable SMTP AUTH for all Mailboxes  
        2    {SmtpClientAuthenticationDisabled}

           # Disable SMTP AUTH globally in your organization
        3    {SMTPAUTHGlobally}

           # Export SMTP AUTH for all mailboxes
        4    {ExportSMTPAUTHList}

           # Enable SMTP AUTH for multiple mailboxes (CSV)
        5    {EnableSMTPAUTH} 

           # Disable POP for multiple mailboxes (CSV)
        6    {DisablePOP}
            
           # Disable IMAP for multiple mailboxes (CSV)      
        7    {DisableIMAP}   
           
           # Disable IMAP for multiple mailboxes (CSV)      
        8    {LegacyProtocols}   

           # Exit Menu and logout from Menu 
       10    {If ($choose -ne -10){ 
              Write-Host "   Exiting from Menu....." -ForegroundColor Cyan}{}}}
             
          # Execute cmdlets from the script in order
}            While ($opt -ne 10)


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