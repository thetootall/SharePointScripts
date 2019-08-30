#reference 
#https://www.catapultsystems.com/blogs/pull-onedrive-for-business-usage-using-powershell/
#https://www.jijitechnologies.com/blogs/how-to-get-the-storage-used-by-office365-groups
#https://gallery.technet.microsoft.com/Connect-to-Exchange-Online-7d7365e0

$value = get-module *sharepoint*
if ($value -eq $null){
Write-host "You do not have the Sharepoint online module installed; please wait"
Start-process "https://www.microsoft.com/en-us/download/details.aspx?id=35588"
Exit
}

# Get values for input parameters:

$tenant = Read-host "Enter your onmicrosoft tenant name"
$SPOAdminCenterUrl="https://$tenant-admin.sharepoint.com/"

# Connect to SharePoint Online and Exchange Online 

Write-host "Connecting to Sharepoint Online Powershell, ensure module is installed" -BackgroundColor Black -ForegroundColor Yellow
Connect-SPOService -Url $SPOAdminCenterUrl -Credential $credential

Write-host "Connecting to Exchange Online Powershell" -BackgroundColor Black -ForegroundColor Yellow

  $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1) 
  If ($MFAExchangeModule -eq $null) 
  { 
   Write-Host "'nPlease install Exchange Online MFA Module" -ForegroundColor yellow 
   Write-Host "You can install module using below blog : `nLink `nOR you can install module directly by entering "Y"`n:
   $Confirm= "Read-Host "Are you sure you want to install module directly? [Y] Yes [N] No"
   if($Confirm -match "[yY]") 
   { 
     Write-Host Yes 
     Start-Process "iexplore.exe" "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application" 
   } 
   else 
   { 
    Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/' 
    Exit 
   } 
   $Confirmation= Read-Host "Have you installed Exchange Online MFA Module? [Y] Yes [N] No"
   if($Confirmation -match "[yY]") 
   { 
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1) 
    If ($MFAExchangeModule -eq $null) 
    { 
     Write-Host "Exchange Online MFA module is not available" -ForegroundColor red 
     Exit 
    } 
   } 
   else 
   {  
    Write-Host Exchange Online PowerShell Module is required 
    Start-Process 'https://o365reports.com/2019/04/17/connect-exchange-online-using-mfa/' 
    Exit 
   }    
  } 
   
  #Importing Exchange MFA Module 
  . "$MFAExchangeModule" 
  Connect-EXOPSSession -WarningAction SilentlyContinue 

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credential -Authentication Basic -AllowRedirection
#Import-PSSession $Session

# Getting Value from all Office 365 Groups

$Groups = Get-UnifiedGroup
$Groups | Foreach-Object{
$Group = $_
$site=Get-SPOSite -Identity $Group.SharePointSiteUrl -Detailed
      New-Object -TypeName PSObject -Property @{
      GroupName=$site.Title
      CurrentStorageInMB=$site.StorageUsageCurrent
      StorageQuotaInMB=$site.StorageQuota
      StorageQuotaWarningLevelInMB=$site.StorageQuotaWarningLevel
	Type="O365Group"
}}|select GroupName, CurrentStorageInMB, StorageQuotaInMB, StorageQuotaWarningLevelInMB,Type | export-csv o365groups.csv

# Get Value of Sharepoint sites

Get-SPOSite -Limit All -Detailed | select owner,storageusagecurrent,storagequota,storagequotawarninglevel,Url | export-csv sharepoint.csv -append

# Get Value of OneDrive sites

Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | select owner,storageusagecurrent,storagequota,storagequotawarninglevel,Url  | export-csv sharepoint.csv -append

Write-host "Collection complete; please email " -BackgroundColor Black -ForegroundColor Yellow

Get-PSSession | Remove-PSSession
Disconnect-SPOService
