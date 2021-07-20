
#Set the SPO URL
$tenant = Read-host "Enter your onmicrosoft tenant vanity name"
$SPOAdminCenterUrl="https://$tenant-admin.sharepoint.com/"

#Check for SPO v2 module inatallation
 $Module = Get-Module *sharepoint* -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host "Sharepoint Online module is not available"  -ForegroundColor yellow  
  $SPOConfirm= Read-Host "Do you want to install the SPO module? [Y] Yes [N] No"
  if($SPOConfirm -match "[yY]") 
  { 
   Write-host "Installing Sharepoint Online PowerShell module"
   Install-Module Microsoft.Online.SharePoint.PowerShell -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host "Sharepoint module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
   Exit
  }
 } 

#Check for EXO v2 module inatallation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host "Exchange Online PowerShell V2 module is not available"  -ForegroundColor yellow  
  $EXOConfirm= Read-Host "Do you want to install the EXO module? [Y] Yes [N] No"
  if($EXOConfirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host "EXO V2 module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
   Exit
  }
 } 

#Connect to EXO & SPO Modules

Try {
    $TestEXO = (Get-OrganizationConfig).ExchangeVersion  }
Catch
     { Write-Host "Connecting to Exchange Online" -ForegroundColor Green
Connect-ExchangeOnline }
Try {
    $TestSPO = (Get-SPOTenant).StorageQuota }
Catch
    { Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
Connect-SPOService -Url $SPOAdminCenterUrl -Credential $credential }


# Getting Value from all Office 365 Groups

$Groups = Get-UnifiedGroup
$Groups | Foreach-Object{
$Group = $_
$site=Get-SPOSite -Identity ($Group).SharePointSiteUrl -Detailed
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

Write-host "Collection complete; please email o365groups.csv & sharepoint.csv to consultant " -BackgroundColor Black -ForegroundColor Yellow

Get-PSSession | Remove-PSSession
Disconnect-SPOService
