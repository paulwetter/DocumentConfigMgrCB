<#
Functionality: This script creates a new Software Update Group in Microsoft System Center 2012 Configuration Manager

How does it work: create-SoftwareUpdateGroup.ps1 -UpdateGroupName $Name -KnowledgeBaseIDs $KBID -SiteCode

KnowledgeBaseID can contain comma separated KnowledgeBase IDs like 981852,16795779

Author: David O'Brien, david.obrien@sepago.de

Date: 02.12.2012

#>

param (
[string]$SiteCode,
[string]$UpdateGroupName,
[array]$KnowledgeBaseIDs,
[string]$DateUpdatesCreated,
[string]$LogFilePath,
[switch]$UseCSV,
[string]$CSVFilePath
)

Function create-Group {


[array]$CIIDs = @()

if ($UseCSV)
    {
        $KnowledgeBaseIDs = Get-Content $CSVFilePath
        foreach ($CIID in $KnowledgeBaseIDs)
            {
                $CIIDs += $CIID
            }
    }
else 
    {
        $KnowledgeBaseIDs = (gwmi -ns root\sms\site_$($SiteCode) -Class SMS_softwareupdate | where {$_.dateposted -like "$($DateUpdatesCreated)*"}).ci_id
    }
<#
foreach ($KBID in $KnowledgeBaseIDs)
    {
        $CIID = (gwmi -ns root\sms\site_$($SiteCode) -class sms_softwareupdate | where {$_.ArticleID -eq $KBID }).CI_ID
        if ($CIID -eq $null)
            {
                Write-Log "The update with KB ID $($KBID) could not be found in the database and will be ignored."
            }
        else 
            {
                $CIIDs += $CIID
            }
        
    }
#>
if (-not $UseCSV) 
    {
        foreach ($CIID in $KnowledgeBaseIDs)
            {
                $CIIDs += $CIID
                write-log "The Update with CI_ID $($CIID) has been added to the Update List"
            }
    }

$SMS_CI_LocalizedProperties = "SMS_CI_LocalizedProperties"
$class_Localization = [wmiclass]""
$class_Localization.psbase.Path ="ROOT\SMS\Site_$($SiteCode):$($SMS_CI_LocalizedProperties)"

$Localization = $class_Localization.CreateInstance()

 
$Localization.DisplayName = $UpdateGroupName
$Localization.LocaleID = 1033
 
$Description += $Localization
 
$SMSAuthorizationList = "SMS_AuthorizationList"
$class_AuthList = [wmiclass]""
$class_AuthList.psbase.Path ="ROOT\SMS\Site_$($SiteCode):$($SMSAuthorizationList)"
$AuthList = $class_AuthList.CreateInstance() 

$AuthList.Updates = $CIIDs
$AuthList.LocalizedInformation = $Description
$AuthList.Put() |Out-Null
}


function write-log([string]$info){            
    if (($loginitialized -eq $false) -and (-not (Test-Path $logfile)))
        {            
            $FileHeader > $logfile            
            $script:loginitialized = $True            
        }            
    
    $time = get-date -format G 
    $time + " " + $info | Out-File -FilePath $logfile -Append
    
}            
            
<#---------Logfile Info----------#>            
$script:logfile = "$($LogFilePath)\$($MyInvocation.MyCommand.Name)-$(get-date -format ddMMyy).log"            
$script:Seperator = @"

$("-" * 25)

"@            
$script:loginitialized = $false            
$script:FileHeader = @"
$seperator
***Application Information***
Filename:  $($MyInvocation.MyCommand.Name)
Created by:  David O'Brien
Last Modified:  $(Get-Date -Date (get-item .\$($MyInvocation.MyCommand.Name)).LastWriteTime -f dd/MM/yyyy)
"@            


create-Group

