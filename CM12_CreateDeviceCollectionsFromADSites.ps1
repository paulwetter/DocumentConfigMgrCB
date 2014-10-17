Function Get-SiteCode
{
    $wqlQuery = “SELECT * FROM SMS_ProviderLocation”
    $a = Get-WmiObject -Query $wqlQuery -Namespace “root\sms” #-ComputerName $SMSProvider
    $a | ForEach-Object {
        if($_.ProviderForLocalSite)
            {
                $script:SiteCode = $_.SiteCode
            }
    }
return $SiteCode
}

Function Import-CM12Module
{
#Import the CM12 Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
    {
        Write-Output "$(Get-Date):   ConfigMgr module has not been imported yet, will import it now."
        Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH -parent) ConfigurationManager.psd1) | Out-Null
    }

#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }
}

Function Create-CollectionFolder{	$script:FolderName = "All Systems in AD Sites"    if (-not (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\sms\site_$SiteCode -Filter "ObjectType = '5000' and Name = '$($FolderName)'"))        {            Write-Output "$(Get-Date):   Folder does not exist, creating it."            $CollectionFolderArgs = @{	        Name = $FolderName;	        ObjectType = "5000"; 		# 5000 is for Collection_Device, 5001 is for Collection_User	        ParentContainerNodeid = "0" # ParentContainerNodeID is '0' if Folder underneath root folder, otherwise ParentContainerNodeID needs to be evaluated	        }	        Set-WmiInstance -Class SMS_ObjectContainerNode -arguments $CollectionFolderArgs -namespace "root\SMS\Site_$SiteCode" | Out-Null        }}

Function Move-CMCollection
{
[int]$script:TargetFolder = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\sms\site_$SiteCode -Filter "ObjectType = '5000' and Name = '$($FolderName)'").ContainerNodeID
$Parameters = ([wmiclass]"root\SMS\Site_$($SiteCode):SMS_ObjectContainerItem").psbase.GetMethodParameters("MoveMembers")
$Parameters.ObjectType = 5000
$Parameters.ContainerNodeID = 0
$Parameters.TargetContainerNodeID = $TargetFolder
$Parameters.InstanceKeys = $Coll.CollectionID

try {
        $Output = ([wmiclass]"root\SMS\Site_$($SiteCode):SMS_ObjectContainerItem").psbase.InvokeMethod("MoveMembers",$Parameters,$null)
        if ($Output.ReturnValue -eq "0")
            {
                Write-Output "Collection successfully moved to Folder $($FolderName)."
            }
    }
catch [Exception]
    {
        Write-Error -Message "Something went wrong."
    }
}


Get-SiteCode
Import-CM12Module
Create-CollectionFolder

$SiteDescription=@{}
$SitesDN="LDAP://CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")

foreach ($Site in $([adsi] $SitesDN).psbase.children)
    {
        if ($Site.objectClass -eq "site")
            {
                $Name = ([string]$Site.cn).toUpper()
                $SiteDescription[$Name] = $Site.Description
            }
    }


foreach ($ADSite in $SiteDescription.Keys)
    {
        if (-not (Get-CMDeviceCollection -Name "All Systems in AD Site $($ADSite)"))
            {
                Write-Output "$(Get-Date):   Creating Device Collection `"All Systems in AD Site $($ADSite)`""
                $Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7
                $Coll = New-CMDeviceCollection -LimitingCollectionId SMSDM003 -Name "All Systems in AD Site $($ADSite)" -RefreshSchedule $Schedule -Comment "All Systems in AD Site $($ADSite)"
                Add-CMDeviceCollectionQueryMembershipRule -CollectionId $Coll.CollectionID -RuleName "All Systems in AD Site $($ADSite)" -QueryExpression "SELECT SMS_R_System.Name, SMS_R_System.ADSiteName FROM SMS_R_System where SMS_R_System.ADSiteName = `"$ADSite`""
                Move-CMCollection
            }
        else
            {
                Write-Output "$(Get-Date):   Device Collection `"All Systems in AD Site $($ADSite)`" already exists, skipping it."
            }
    }