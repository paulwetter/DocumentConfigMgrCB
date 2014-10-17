<#
.SYNOPSIS
	Deletes packages and folders beneath a given folder structure.
.DESCRIPTION
	Deletes packages and folders beneath a given folder structure.
.PARAMETER SiteCode
    ConfigMgr Site SiteCode
    This parameter is mandatory!
    This parameter has an alias of SC.
.PARAMETER ManagementPoint
    FQDN of a ManagementPoint in this hierarchy. 
    This parameter is mandatory!
    This parameter has an alias of MP.
.PARAMETER FolderPath
    This parameter expects the path to the folder UNDER which you want to delete ALL packages and ALL folders.
    This parameter is mandatory!
    This parameter has an alias of FP.
.EXAMPLE
	PS C:\PSScript > .\delete-folderstructure.ps1 -SiteCode PR1 -ManagementPoint CM12.do.local -FolderPath "Software\HelpDesk"

    This will use PR1 as Site Code.
    This will use CM12.do.local as Management Point.
    This will use "Software\HelpDesk" as the path to the folder under which you want to delete content. ALL content beneath the folder HelpDesk and ALL packages will be deleted. USE WITH CAUTION!!!
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.LINK
	http://www.david-obrien.net
.NOTES
	NAME: delete-folderstructure.ps1
	VERSION: 1.0
	AUTHOR: David O'Brien
	LASTEDIT: June 20, 2013
    Change history:
.REMARKS
	To see the examples, type: "Get-Help .\delete-folderstructure.ps1 -examples".
	For more information, type: "Get-Help .\delete-folderstructure.ps1 -detailed".
    This script will only work with Powershell 3.0.
#>



[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]
param(
[parameter(
	Position = 1, 
	Mandatory=$true )
	] 
	[Alias("SC")]
	[ValidateNotNullOrEmpty()]
	[string]$SiteCode="",
    
    [parameter(
	Position = 2, 
	Mandatory=$true )
	] 
	[Alias("MP")]
	[ValidateNotNullOrEmpty()]
	[string]$ManagementPoint="",

    [parameter(
	Position = 3, 
	Mandatory=$true )
	] 
	[Alias("FP")]
	[ValidateNotNullOrEmpty()]
	[string]$FolderPath=""
)
<#
#Import the CM12 Powershell cmdlets
Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null
#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }
#>

$Packages = @()
$ChildFolders = @()
$Children = $null
$IDPath = @()
$GreatChildFolders = $null
$ChildFolders = $null
$Folders = $null

[array]$Folders = $FolderPath.Split("\")

$i = 0
foreach ($Folder in $Folders)
    {
        $FolderID = $null
        if ($i -eq 0)
            {
                $RootFolder = "0"
            }                
        $FolderID = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "Name = '$($Folder)' and ObjectType = '2' and ParentContainerNodeID = '$($RootFolder)'").ContainerNodeID
        $RootFolder = $FolderID
        $IDPath += $FolderID
        $i++
    }

$ParentFolder = $StartFolder = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "ContainerNodeID = '$($IDPath[-1])'").ContainerNodeID


$Children = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "ParentContainerNodeID = '$($ParentFolder)'").ContainerNodeID
$ChildFolders += $Children

foreach ($Child in $ChildFolders)
    {
        try 
            {
                $GreatChildFolders = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "ParentContainerNodeID = '$($Child)'").ContainerNodeID 
            }   
        catch [System.Management.Automation.PropertyNotFoundException] 
            {
                Write-Verbose "This was the last folder."
            }
        
        
        $ChildFolders += $GreatChildFolders
    }

Write-Host "Folders to be deleted: $($ChildFolders)"

foreach ($ChildFolder in $ChildFolders)
    {
        try 
            {
                $Packages += (Get-WmiObject -Class SMS_ObjectContainerItem -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "ContainerNodeID = '$($ChildFolder)'").InstanceKey
            }   
        catch [System.Management.Automation.PropertyNotFoundException] 
            {
                Write-Verbose "This was the last Package."
            }
    }

Write-Host "Packages to be deleted: $($Packages)"

if ((Read-Host -Prompt "Are you sure you want to delete these folders and packages? [true]") -eq $true)
    {

        foreach ($Pkg in $Packages)
            {
                try
                    {
                        $Pkg = (Get-WmiObject -Class SMS_Package -Namespace root\sms\site_$SiteCode -Filter "PackageID = '$($Pkg)'").__PATH
                        Remove-WmiObject -Path $Pkg
                    }
                catch [System.Management.Automation.PropertyNotFoundException] 
                    {
                        Write-Verbose "This was the last Package."
                    }
            }
        foreach ($Fld in $ChildFolders)
            {
                try
                    {
                        $Fld = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\site_$($SiteCode) -ComputerName $($ManagementPoint) -Filter "ContainerNodeID = '$($Fld)'").__PATH
                        Remove-WmiObject -Path $Fld -ErrorAction SilentlyContinue
                    }
                catch [System.Management.Automation.PropertyNotFoundException] 
                    {
                        Write-Verbose "This was the last folder."
                    }
            }
    }