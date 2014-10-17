param
(
$SMSProvider,
$BootImageName,
$BootMediaPath,
$MountPath,
$AdaptivaClientSources
)

<#
Execute like this:
New-AdaptivaBootImage.ps1 -SMSProvider %NameOfSMSProvider% -BootImageName %NameOfExistingPE% -BootMediaPath %PathToCreatedISO% -MountPath %PathToMountFolder% -AdaptivaClientSources %PathToFolderWith OneSiteDownloader%


$BootImageName = "WinPE5PwithDART" # Boot Image you like to copy and have the Adaptiva Binaries copied to
$BootMediaPath = "\\dc01\sources\OSD\AdaptivaMedia.iso" # The ISO you manually created beforehand
$MountPath = "C:\mount" # The path where you would like the WIM to be mounted
$AdaptivaClientSources = "\\dc01\sources\Software\AdaptivaClient" #The path to your OneSiteDownloader.exe and OneSiteDownloader64.exe

Author: David O'Brien, obrien.david@outlook.com
#>

Function Get-SiteCode
{
    $wqlQuery = “SELECT * FROM SMS_ProviderLocation”
    $a = Get-WmiObject -Query $wqlQuery -Namespace “root\sms” -ComputerName $SMSProvider
    $a | ForEach-Object {
        if($_.ProviderForLocalSite)
            {
                $script:SiteCode = $_.SiteCode
            }
    }
    return $SiteCode
}


Get-SiteCode
#Import the CM12 Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
    {
        Write-Verbose "$(Get-Date):   CM12 module has not been imported yet, will import it now."
        Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null
    }

if (-not (Test-Path $MountPath))
    {
        Write-Output "Creating the Folder where the Boot Image will be mounted"
        md $MountPath -Force
    }

Write-Output "Mounting the previously created Boot Medium ISO"
$MountedISO = Mount-DiskImage -ImagePath $BootMediaPath -PassThru 

Write-Output "Getting the mounted ISO's drive letter"
$Volume = (Get-Volume -FileSystemLabel "Configuration Manager 2012").DriveLetter.ToString()

Write-Output "Getting the Boot Image properties"
#$BootImage = Get-WmiObject -Class SMS_BootImagePackage -Namespace root\sms\site_$SiteCode -Filter "Name = '$($BootImageName)'"

$BootImage = Get-CimInstance -ClassName SMS_BootImagePackage -Namespace root\sms\site_$SiteCode -Filter "Name = '$($BootImageName)'"

Write-Output "Creating a copy of the original WIM. This copy will now be edited."
$CopiedBootImage = Copy-Item $BootImage.PkgSourcePath -Destination $(Join-Path $(Split-Path $BootImage.PkgSourcePath -Parent) $BootImageName'_Adaptiva.wim') -PassThru -Force

Write-Output "Mounting the Boot Image."
Mount-WindowsImage -Path $MountPath -ImagePath $CopiedBootImage.FullName -Index 1

Write-Output "Creating the SMS\DATA Folders"
md -Path $(Join-Path $MountPath SMS\DATA)

Set-Location "$($Volume):"

Write-Output "Copying the needed Files to the new Boot Image."
Copy-Item .\SMS\DATA\* -Destination $MountPath\SMS\DATA -Force | Out-Null
Copy-Item $AdaptivaClientSources\OneSiteDownloader.exe -Destination $MountPath -Force | Out-Null
Copy-Item $AdaptivaClientSources\OneSiteDownloader64.exe -Destination $MountPath -Force | Out-Null

Write-Output "Saving changes to the WIM file and dismounting it."
Dismount-WindowsImage -Path $MountPath -Save 
Dismount-DiskImage -ImagePath $BootMediaPath 

Write-Output "Creating new CM12 Boot Image from this WIM file."
#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }
New-CMBootImage -Path $CopiedBootImage.FullName -Index 1 -Name "$($BootImageName)_AdaptivaOneSite" -Verbose

Set-Location C: