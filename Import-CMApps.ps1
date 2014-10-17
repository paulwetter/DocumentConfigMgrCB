Param(	
    
    [parameter(
	Position = 0, 
	Mandatory=$true )
	] 
	[Alias("SMS")]
    [ValidateScript({
        $ping = New-Object System.Net.NetworkInformation.Ping
        $ping.Send("$_", 5000)})]
	[ValidateNotNullOrEmpty()]
	[string]$SMSProvider="",
    
    [parameter(
	Position = 1, 
	Mandatory=$true )
	] 
    [string]$ImportFolder
)


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

$SiteCode = Get-SiteCode

#Import the CM12 Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
    {
        Write-Verbose "CM12 module has not been imported yet, will import it now."
        Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null
    }

$Exports = Get-ChildItem -Path $ImportFolder -Filter *.zip

#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }

foreach ($Export in $Exports)
    {
        try 
            {
                Import-CMApplication -FilePath $($Export.FullName) -ImportActionType DirectImport
            }
        catch 
            {      
            }
    }