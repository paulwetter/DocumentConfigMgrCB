[cmdletBinding()]

param (
  [parameter(
	Position = 0, 
	Mandatory=$true)
	] 
	[Alias('SMS')]
	[ValidateScript({Test-Connection -ComputerName $_ -Count 1 -Quiet})]
  [ValidateNotNullOrEmpty()]
	[string]$SMSProvider='')

Function Get-SiteCode
{
    $wqlQuery = 'SELECT * FROM SMS_ProviderLocation'
    $a = Get-WmiObject -Query $wqlQuery -Namespace 'root\sms' -ComputerName $SMSProvider
    $a | ForEach-Object {
        if($_.ProviderForLocalSite)
            {
                $script:SiteCode = $_.SiteCode
            }
    }
    return $SiteCode
}


Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
Set-Location "$($SiteCode):"
