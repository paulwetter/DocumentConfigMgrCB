[cmdletBinding()]

param (
    [parameter(Mandatory=$true)]
    [String]$DPName,
    [parameter(Mandatory=$true)]
    [String]$SMSProvider
)

#region Start Functions

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
}

Function Execute-Query {
[cmdletBinding()]

param (
    [parameter(Mandatory=$true)]
    [String]$DPName
)

#MessageID 9504 is "Distribution Point un-installation successfully completed on server"
$Query = "SELECT * FROM SMS_StatMsgWithInsStrings WHERE MessageID='9504' AND InsString2 = '$DPName' AND Component = 'SMS_DISTRIBUTION_MANAGER' AND Win32Error = '0'"
$result = Get-CimInstance -Namespace root\sms\site_$SiteCode -Query $Query
#check time if recent enough

$currenttime = get-date
#Compare-Object 
$timediff = ($currenttime - $result.Time)

if (($timediff.Hour -eq '0') -and ($timediff.Minute -lt '15')) {
    return $result
    }
else {
    return $null
    }

}

#endregion Functions


Get-SiteCode

do {
    Start-Sleep -Seconds 15
    Write-Output 'Checking for DP un-install status'; 
    $result = Execute-Query -DPName $DPName; 
} 
while ([string]::IsNullOrEmpty($result))

Write-Output "$DPName is successfully un-installed"