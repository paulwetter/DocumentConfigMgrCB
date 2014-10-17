param (
$ServerName,
$SMSProvider
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
}

Get-SiteCode

$ServerObjectWMI = Get-WmiObject -Class SMS_SCI_SysResUse -Namespace root\SMS\Site_$SiteCode -ComputerName $SMSProvider -Filter "NetworkOSPath LIKE '%\\$ServerName%' AND RoleName = 'SMS Site System'"
$ServerObjectWMI.Delete()