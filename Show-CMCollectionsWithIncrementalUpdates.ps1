param (
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
return $SiteCode
}

$SiteCode = Get-SiteCode

Get-WmiObject -Class SMS_Collection -Namespace Root\SMS\Site_$SiteCode -ComputerName $SMSProvider -Filter "RefreshType = '4' AND CollectionType = '2'" | Select-Object Name, CollectionID, MemberCount, LimitToCollectionName, LimitToCollectionID | Out-GridView
Get-WmiObject -Class SMS_Collection -Namespace Root\SMS\Site_$SiteCode -ComputerName $SMSProvider -Filter "RefreshType = '4' AND CollectionType = '1'" | Select-Object Name, CollectionID, MemberCount, LimitToCollectionName, LimitToCollectionID | Out-GridView
