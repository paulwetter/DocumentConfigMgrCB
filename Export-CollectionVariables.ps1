<#
.\Export-CollectionVariables.ps1 -SMSProvider localhost -CollectionName Parent -OutputFile $env:TEMP\collvars.csv
#>


param
(
$SMSProvider,
$CollectionName,
$OutputFile
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

Function Get-CMCollectionID
{
    $script:CollectionID = Get-WmiObject -Class SMS_Collection -Namespace root\SMS\Site_$SiteCode -ComputerName $SMSProvider -Filter "Name = '$($CollectionName)'"
    return $CollectionID.CollectionID
}

$SiteCode = Get-SiteCode
$CollectionID = Get-CMCollectionID


$CollSettings = Get-WmiObject -Class SMS_CollectionSettings -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "CollectionID = '$($CollectionID)'"

$CollSettings = [wmi]$($CollSettings).__PATH

$Variables = $CollSettings.CollectionVariables

foreach ($Var in $Variables)

    {
        $Var | Select Name,Value | Export-Csv -Append -Path $OutputFile
    }