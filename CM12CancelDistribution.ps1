param (
$DPServer,
$SMSProvider,
$AppName,
[switch]$Package,
[switch]$Application
)

$Colon = ':'
$Class = 'SMS_DistributionPoint'

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

$NALPath = (Get-WmiObject -Class SMS_DistributionPointInfo -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "NALPath LIKE '%$DPServer%'").NALPath

if ($Package)
    {
        $PackageID = (Get-WmiObject -Class SMS_Package -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "Name = '$AppName'").PackageID
    }
elseif ($Application)
    {
        Write-Output "Looking for Application $AppName"
        $ApplicationObject = Get-WmiObject SMS_Application -Namespace root\SMS\Site_$SiteCode -ComputerName $SMSProvider -Filter "LocalizedDisplayName = '$AppName'"
        $PackageID = (Get-WmiObject -Class SMS_DistributionPoint -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Filter "SecureObjectID = '$($ApplicationObject.ModelName)'").PackageID
    }

$DP = [WmiClass]"\\$SMSProvider\ROOT\SMS\site_$SiteCode$Colon$Class"

$inParams = $DP.psbase.GetMethodParameters('CancelDistribution')

$inParams.NALPath = $NALPath
$inParams.PackageId = $PackageID

Write-Output "Going to cancel Distribution of Content $AppName with PackageID $PackageID to DistributionPoint $DPServer"

$Job = $DP.PSBase.InvokeMethod('CancelDistribution', $inParams, $Null)

if ($Job.ReturnValue -eq 0)
    {
        Write-Output "Distribution Job cancelled."
    }
else
    {
        Write-Output "Cancelling Distribution did not succeed."
    }
