param (
$SMSProvider,
$StartingIP,
$EndingIP
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
}Function New-IPRange ($start, $end) {
# created by Tobias Weltner, MVP PowerShell, http://powershell.com/cs/blogs/tobias/archive/2011/02/20/creating-ip-ranges-and-other-type-magic.aspx
$ip1 = ([System.Net.IPAddress]$start).GetAddressBytes()
[Array]::Reverse($ip1)
$ip1 = ([System.Net.IPAddress]($ip1 -join '.')).Address

$ip2 = ([System.Net.IPAddress]$end).GetAddressBytes()
[Array]::Reverse($ip2)
$ip2 = ([System.Net.IPAddress]($ip2 -join '.')).Address

for ($x=$ip1; $x -le $ip2; $x++) {
$ip = ([System.Net.IPAddress]$x).GetAddressBytes()
[Array]::Reverse($ip)
$ip -join '.'
}
}$RangeGiven = $null$RangeExisting = $nullGet-SiteCode | Out-Nullif (-not ($SiteCode))    {        Write-Error "SiteCode could not be determined."        exit 1    }$Boundaries = Get-WmiObject -Namespace root\sms\site_$SiteCode -ComputerName $SMSProvider -Query "SELECT * FROM SMS_Boundary WHERE BoundaryType ='3'"$RangeGiven = New-IPRange $StartingIP $EndingIPforeach ($Boundary in $Boundaries)    {        $BoundaryStart = $Boundary.Value.Split('-')[0]        $BoundaryEnd   = $Boundary.Value.Split('-')[1]        $RangeExisting = New-IPRange $BoundaryStart $BoundaryEnd        $index = 0        while ($index -lt $($RangeExisting.length))            {                $index2 = 0                while ($index2 -lt $($RangeGiven.length))                    {                        if ($($RangeExisting[$index]) -eq $($RangeGiven[$index2]))                            {                                Write-Output "IP $($RangeGiven[$index2]) already exists in Boundary $($Boundary.DisplayName)"                             }                    $index2++                    }                $index++            }    }