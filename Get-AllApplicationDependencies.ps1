param (
[string]$SMSProvider,
[string]$ApplicationName
)

$Subs = @()
$Subdependencies = $null


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

##################### MAIN SCRIPT STARTS HERE #######################


#Import the CM12 Powershell cmdlets
if (-not (Test-Path -Path $SiteCode))
    {
        Write-Verbose "CM12 module has not been imported yet, will import it now."
        Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length – 5) + '\ConfigurationManager.psd1') | Out-Null
    }
#CM12 cmdlets need to be run from the CM12 drive
Set-Location "$($SiteCode):" | Out-Null
if (-not (Get-PSDrive -Name $SiteCode))
    {
        Write-Error "There was a problem loading the Configuration Manager powershell module and accessing the site's PSDrive."
        exit 1
    }


$App = Get-CMDeploymentType -ApplicationName "$($ApplicationName)"

$Dependencies = Get-WmiObject -Class SMS_CIRelation -Namespace root\sms\site_$SiteCode -Filter "FromCIID = $($App.CI_ID)"

foreach ($Dependency in $Dependencies)
    {
        $DependentDTs = Get-WmiObject -Class SMS_DeploymentType -Namespace root\sms\site_$SiteCode -Filter "CI_ID = $($Dependency.ToCIID)"
        foreach ($DependentDT in $DependentDTs)
            {
                
                Write-Host "Found Dependency: " (Get-WmiObject -Class SMS_DeploymentType -Namespace root\sms\site_$SiteCode -Filter "CI_ID = $($Dependency.ToCIID)").LocalizedDisplayName -ForegroundColor Magenta
                try 
                    {
                        $Subdependencies = Get-WmiObject -Class SMS_CIRelation -Namespace root\sms\site_$SiteCode -Filter "(FromCIID = $($DependentDT.CI_ID)) AND (RelationType='10')"
                    }
                catch
                    {
                        ""
                    }

        if (-not [string]::IsNullOrEmpty($Subdependencies))
            {
                Write-Host "Checking Subdependencies for Dependency $($DependentDT.LocalizedDisplayName)" -ForegroundColor DarkGreen
                foreach ($Subdependency in $Subdependencies)
                    {
                        "Found Subdependency: " + (Get-WmiObject -Class SMS_DeploymentType -Namespace root\sms\site_$SiteCode -Filter "CI_ID = $($Subdependency.ToCIID)").LocalizedDisplayName
                        $Subs = @()
                        
                        try 
                            {
                                
                                $Subs += Get-WmiObject -Class SMS_CIRelation -Namespace root\sms\site_$SiteCode -Filter "(FromCIID = $($Subdependency.ToCIID)) AND (RelationType='10')"
                                
                                   
                                   function Get-RecursevelyAllSubdependencies 
                                    {
                                        param ($Sub)

                                       "Found Grandchild of $((Get-WmiObject -Class SMS_DeploymentType -Namespace root\sms\site_$SiteCode -Filter "CI_ID = $($Subdependency.ToCIID)").LocalizedDisplayName): " + (Get-WmiObject -Class SMS_DeploymentType -Namespace root\sms\site_$SiteCode -Filter "CI_ID = $($Sub.ToCIID)").LocalizedDisplayName
                                       try
                                        {
                                            $Sub = Get-WmiObject -Class SMS_CIRelation -Namespace root\sms\site_$SiteCode -Filter "(FromCIID = $($Sub.ToCIID)) AND (RelationType='10')" -ErrorAction Stop
                                            if ($Sub)
                                                { Get-RecursevelyAllSubdependencies $Sub }
                                        }
                                       catch
                                        {
                                        }

        
                                    }
                                    foreach ($Sub in $Subs)
                                        {
                                            Get-RecursevelyAllSubdependencies $Sub
                                        }

                            }
                        catch
                            {
                                ""
                            }   
                    }
                }
                ""
                }
    }

    Set-Location c:
