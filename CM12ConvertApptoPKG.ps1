# .\CM12ConvertApptoPKG.ps1 -SiteCode DE0 -SMSProvider localhost -ApplicationName "DE_Microsoft_MBAMClient_2.1"


param (
[string]$SiteCode,
[string]$SMSProvider,
[string]$ApplicationName
)

function Get-ExecuteWqlQuery($siteServerName, $query)
{
  $returnValue = $null
  $connectionManager = New-Object Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine.WqlConnectionManager
  
  if($connectionManager.Connect($siteServerName))
  {
      $result = $connectionManager.QueryProcessor.ExecuteQuery($query)
      
      foreach($i in $result.GetEnumerator())
      {
        $returnValue = $i
        break
      }
      
      $connectionManager.Dispose() 
  }
  
  $returnValue
}

function Get-ApplicationObjectFromServer($appName,$siteServerName)
{    
    $resultObject = Get-ExecuteWqlQuery $siteServerName "select thissitecode from sms_identification" 
    $siteCode = $resultObject["thissitecode"].StringValue
    
    $path = [string]::Format("\\{0}\ROOT\sms\site_{1}", $siteServerName, $siteCode)
    $scope = New-Object System.Management.ManagementScope -ArgumentList $path
    
    $query = [string]::Format("select * from sms_application where LocalizedDisplayName='{0}' AND ISLatest='true'", $appName.Trim())
    
    $oQuery = New-Object System.Management.ObjectQuery -ArgumentList $query
    $obectSearcher = New-Object System.Management.ManagementObjectSearcher -ArgumentList $scope,$oQuery
    $applicationFoundInCollection = $obectSearcher.Get()    
    $applicationFoundInCollectionEnumerator = $applicationFoundInCollection.GetEnumerator()
    
    if($applicationFoundInCollectionEnumerator.MoveNext())
    {
        $returnValue = $applicationFoundInCollectionEnumerator.Current
        $getResult = $returnValue.Get()        
        $sdmPackageXml = $returnValue.Properties["SDMPackageXML"].Value.ToString()
        [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($sdmPackageXml)
    }
}


 function Load-ConfigMgrAssemblies()
 {
     
     $AdminConsoleDirectory = Split-Path $env:SMS_ADMIN_UI_PATH -Parent
     $filesToLoad = "Microsoft.ConfigurationManagement.ApplicationManagement.dll","AdminUI.WqlQueryEngine.dll", "AdminUI.DcmObjectWrapper.dll" 
     
     Set-Location $AdminConsoleDirectory
     [System.IO.Directory]::SetCurrentDirectory($AdminConsoleDirectory)
     
      foreach($fileName in $filesToLoad)
      {
         $fullAssemblyName = [System.IO.Path]::Combine($AdminConsoleDirectory, $fileName)
         if([System.IO.File]::Exists($fullAssemblyName ))
         {   
             $FileLoaded = [Reflection.Assembly]::LoadFrom($fullAssemblyName )
         }
         else
         {
              Write-Output ([System.String]::Format("File not found {0}",$fileName )) -backgroundcolor "red"
         }
      }
 }

# The move-Package function is optional

Function Move-Package
{
    $ParentFolder = Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\Site_$($SiteCode) -Filter "ObjectType = 2 AND Name = 'DE'"
    $TargetFolder = Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\SMS\Site_$($SiteCode) -Filter "ObjectType = 2 AND Name = 'Task Sequence' AND ParentContainerNodeID = '$($ParentFolder.ContainerNodeID)'"

    $Params = ([wmiclass]"\\$($SMSProvider)\root\sms\site_$($SiteCode):SMS_ObjectContainerItem").psbase.GetMethodParameters("MoveMembers")

    $Params.ObjectType = 2
    $Params.ContainerNodeID = 0
    $Params.TargetContainerNodeID = $TargetFolder.ContainerNodeID
    $Params.InstanceKeys = $Package.PackageID

    ([wmiclass]"\\$($SMSProvider)\root\sms\site_$($SiteCode):SMS_ObjectContainerItem").psbase.InvokeMethod("MoveMembers",$Params,$null)
}

$LocBefore = $null
$LocBefore = Get-Location
Load-ConfigMgrAssemblies 
Set-Location "$($SiteCode):"


$application = [wmi](Get-WmiObject SMS_Application -Namespace root\sms\site_$($SiteCode) |  where {($_.LocalizedDisplayName -eq "$($ApplicationName)") -and ($_.IsLatest)}).__PATH

$global:applicationXML = Get-ApplicationObjectFromServer "$($ApplicationName)" $SMSProvider

if ($applicationXML.DeploymentTypes -ne $null)
    { 
        foreach ($a in $applicationXML.DeploymentTypes)
            {
                $a 
                $InstallCommandLine = $a.Installer.InstallCommandLine
                $ContentPath = $a.Installer.Contents[0].Location
            }
    }
<#

if (Get-CMPackage -Name "$($ApplicationName)")
    {
        Write-Output "Package already exists"
        Set-Location $LocBefore
        exit
    }

New-CMPackage -Name "$($ApplicationName)" -Path $ContentPath 
New-CMProgram -PackageName "$($ApplicationName)" -StandardProgramName "Install $($ApplicationName)" -RunMode RunWithAdministrativeRights -UserInteraction $false -RunType Hidden -ProgramRunType WhetherOrNotUserIsLoggedOn -CommandLine $InstallCommandLine 

$Package = Get-CMPackage -Name "$($ApplicationName)"

$Program = Get-WmiObject -Class sms_program -Namespace root\sms\site_$SiteCode -Filter "PackageID = '$($Package.PackageID)'"

 if (-not ($Program.ProgramFlags -band 0x00000001))
    {
        $Program.ProgramFlags = $Program.ProgramFlags -bxor 0x00000001
        $Program.put()
    }

#Move-Package

#>

Set-Location $LocBefore