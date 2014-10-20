#region declare variables
[CmdletBinding()]
param()

$SQLNodes = @('SQL01AO', 'SQL02AO')
$AlwaysOnInstance = 'ALWAYSON'
$AvailabilityGroupName = 'AG01'

$location = Split-Path $MyInvocation.MyCommand.Path -Parent

#endregion declare variables

#region helper functions
Function Get-SQLAlwaysOnReplicaNode {
  # Full credit to Brian P ODwyer http://www.mssqltips.com/sqlservertip/3206/finding-primary-replicas-for-sql-server-2012-alwayson-availability-groups-with-powershell/
  # small changes to $connectionstring and output by David O'Brien
  [CmdletBinding()]
  param (
    [array]$SQLNodes,   
    [string]$SQLInstance,
    [ValidateSet('PRIMARY','SECONDARY')] 
    [string]$Role
  )
  
  ## Setup dataset to hold results
  $dataset = New-Object System.Data.DataSet
  ## populate variable with collection of SQL instances
  #$serverlist='SQL01AO\AlwaysOn', 'SQL02AO\AlwaysOn'
  ## Setup connection to SQL server inside loop and run T-SQL against instance 
  foreach($Server in $SQLNodes) {
    if ([string]::IsNullOrEmpty($SQLInstance)) {
      $connectionString = "Provider=sqloledb;Data Source=$Server;Initial Catalog=Master;Integrated Security=SSPI;"
    }
    else {
      $connectionString = "Provider=sqloledb;Data Source=$Server\$SQLInstance;Initial Catalog=Master;Integrated Security=SSPI;"
    }
    ## place the T-SQL in variable to be executed by OLEDB method 
    
    $sqlcommand="
      IF SERVERPROPERTY ('IsHadrEnabled') = 1
      BEGIN
      SELECT
      AGC.name
      , RCS.replica_server_name
      , ARS.role_desc
      , AGL.dns_name
      FROM
      sys.availability_groups_cluster AS AGC
      INNER JOIN sys.dm_hadr_availability_replica_cluster_states AS RCS
      ON
      RCS.group_id = AGC.group_id
      INNER JOIN sys.dm_hadr_availability_replica_states AS ARS
      ON
      ARS.replica_id = RCS.replica_id
      INNER JOIN sys.availability_group_listeners AS AGL
      ON
      AGL.group_id = ARS.group_id
      WHERE
      ARS.role_desc = '$Role'
      END
    "
    ## Connect to the data source and open it
    $connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
    $command = New-Object System.Data.OleDb.OleDbCommand $sqlCommand,$connection
    $connection.Open()
    ## Execute T-SQL command in variable, fetch the results, and close the connection
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
    #$dataset = New-Object System.Data.DataSet
    [void] $adapter.Fill($dataSet)
    $connection.Close()
  }
  ## Return all of the rows from dataset object
  #$dataSet.Tables | FT -AutoSize
  return ($DataSet.Tables[0] | select replica_server_name).replica_server_name
  
  #>
}

Function Check-AGHealth {
  [CmdletBinding()]
  param (
    [parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$AGName
  )
  
  #Set-Location "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AGName"
  
  $objHealth = Test-SqlAvailabilityGroup -Path "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AGName"
  
  if ($objHealth.HealthState -eq 'Healthy') {
    return $true;
  }
  elseif ($objHealth.HealthState -eq 'Unknown') {
    Write-Verbose 'The AG is in Health State unknown. The script might be checking against a Secondary Replica.';
    return $false;
  }
  else {
    Write-Verbose 'The AG is in an unhealthy state. Please check first before updating.';
    return $false;
  }
}

Function Check-ReplicaHealth {
  [CmdletBinding()]
  param()    
  [int]$HealthIssues = 0
  [array]$UnhealthyReplica = @()
  
  Set-Location "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AvailabilityGroupName\AvailabilityReplicas"
  $objReplicaHealth =  dir | Test-SqlAvailabilityReplica
  foreach ($obj in $objReplicaHealth) {
    if ($obj.HealthState -ne 'Healthy') {
      Write-Verbose "$($obj.Name) is in HealthState $($obj.HealthState)."
      $HealthIssues++
      $UnhealthyReplica += $obj.Name
    }
    else {
      Write-Verbose "$($obj.Name) is in HealthState $($obj.HealthState)."
    }
  }
  if ($HealthIssues -ge 1) {
    Write-Verbose "There were $HealthIssues issues in this AvailabilityGroup."
    return $UnhealthyReplica
  }
  else {
    Write-Verbose "There were $HealthIssues issues in this AvailabilityGroup."
    return $UnhealthyReplica
    #return $true
  }
}

Function Check-DBSynchronisation {
  [CmdletBinding()]
  param()
  [array]$NotSynchronisedDBs = @()
  Set-Location "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AvailabilityGroupName\DatabaseReplicaStates"
  $DBHealthStates = dir | Test-SqlDatabaseReplicaState
  foreach ($DBHealthState in $DBHealthStates) {
    if ($DBHealthState.HealthState -ne 'Healthy') {
      $NotSynchronisedDBs += $DBHealthState.Name,$DBHealthState.AvailabilityReplica,$DBHealthState.HealthState
    }
  }
  Set-Location C:\
  return $NotSynchronisedDBs
}

Function Backup-AvailabilityDB {
  [CmdletBinding()]
  param()
  [int]$global:errorcount = 0
  Set-Location "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AvailabilityGroupName\AvailabilityDatabases"
  $DBs = Get-ChildItem -Path "SQLServer:\SQL\$PrimaryReplicaServer\availabilitygroups\$AvailabilityGroupName\AvailabilityDatabases"
  foreach ($DB in $DBs) {
    Write-Verbose "Starting Backup of $($DB.Name)."
    Backup-SqlDatabase "$($DB.Name)" "\\dc01\sources\SQLAlwaysOn\$($DB.Name)_PoSh.bak"
    Backup-SqlDatabase "$($DB.Name)" "\\dc01\sources\SQLAlwaysOn\$($DB.Name)_PoSh.trn" -BackupAction Log
    Start-Sleep -Seconds 5
    
    if ((!(Test-Path "filesystem::\\dc01\sources\SQLAlwaysOn\$($DB.Name)_PoSh.bak")) -or (!(Test-Path "filesystem::\\dc01\sources\SQLAlwaysOn\$($DB.Name)_PoSh.trn"))) {
      Write-Verbose "$($DB.Name) is missing"
      $errorcount++
    }
  }
  #$errorcount
  if ($errorcount -eq 0) {
    Write-Verbose 'this is true'
    return $true;
  }
  else {
    Write-Verbose 'this is false'
    return $false;
  }
  Set-Location $location
}

Function Disable-AutomaticFailover {
  [cmdletbinding()]
  param(
    $Node,
    $AlwaysOnInstance
  )
  
  $ReplicaServer = $Node+'\'+$AlwaysOnInstance
  
  $NewReplicaString = $ReplicaServer.Replace('\','%5C')
  $NewReplicaString
  try {
    "Set-SqlAvailabilityReplica -AvailabilityMode `"SynchronousCommit`" -FailoverMode 'Manual' -Path SQLSERVER:\Sql\$PrimaryReplicaServer\AvailabilityGroups\$AvailabilityGroupName\AvailabilityReplicas\$NewReplicaString"
    Set-SqlAvailabilityReplica -AvailabilityMode "SynchronousCommit" -FailoverMode 'Manual' -Path SQLSERVER:\Sql\$PrimaryReplicaServer\AvailabilityGroups\$AvailabilityGroupName\AvailabilityReplicas\$NewReplicaString
  }
  catch {
    Write-Error $_
  }
}

Function Get-SiteCode {
  [cmdletBinding()]
  param (
    $SMSProvider
  )
  $wqlQuery = 'SELECT * FROM SMS_ProviderLocation'
  $a = Get-WmiObject -Query $wqlQuery -Namespace 'root\sms' -ComputerName $SMSProvider
  $a | ForEach-Object {
    if($_.ProviderForLocalSite)
    {
      $SiteCode = $_.SiteCode
    }
  }
  return $SiteCode
}

Function Add-NodeToConfigMgrCollection {
  [cmdletBinding()]
  
  param (
    $Node,
    $CollectionID,
    $SiteCode,
    $SMSProvider
  )
  
  $Device = Get-WmiObject -ComputerName $SMSProvider -Class SMS_R_SYSTEM -Namespace root\sms\site_$SiteCode -Filter "Name = '$Node'"
  $objColRuledirect = [WmiClass]"\\$SMSProvider\ROOT\SMS\site_$($SiteCode):SMS_CollectionRuleDirect"
  $objColRuleDirect.psbase.properties["ResourceClassName"].value = "SMS_R_System"
  $objColRuleDirect.psbase.properties["ResourceID"].value = $Device.ResourceID
  
  $MC = Get-WmiObject -Class SMS_Collection -ComputerName $SMSProvider -Namespace "ROOT\SMS\site_$SiteCode" -Filter "CollectionID = '$CollectionID'"
  $InParams = $mc.psbase.GetMethodParameters("AddMembershipRule")
  $InParams.collectionRule = $objColRuledirect
  $inparams.PSBase.properties | select name,Value | Format-Table
  $R = $mc.PSBase.InvokeMethod("AddMembershipRule", $inParams, $Null)
}
#endregion helper functions


#region main script starts here
if (Get-Module -Name SQLPS) {
  Write-Verbose 'SQLPS Module present, continue';
}
else {
  Write-Verbose 'SQLPS Module is not present, better import it';
  
  try {
    Import-Module SQLPS -DisableNameChecking -ErrorAction SilentlyContinue;
  }
  catch {
    Write-Error $_
  }
};

if (! (Test-Path 'SQLServer:')) {
  Write-Verbose 'Cannot access the SQLServer PSDrive. Exiting.';
  #exit after here
}
else {
  Set-Location $location
}
    #>
$PrimaryReplicaServer = Get-SQLAlwaysOnReplicaNode -SQLNodes $SQLNodes -SQLInstance $AlwaysOnInstance -Role PRIMARY;
$SecondaryReplicaServer = Get-SQLAlwaysOnReplicaNode -SQLNodes $SQLNodes -SQLInstance $AlwaysOnInstance -Role SECONDARY;

#region executing HealthChecks

if (!(Check-AGHealth -AGName $AvailabilityGroupName)) {
  Set-Location C:\
  Set-Location $location
  throw "AvailabilityGroup $AvailabilityGroupName is in an unhealthy or unknown state."
}
    
$UnhealthyReplica = Check-ReplicaHealth

if (-not [string]::IsNullOrEmpty($UnhealthyReplica)) {
  Set-Location C:\
  Set-Location $location
  throw "Issues found with $UnhealthyReplica. Please check the node for any issues."
}

$NotSynchronisedDBs = Check-DBSynchronisation

if (-not [string]::IsNullOrEmpty($NotSynchronisedDBs)) {
  Set-Location C:\
  Set-Location $location
  throw "Issues found with $NotSynchronisedDBs. Please check the node for any issues."
}

Write-Verbose 'All seems to be fine. Let''s go'

#endregion HealthChecks

#region back up the AlwaysOn DBs
<#
try {
    Backup-AvailabilityDB
}
catch {
    Write-Error $_
}
<#
if (Backup-AvailabilityDB) {
    Write-Verbose 'All Databases backed up' 
    }
else {
    throw "Problem backing up Databases. $errorcount errors."
    }
    #>
#Backup-AvailabilityDB

#endregion backup

foreach ($SQLNode in $SQLNodes)
{
  Disable-AutomaticFailover -Node $SQLNode -AlwaysOnInstance $AlwaysOnInstance -Verbose
}

$OriginalPrimary = $PrimaryReplicaServer

<#

Start Updating one Secondary Node at a time

#>

$SiteCode = Get-SiteCode -SMSProvider CM01.Do.Local
$i = 0
foreach ($SecondaryReplica in $SecondaryReplicaServer) {
  if (-not ($AlreadyPatched -contains $SecondaryReplica.Split('\')[0])) {
    try {
      $i++
      Write-Verbose "Patching Server round $i = $($SecondaryReplica.Split('\')[0])"
      
      #Add current secondary node to ConfigMgr collection to receive its updates
      Add-NodeToConfigMgrCollection -Node $SecondaryReplica.Split('\')[0] -SiteCode $SiteCode -SMSProvider cm01.do.local -CollectionID 'HQ100066' -Verbose
      Start-Sleep -Seconds 30
      
      Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000113}' -ComputerName $SecondaryReplica.Split('\')[0]
      Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000042}' -ComputerName $SecondaryReplica.Split('\')[0]
      
      #Check if all updates have been installed and server finished rebooting
      
      $AlreadyPatched += $SecondaryReplica.Split('\')[0]
    }
    catch {
      Write-Error $_
    }
  }
  else {
    Write-Verbose "$($SecondaryReplica.Split('\')[0]) has already been patched. Skipping."
  }
}
<#
# fail over to one of the secondary nodes and update the primary node, after that, fail over again to the original primary node
Switch-SqlAvailabilityGroup -Path SQLSERVER:\Sql\$(Get-Random -InputObject $SecondaryReplicaServer)\AvailabilityGroups\$AvailabilityGroupName -Verbose
Add-NodeToConfigMgrCollection -Node $PrimaryReplicaServer.Split('\')[0] -SiteCode $SiteCode -SMSProvider cm01.do.local -CollectionID 'HQ100066' -Verbose
Start-Sleep -Seconds 30

Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000113}' -ComputerName $PrimaryReplicaServer.Split('\')[0]
Invoke-WmiMethod -Namespace root\ccm -Class sms_client -Name TriggerSchedule '{00000000-0000-0000-0000-000000000042}' -ComputerName $PrimaryReplicaServer.Split('\')[0]

#Check if all updates have been installed and server finished rebooting


#If the primary node is finished updating, fail over again to the Primary
Switch-SqlAvailabilityGroup -Path SQLSERVER:\Sql\$PrimaryReplicaServer\AvailabilityGroups\$AvailabilityGroupName -Verbose
#>
#$SecondaryReplicaServer
Set-Location C:\
Set-Location $location
#endregion that's it!
