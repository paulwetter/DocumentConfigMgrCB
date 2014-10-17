param (

[string]$TaskSequenceName,
[string]$TargetFolderName,
[string]$SiteCode,
[string]$CMProvider
)

[int]$ObjectID = 20
[string]$TaskSequence = ""
$TaskSequence = (Get-WmiObject -Class SMS_TaskSequencePackage -Namespace root\sms\site_$SiteCode -Filter "Name = '$($TaskSequenceName)'" -ComputerName $CMProvider).PackageID
[int]$SourceFolder = (Get-WmiObject -Class SMS_ObjectContainerItem -Namespace root\sms\site_$SiteCode -Filter "InstanceKey = '$($TaskSequence)'" -ComputerName $CMProvider).ContainerNodeID
[int]$TargetFolder = (Get-WmiObject -Class SMS_ObjectContainerNode -Namespace root\sms\site_$SiteCode -Filter "ObjectType = '20' and Name = '$($TargetFolderName)'" -ComputerName $CMProvider).ContainerNodeID

$Parameters = ([wmiclass]"\\$($CMProvider)\root\SMS\Site_$($SiteCode):SMS_ObjectContainerItem").psbase.GetMethodParameters("MoveMembers")


$Parameters.ObjectType = $ObjectID
$Parameters.ContainerNodeID = $SourceFolder
$Parameters.TargetContainerNodeID = $TargetFolder
$Parameters.InstanceKeys = $TaskSequence

try {
        $Output = ([wmiclass]"\\$($CMProvider)\root\SMS\Site_$($SiteCode):SMS_ObjectContainerItem").psbase.InvokeMethod("MoveMembers",$Parameters,$null)
        if ($Output.ReturnValue -eq "0")
            {
                Write-Host "Task Sequence $($TaskSequence) successfully moved to Folder $($TargetFolderName)."
            }
    }
catch [Exception]
    {
        Write-Error -Message "Something went wrong."
    }