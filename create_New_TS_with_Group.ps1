$Computer = "."
$Class = "SMS_TaskSequencePackage"
$Method = "GetSequence"
$SourceSequence = $Null
$MC = [WmiClass]"\\$Computer\ROOT\SMS\site_PR1:$Class"


$InParams = $mc.psbase.GetMethodParameters($Method)

$TaskSequence = Get-WmiObject -Class SMS_TaskSequencePackage -Namespace root\sms\site_PR1 -Filter "Name = 'Deploy_Win8'"

$InParams.TaskSequencePackage = $TaskSequence
$SourceSequence = $mc.PSBase.InvokeMethod($Method, $inParams, $Null)

# create a blank Sequence
$Sequence = ([WmiClass]"ROOT\SMS\site_PR1:SMS_TaskSequencePackage").ImportSequence('<sequence version="3.00"/>').TaskSequence
$Steps = @()
foreach ($Step in $SourceSequence.TaskSequence.Steps)
    {
        $Steps += $Step              
    }

$Sequence.Steps += $Steps

# create the new Task Sequence Package object
$EmptySequence = ([WmiClass]"ROOT\SMS\site_PR1:SMS_TaskSequencePackage").CreateInstance()
$EmptySequence.Name = "NewTS"
$EmptySequence.Description = "New Task Sequence"
$EmptySequence.Category = "OSD"

# commit the new Task Sequence
([WmiClass]"ROOT\SMS\site_PR1:SMS_TaskSequencePackage").SetSequence($EmptySequence,$Sequence)