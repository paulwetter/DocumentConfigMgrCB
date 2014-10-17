
$SiteCode = "HQ1"
$MachineName = "DPM01"

$ResourceID = (Get-WmiObject -Class SMS_R_System -Namespace "root\sms\site_$SiteCode" -Filter "Name like `'$MachineName`'").ResourceID

$Resource = [wmi]"\\.\root\sms\site_$($SiteCode):SMS_R_System.resourceID=$ResourceID"

$Resource.psbase.delete()
