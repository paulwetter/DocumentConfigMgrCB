[wmi]$collsetting = Get-WmiObject -Class sms_collectionsettings -Namespace root\sms\site_pri -Filter "CollectionID='PRI0001E'"
$collsetting = $collsetting.__PATH
$collsetting.PowerConfigs
