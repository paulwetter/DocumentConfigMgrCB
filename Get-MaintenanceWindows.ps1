param (
[string]$SiteCode,
[string]$FilePath
)

$CollSettings = ""
[array]$CollIDs = @()

Function Convert-NormalDateToConfigMgrDate {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$starttime
    )

    return [System.Management.ManagementDateTimeconverter]::ToDateTime($starttime)
}

Function Read-ScheduleToken {

$SMS_ScheduleMethods = "SMS_ScheduleMethods"
$class_SMS_ScheduleMethods = [wmiclass]""
$class_SMS_ScheduleMethods.psbase.Path ="ROOT\SMS\Site_$($SiteCode):$($SMS_ScheduleMethods)"
        
$script:ScheduleString = $class_SMS_ScheduleMethods.ReadFromString($ServiceWindow.ServiceWindowSchedules)
return $ScheduleString
}

############### Main script starts here ######################

#Collecting all collections with Maintenance windows configured
$Collections = Get-WmiObject -Class SMS_Collection -Namespace root\SMS\Site_$($SiteCode) | Where-Object {$_.ServiceWindowsCount -gt 0}

#get the collection IDs of these collections
foreach ($Collection in $Collections)
    {
        $CollIDs += $Collection.CollectionID

    }

#get the maintenance window information
foreach ($CollectionID in $CollIDs)
    {   

        $CollName = (Get-WmiObject -Class SMS_Collection -Namespace root\sms\Site_$($SiteCode) | Where-Object {$_.CollectionID -eq "$($CollectionID)"}).Name
        "Working on Collection $($CollName)" | Out-File -FilePath $FilePath -Append
        "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\" | Out-File -FilePath $FilePath -Append
        $CollSettings = Get-WmiObject -class sms_collectionsettings -Namespace root\sms\site_$($SiteCode) | Where-Object {$_.CollectionID -eq "$($CollectionID)"}
        
        $CollSettings = [wmi]$CollSettings.__PATH
        
        #$CollSettings.Get() | Out-Null
        
        $ServiceWindows = $($CollSettings.ServiceWindows)
        
        $ServiceWindows = [wmi]$ServiceWindows.__PATH
               
        foreach ($ServiceWindow in $ServiceWindows)
            {
                
                $ScheduleString = Read-ScheduleToken
                
                "Working on maintenance window $($ServiceWindow.Name)" | Out-File -FilePath $FilePath -Append
                #$ServiceWindow.Description
                #$starttime = (Convert-NormalDateToConfigMgrDate $ScheduleString.TokenData.starttime)
                
                switch ($ServiceWindow.ServiceWindowType)
                    {
                        0 {"This is a Task Sequence maintenance window" | Out-File -FilePath $FilePath -Append}
                        1 {"This is a general maintenance window" | Out-File -FilePath $FilePath -Append}                        
                    }   
                switch ($ServiceWindow.RecurrenceType)
                    {
                        1 {"This maintenance window occurs only once on $($startTime) and lasts for $($ScheduleString.TokenData.HourDuration) hour(s) and $($ScheduleString.TokenData.MinuteDuration) minute(s)." | Out-File -FilePath $FilePath -Append}
                        2 
                            {
                                if ($ScheduleString.TokenData.DaySpan -eq "1")
                                    {
                                        $daily = "daily"
                                    }
                                else
                                    {
                                        $daily = "every $($ScheduleString.TokenData.DaySpan) days"
                                    }
                        
                                "This maintenance window occurs $($daily)." | Out-File -FilePath $FilePath -Append
                            }
                        3 
                            {
                                switch ($ScheduleString.TokenData.Day)
                                    {
                                        1 {$weekday = "Sunday"}
                                        2 {$weekday = "Monday"}
                                        3 {$weekday = "Tuesday"}
                                        4 {$weekday = "Wednesday"}
                                        5 {$weekday = "Thursday"}
                                        6 {$weekday = "Friday"}
                                        7 {$weekday = "Saturday"}
                                    }
                                
                                "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofWeeks) week(s) on $($weekday) and lasts $($ScheduleString.TokenData.HourDuration) hour(s) and $($ScheduleString.TokenData.MinuteDuration) minute(s) starting on $($startTime)." | Out-File -FilePath $FilePath -Append}
                        4 
                            {
                                switch ($ScheduleString.TokenData.Day)
                                    {
                                        1 {$weekday = "Sunday"}
                                        2 {$weekday = "Monday"}
                                        3 {$weekday = "Tuesday"}
                                        4 {$weekday = "Wednesday"}
                                        5 {$weekday = "Thursday"}
                                        6 {$weekday = "Friday"}
                                        7 {$weekday = "Saturday"}
                                    }
                                switch ($ScheduleString.TokenData.weekorder)
                                    {
                                        0 {$order = "last"}
                                        1 {$order = "first"}
                                        2 {$order = "second"}
                                        3 {$order = "third"}
                                        4 {$order = "fourth"}
                                    }
                                
                                "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofMonths) month(s) on every $($order) $($weekday)" | Out-File -FilePath $FilePath -Append
                            }

                        5 
                            {
                                if ($ScheduleString.TokenData.MonthDay -eq "0")
                                    { 
                                        $DayOfMonth = "the last day of the month"
                                    }
                                else
                                    {
                                        $DayOfMonth = "day $($ScheduleString.TokenData.MonthDay)"
                                    }
                                "This maintenance window occurs every $($ScheduleString.TokenData.ForNumberofMonths) month(s) on $($DayOfMonth)." | Out-File -FilePath $FilePath -Append                                                  
                                "It lasts $($ScheduleString.TokenData.HourDuration) hours and $($ScheduleString.TokenData.MinuteDuration) minutes." | Out-File -FilePath $FilePath -Append
                            }

                    }
                switch ($ServiceWindow.IsEnabled)
                    {
                        true {"The maintenance window is enabled" | Out-File -FilePath $FilePath -Append}
                        false {"The maintenance window is disabled" | Out-File -FilePath $FilePath -Append}
                    }
                "Going to next Maintenance window" | Out-File -FilePath $FilePath -Append
                "---------------------------------------------" | Out-File -FilePath $FilePath -Append
            }
        "No more maintenance windows present on this collection. Going to next collection." | Out-File -FilePath $FilePath -Append
        "###############################################" | Out-File -FilePath $FilePath -Append
    }
"No more maintenance windows present. Exiting documentation script"  | Out-File -FilePath $FilePath -Append