Function Get-HumanReadableSchedule {
    [CmdletBinding()]
    param(
        $Schedule
    )
    if ($Schedule.HourDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.HourDuration) hours"
    } 
    elseif ($Schedule.MinuteDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.MinuteDuration) minutes"
    } 
    elseif ($Schedule.DayDuration -gt 0) {
        $HrDuration = ", with a duration of $($Schedule.DayDuration) days"
    }

    if ($Schedule.DaySpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.DaySpan) days effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.HourSpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.HourSpan) hours effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.MinuteSpan -gt 0) {
        $HrSched = "Occurs every $($Schedule.MinuteSpan) minutes effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.ForNumberOfWeeks) {
        $HrSched = "Occurs every $($Schedule.ForNumberOfWeeks) weeks on $(Convert-WeekDay $Schedule.Day) effective $($Schedule.StartTime)$HrDuration"
    }
    elseif ($Schedule.ForNumberOfMonths) {
        if ($Schedule.MonthDay -gt 0) {
            $HrSched = "Occurs on day $($Schedule.MonthDay) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
        elseif ($Schedule.MonthDay -eq 0) {
            $HrSched = "Occurs the last day of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
        elseif ($Schedule.WeekOrder -gt 0) {
            switch ($Schedule.WeekOrder) {
                0 { $order = 'last' }
                1 { $order = 'first' }
                2 { $order = 'second' }
                3 { $order = 'third' }
                4 { $order = 'fourth' }
            }
            $HrSched = "Occurs the $($order) $(Convert-WeekDay $Schedule.Day) of every $($Schedule.ForNumberOfMonths) months effective $($Schedule.StartTime)$HrDuration"
        }
    }
    elseif ($Schedule.HourDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.HourDuration) hours"
    } 
    elseif ($Schedule.MinuteDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.MinuteDuration) minutes"
    } 
    elseif ($Schedule.DayDuration -gt 0) {
        $HrSched = "Occurs on $($Schedule.StartTime), with a duration of $($Schedule.DayDuration) days"
    }
    else {
        $HrSched = "Could not calculate schedule"
    }
    return $HrSched
}

Function Get-PWCMDiscoveryMethod {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteServer = $env:COMPUTERNAME,
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteName = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM\CcmEval -Name LastSiteCode -ErrorAction SilentlyContinue).LastSiteCode,
        [Parameter(Mandatory = $false)]
        [ValidateSet('ActiveDirectoryForestDiscovery', 'ActiveDirectoryGroupDiscovery', 'ActiveDirectorySystemDiscovery', 'ActiveDirectoryUserDiscovery', 'NetworkDiscovery', 'HeartbeatDiscovery')]
        [ValidateNotNullOrEmpty()]
        $DiscoveryMethod
    )
    
    switch ($DiscoveryMethod) {
        'ActiveDirectoryForestDiscovery' { $DMX = 1 }
        'ActiveDirectoryGroupDiscovery' { $DMX = 2 }
        'ActiveDirectorySystemDiscovery' { $DMX = 3 }
        'ActiveDirectoryUserDiscovery' { $DMX = 4 }
        'NetworkDiscovery' { $DMX = 5 }
        'HeartbeatDiscovery' { $DMX = 6 }
        Default { $DMX = 7 }
    }

    #region AD Group Discovery
    If ($dmx -eq 2 -or $DMX -eq 7) {
        $ADGDHash = @{ }
        $ADGroupDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_SECURITY_GROUP_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName"
        $ADGDHash.Add('DiscoveryMethod', 'Active Directory Group Discovery')
        foreach ($Prop in $ADGroupDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('IncrementalSyncState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('IncrementalSyncState', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    $IncSchedule = Get-HumanReadableSchedule -Schedule $schedule
                    $ADGDHash.Add('IncrementalSyncSchedule', "$IncSchedule")
                }
                'Full Sync Schedule' {
                    $GDFullSched = Convert-CMSchedule $prop.Value1
                    $FullSchedule = Get-HumanReadableSchedule -Schedule $GDFullSched
                    $ADGDHash.Add('FullSyncSchedule', "$FullSchedule")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADGDHash.Add('DiscoveryState', 'Enabled')
                    }
                    else {
                        $ADGDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Discover DG Membership' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('DistributionGroupDiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('DistributionGroupDiscoveryState', 'Disabled')
                    }
                }
                'Enable Filtering Expired Logon' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('FilterExpiredLogonState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('FilterExpiredLogonState', 'Disabled')
                    }
                }
                'Days Since Last Logon' {
                    $ADGDHash.Add('FilterExpiredLogonTime', $Prop.Value)
                }
                'Enable Filtering Expired Password' {
                    If ($Prop.Value -eq 1) {
                        $ADGDHash.Add('FilterExpiredPasswordState', 'Enabled')
                    }
                    Else {
                        $ADGDHash.Add('FilterExpiredPasswordState', 'Disabled')
                    }
                }
                'Days Since Last Password Set' {
                    $ADGDHash.Add('FilterExpiredPasswordTime', $Prop.Value)
                }
            }    
        }
        $adgroupd = @()
        $ADGroupSearch = @()
        $ADGroupSearchCred = @()
        $start = 0
        foreach ($List in $ADGroupDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { $two = $value }
                            3 { $three = $value }
                            4 {
                                $adgroupd += [pscustomobject]@{'Domain' = $one; 'val1' = $two; 'val2' = $three; 'val3' = $value }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'Search Bases:*' {
                    $SearchDomain = $List.PropertyListName -replace 'Search Bases:', ''
                    $ADGroupSearch += [pscustomobject]@{'Domain' = $SearchDomain; 'SearchBase' = "$($List.Values -join ';')" }
                }
                'AD Accounts:*' {
                    $AccountDomain = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADGroupSearchCred += [pscustomobject]@{'Domain' = $AccountDomain; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        $ADGroupSearchLocations = @()
        foreach ($ADGD in $adgroupd) {
            If ($ADGD.val1 -eq 0) {
                $ADGDType = 'Location'
            }
            else {
                $ADGDType = 'Groups'
            }
            if ($ADGD.val2 -eq 0) {
                $ADGDRecursive = 'Yes'
            }
            else {
                $ADGDRecursive = 'No'
                IF ($ADGDType -like 'Groups') { $ADGDRecursive = 'Not Applicable' }
            }
            $ADGDAccount = 'Site Server'
            foreach ($account in $ADGroupSearchCred) {
                If ($account.Domain -like $ADGD.Domain) {
                    $ADGDAccount = "$($account.Account)"
                }
            }
            foreach ($SB in $ADGroupSearch) {
                If ($SB.Domain -like $ADGD.Domain) {
                    $ADGDSearchBase = "$($SB.SearchBase)"
                }
            }
            $ADGroupSearchLocations += [pscustomobject]@{'Name' = "$($ADGD.Domain)"; 'Type' = "$ADGDType"; 'Recursive' = "$ADGDRecursive"; 'Account' = "$ADGDAccount"; 'SearchBase' = $ADGDSearchBase }
        }
        $ADGDHash.Add('SearchLocations', $ADGroupSearchLocations)
        [PSCustomObject]$ADGDHash
    }
    #endregion AD Group Discovery

    #region AD Forest Discovery
    If ($dmx -eq 1 -or $DMX -eq 7) {
        $ADFDHash = @{ }
        $ADFDHash.Add('DiscoveryMethod', 'Active Directory Forest Discovery')
        $ADForestDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_FOREST_DISCOVERY_MANAGER|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName"
        foreach ($Prop in $ADForestDiscovery.Props) {
            switch ($Prop.PropertyName) {
                'Startup Schedule' {
                    $FDSchedule = Convert-CMSchedule $prop.Value1
                    $ADFDHash.Add('SyncSchedule', "$(Get-HumanReadableSchedule -Schedule $FDSchedule)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADFDHash.Add('DiscoveryState', 'Enabled')
                    }
                    else {
                        $ADFDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Enable AD Site Boundary Creation' {
                    If ($Prop.Value -eq 1) {
                        $ADFDHash.Add('ADSiteBoundaryCreation', 'Enabled')
                    }
                    Else {
                        $ADFDHash.Add('ADSiteBoundaryCreation', 'Disabled')
                    }
                }
                'Enable Subnet Boundary Creation' {
                    If ($Prop.Value -eq 1) {
                        $ADFDHash.Add('SubnetBoundaryCreation', 'Enabled')
                    }
                    Else {
                        $ADFDHash.Add('SubnetBoundaryCreation', 'Disabled')
                    }
                }
            }    
        }
        [PSCustomObject]$ADFDHash
    }
    #endregion AD Forest Discovery

    #region Heartbeat Discovery
    If ($dmx -eq 6 -or $DMX -eq 7) { 
        $HBDHash = @{ }
        $HBDHash.Add('DiscoveryMethod', 'Heartbeat Discovery')
        $HeartbeatDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_ClientConfig WHERE FileType=2 AND ItemName='Client Properties' AND ItemType='Client Configuration'" -Namespace "ROOT\SMS\site_$SiteName"
        Foreach ($prop in $HeartbeatDiscovery.Props) {
            switch ($prop.PropertyName) {
                'DDR Refresh Interval' {
                    $schedule = Convert-CMSchedule $prop.Value2
                    If ($schedule.DaySpan -ne 0){
                        If(($schedule.DaySpan % 7) -eq 0){
                            $weeks = $schedule.DaySpan/7
                            $DDRInterval = "Every $weeks week(s)"
                        } else {
                            $DDRInterval = "Every $($schedule.DaySpan) day(s)"
                        }
                    } elseif ($schedule.HourSpan -ne 0) {
                        $DDRInterval = "Every $($schedule.HourSpan) hour(s)"
                    }
                    $HBDHash.Add('DDRRefreshInterval', "$DDRInterval")
                }
                'Enable Heartbeat DDR' { 
                    If ($Prop.Value -eq 1) {
                        $HBDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $HBDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
            }
        }
        [PSCustomObject]$HBDHash
    }
    #endregion Heartbeat Discovery

    #region Network Discovery
    If ($dmx -eq 5 -or $DMX -eq 7) {
        $NDHash = @{ }
        $NDHash.Add('DiscoveryMethod', 'Network Discovery')
        $NetworkDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_NETWORK_DISCOVERY|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName"
        Foreach ($prop in $NetworkDiscovery.Props) {
            switch ($prop.PropertyName) {
                'Discovery Enabled' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Subnet Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalSubnets', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalSubnets', 'Disabled')
                    }
                }
                'Domain Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalDomain', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalDomain', 'Disabled')
                    }
                }
                'Router Hop Count' {
                    $NDHash.Add('SNMPMaxHopCount', $prop.Value1)
                }
                'DHCP Include Local' {
                    If ($Prop.Value1 -eq "TRUE") {
                        $NDHash.Add('IncludeLocalDHCP', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('IncludeLocalDHCP', 'Disabled')
                    }
                }
                'ICMP Ping Timeout' {
                    If ($Prop.Value1 -eq 2000) {
                        $NDHash.Add('SlowNetwork', 'Enabled')
                    }
                    Else {
                        $NDHash.Add('SlowNetwork', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    If ($prop.Value1) {
                        $SchedVal = $prop.Value1 -split ''
                        $schedules = foreach ($bb in $SchedVal) {
                            $cc = "$cc$bb"
                            If ($cc.length -eq 16) {
                                $cc
                                $cc = ''
                            }
                        }
                        Remove-Variable cc, bb
                        $NDschedules = @()
                        foreach ($schedule in $schedules) {
                            $Sched = Convert-CMSchedule $schedule
                            $NDschedules += "$(Get-HumanReadableSchedule -Schedule $Sched)"
                        }
                        $NDHash.Add('DiscoverySchedule', $NDschedules)
                    }
                }
            }
        }
        $NDSubnets = @()
        $NDDomains = @()
        $NDSnmp = @()
        $NDSnmpDevices = @()
        $NDDHCPServers = @()
        foreach ($List in $NetworkDiscovery.PropLists) {
            Switch ($List.PropertyListName) {
                'Subnet Include' {
                    Foreach ($Value in $List.Values) {
                        $SN = $Value.Split(' ')
                        $NDSubnets += [pscustomobject]@{'Network' = "$($SN[0])"; 'Subnet Mask' = "$($SN[1])"; 'Search' = 'Enabled' }
                    }
                }
                'Subnet Exclude' {
                    Foreach ($Value in $List.Values) {
                        $SN = $Value.Split(' ')
                        $NDSubnets += [pscustomobject]@{'Network' = "$($SN[0])"; 'Subnet Mask' = "$($SN[1])"; 'Search' = 'Disabled' }
                    }
                }
                'Domain Include' {
                    Foreach ($Value in $List.Values) {
                        $NDDomains += [pscustomobject]@{'Domain' = "$Value"; 'Search' = 'Enabled' }
                    }
                }
                'Domain Exclude' {
                    Foreach ($Value in $List.Values) {
                        $NDDomains += [pscustomobject]@{'Domain' = "$Value"; 'Search' = 'Disabled' }
                    }
                }
                'Community Names' {
                    foreach ($Value in $List.Values) {
                        $NDSnmp += "$Value"
                    }
                }
                'Address Include' {
                    foreach ($Value in $List.Values) {
                        $NDSnmpDevices += "$Value"
                    }
                }
                'DHCP Include' {
                    foreach ($Value in $List.Values) {
                        $NDDHCPServers += "$Value"
                    }
                }
                'Network Discovery Protocols' {
                    #{DHCP, OSPF} = always set to this.
                }
                'Address Discovery Protocols' {
                    Switch ($List.values -join ',') {
                        'OSPF,RIP' { $NDADP = 1 }
                        'OSPF,RIP,NETBIOS,DHCP' { $NDADP = 2 }
                    }
                }
                'Address Validation Protocols' {
                    Switch ($List.values -join ',') {
                        'ICMP,NAME_RESOLVE' { $NDAVP = 1 }
                        'ICMP,NAME_RESOLVE,NETBIOS' { $NDAVP = 2 }
                    }
                }
            }
        }
        If ($NDADP -eq 1) {
            $NDType = 'Topology'
        }
        else {
            if ($NDADP -eq 2) {
                If ($NDAVP -eq 1) {
                    $NDType = 'Topology and client'
                }
                else {
                    if ($NDAVP -eq 2) {
                        $NDType = 'Topology, client, and client operating system'
                    }
                }
            }
        }
        $NDHash.Add('DiscoveryType', $NDType)
        $NDHash.Add('SearchSubnets', $NDSubnets)
        $NDHash.Add('SearchDomains', $NDDomains)
        $NDHash.Add('SearchSNMPCommunities', $NDSnmp)
        $NDHash.Add('SearchSNMPDevices', $NDSnmpDevices)
        $NDHash.Add('SearchDHCPServers', $NDDHCPServers)
        [PSCustomObject]$NDHash
    }
    #endregion Network Discovery

    #region AD System Discovery
    If ($dmx -eq 3 -or $DMX -eq 7) {
        $ADSDHash = @{ }
        $ADSDHash.Add('DiscoveryMethod', 'Active Directory System Discovery')
        $ADSystemDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_SYSTEM_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName"
        foreach ($Prop in $ADSystemDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('IncrementalSync', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('IncrementalSync', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    If ($schedule.MinuteSpan -ne 0) {
                        $ADSDHash.Add('IncrementalSyncSchedule', $schedule.MinuteSpan)
                    }
                }
                'Full Sync Schedule' {
                    $SDFullSched = Convert-CMSchedule $prop.Value1
                    $ADSDHash.Add('FullSyncSchedule', "$(Get-HumanReadableSchedule -Schedule $SDFullSched)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADSDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
                'Enable Filtering Expired Logon' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('FilterExpiredLogon', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('FilterExpiredLogon', 'Disabled')
                    }
                }
                'Days Since Last Logon' {
                    $ADSDHash.Add('FilterExpiredLogonTime', $Prop.Value)
                }
                'Enable Filtering Expired Password' {
                    If ($Prop.Value -eq 1) {
                        $ADSDHash.Add('FilterExpiredPassword', 'Enabled')
                    }
                    Else {
                        $ADSDHash.Add('FilterExpiredPassword', 'Disabled')
                    }
                }
                'Days Since Last Password Set' {
                    $ADSDHash.Add('FilterExpiredPasswordTime', $Prop.Value)
                }
            }    
        }
        $ADContainerDiscovery = @()
        $AdditionalADAttributes = @()
        $ADContainerSearchCreds = @()
        $ADContainerExclusions = @()
        foreach ($List in $ADSystemDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { 
                                If ($value -eq 0) {
                                    $two = 'Yes'
                                }
                                else {
                                    $two = 'No'
                                }
                            }
                            3 {
                                If ($value -eq 1) {
                                    $three = 'Excluded'
                                }
                                else {
                                    $three = 'Included'
                                }
                                $ADContainerDiscovery += [pscustomobject]@{'Container' = $one; 'Recursive' = $two; 'Groups' = $three; 'Account' = 'Site Server'; 'Exclusions' = 'None' }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'AD Attributes' {
                    $AdditionalADAttributes += $List.values
                }
                'AD Containers Exclusions' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 {
                                $ADContainerExclusions += [pscustomobject]@{'Container' = $one; 'Exclusions' = $value }
                                Remove-Variable one
                                $start = 0
                            }
                        }
                    }
                }
                'AD Accounts:*' {
                    $AccountContainer = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADContainerSearchCreds += [pscustomobject]@{'Container' = $AccountContainer; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        foreach ($account in $ADContainerSearchCreds) {
            foreach ($Container in $ADContainerDiscovery) {
                If ($account.Container -like $Container.Container) {
                    $Container.Account = $account.Account
                }
            }
        }
        foreach ($Exclusion in $ADContainerExclusions) {
            foreach ($Container in $ADContainerDiscovery) {
                If ($Exclusion.Container -like $Container.Container) {
                    $Container.Exclusions = $Exclusion.Exclusions
                }
            }
        }
        $ADSDHash.Add('ActiveDirectoryContainers', $ADContainerDiscovery)
        $ADSDHash.Add('ActiveDirectoryAttributes', $AdditionalADAttributes)
        [PSCustomObject]$ADSDHash
    }
    #endregion AD System Discovery

    #region AD User Discovery
    If ($dmx -eq 4 -or $DMX -eq 7) {
        $ADUDHash = @{ }
        $ADUDHash.Add('DiscoveryMethod', 'Active Directory User Discovery')
        $ADUserDiscovery = Get-WmiObject -Query "SELECT * FROM SMS_SCI_Component WHERE FileType=2 AND ItemName='SMS_AD_USER_DISCOVERY_AGENT|SMS Site Server' AND ItemType='Component'" -Namespace "ROOT\SMS\site_$SiteName"
        foreach ($Prop in $ADUserDiscovery.Props) {
            #schedule and suches
            switch ($Prop.PropertyName) {
                'Enable Incremental Sync' {
                    If ($Prop.Value -eq 1) {
                        $ADUDHash.Add('IncrementalSync', 'Enabled')
                    }
                    Else {
                        $ADUDHash.Add('IncrementalSync', 'Disabled')
                    }
                }
                'Startup Schedule' {
                    $schedule = Convert-CMSchedule $prop.Value1
                    If ($schedule.MinuteSpan -ne 0) {
                        $ADUDHash.Add('IncrementalSyncSchedule', $schedule.MinuteSpan)
                    }
                }
                'Full Sync Schedule' {
                    $SDFullSched = Convert-CMSchedule $prop.Value1
                    $ADUDHash.Add('FullSyncSchedule', "$(Get-HumanReadableSchedule -Schedule $SDFullSched)")
                }
                'SETTINGS' {
                    If ($Prop.Value1 -eq 'ACTIVE') {
                        $ADUDHash.Add('DiscoveryState', 'Enabled')
                    }
                    Else {
                        $ADUDHash.Add('DiscoveryState', 'Disabled')
                    }
                }
            }    
        }
        $ADUserContainerDiscovery = @()
        $AdditionalADUserAttributes = @()
        $ADUserSearchCreds = @()
        $ADUserContainerExclusions = @()
        foreach ($List in $ADUserDiscovery.PropLists) {
            #Domains and Groups
            switch -wildcard ($List.PropertyListName) {
                'AD Containers' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 { 
                                If ($value -eq 0) {
                                    $two = 'Yes'
                                }
                                else {
                                    $two = 'No'
                                }
                            }
                            3 {
                                If ($value -eq 1) {
                                    $three = 'Excluded'
                                }
                                else {
                                    $three = 'Included'
                                }
                                $ADUserContainerDiscovery += [pscustomobject]@{'Container' = $one; 'Recursive' = $two; 'Groups' = $three; 'Account' = 'Site Server'; 'Exclusions' = 'None' }
                                Remove-Variable one, two, three
                                $start = 0
                            }
                        }
                    }
                }
                'AD Attributes' {
                    $AdditionalADUserAttributes += $List.values
                }
                'AD Containers Exclusions' {
                    $start = 0
                    foreach ($value in $List.values) {
                        $start++
                        switch ($start) {
                            1 { $one = $value }
                            2 {
                                $ADUserContainerExclusions += [pscustomobject]@{'Container' = $one; 'Exclusions' = $value }
                                Remove-Variable one
                                $start = 0
                            }
                        }
                    }
                }
                'AD Accounts:*' {
                    $AccountContainer = $List.PropertyListName -replace 'AD Accounts:', ''
                    $ADUserSearchCreds += [pscustomobject]@{'Container' = $AccountContainer; 'Account' = "$($List.Values[0])" }
                }
            }
        }
        foreach ($account in $ADUserSearchCreds) {
            foreach ($Container in $ADUserContainerDiscovery) {
                If ($account.Container -like $Container.Container) {
                    $Container.Account = $account.Account
                }
            }
        }
        foreach ($Exclusion in $ADUserContainerExclusions) {
            foreach ($Container in $ADUserContainerDiscovery) {
                If ($Exclusion.Container -like $Container.Container) {
                    $Container.Exclusions = $Exclusion.Exclusions
                }
            }
        }
        $ADUDHash.Add('ActiveDirectoryContainers', $ADUserContainerDiscovery)
        $ADUDHash.Add('ActiveDirectoryAttributes', $AdditionalADUserAttributes)
        [PSCustomObject]$ADUDHash
    }
    #endregion User Discovery
}