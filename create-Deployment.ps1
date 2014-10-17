

$ApplicationAssignmentClass = [wmiclass] "\\$($SMSProvider)\root\sms\site_$($SiteCode):SMS_ApplicationAssignment"

$newApplicationAssingment = $ApplicationAssignmentClass.CreateInstance()
$newApplicationAssingment.ApplicationName = "PDFCreator"
$newApplicationAssingment.AssignmentName = "Deploy PDFCreator"
$newApplicationAssingment.AssignedCIs                 = 16781957
$newApplicationAssingment.CollectionName                  = "All Desktops"
$newApplicationAssingment.CreationTime                    = "20130101120000.000000+***"
$newApplicationAssingment.LocaleID                        = 1043
$newApplicationAssingment.SourceSite                      = "PRI"
$newApplicationAssingment.StartTime                       = "20130101120000.000000+***"
$newApplicationAssingment.SuppressReboot                  = $true
$newApplicationAssingment.NotifyUser                      = $true
$newApplicationAssingment.TargetCollectionID              = "PRI0000C"
                 
$newApplicationAssingment.OfferTypeID = 2
$newApplicationAssingment.WoLEnabled  = $false
$newApplicationAssingment.RebootOutsideOfServiceWindows = $false
$newApplicationAssingment.OverrideServiceWindows  = $false
$newApplicationAssingment.UseGMTTimes = $true
[void] $newApplicationAssingment.Put()