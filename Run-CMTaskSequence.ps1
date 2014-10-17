$UI = New-Object -ComObject "UIResource.UIResourceMgr"

$UI.GetAvailableApplications()

$ProgramID = "*"
$PackageID = "PRI00061"

$UI.ExecuteProgram($ProgramID,$PackageID,$true)
