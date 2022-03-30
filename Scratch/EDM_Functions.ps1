Function Get-FilterEDM {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [xml]
        $EnhansedDetectionMethods,
        [Parameter(Mandatory=$true)]
        $RuleExpression,
        [Parameter(Mandatory=$false)]
        [hashtable]$DMSummary = @{}
    )
    foreach ($Expression in $RuleExpression) {
        If ($Expression.Operator -eq 'And') {
            Write-Verbose "adding an And"
            $DMSummary.Add('And',@{})
            $TheseDetails = Get-FilterEDM -EnhansedDetectionMethods $EnhansedDetectionMethods -RuleExpression $Expression.Operands.Expression
            foreach ($key in $TheseDetails.keys){
                $DMSummary.And.Add($key,$($TheseDetails.$key))
            }
        }
        ElseIf ($Expression.Operator -eq 'Or') {
            Write-Verbose "adding an Or"
            $DMSummary.Add('Or',@{})
            $TheseDetails = Get-FilterEDM -EnhansedDetectionMethods $EnhansedDetectionMethods -RuleExpression $Expression.Operands.Expression
            foreach ($key in $TheseDetails.keys){
                $DMSummary.Or.Add($key,$($TheseDetails.$key))
            }
        }
        Else {
            if ($DMSummary.keys -notcontains 'Settings'){
                $DMSummary.Add('Settings',@())
            }
            $SettingLogicalName = $Expression.Operands.SettingReference.SettingLogicalName
            Switch($Expression.Operands.SettingReference.SettingSourceType){
                'Registry'{
                    Write-Verbose "registry Setting"
                    $RegSetting = $EnhansedDetectionMethods.EnhancedDetectionMethod.Settings.SimpleSetting | Where-Object {$_.LogicalName -eq "$SettingLogicalName"}
                    $DMSummary.Settings += @{'RegSetting' = [PSCustomObject]@{
                            RegHive      = $RegSetting.RegistryDiscoverySource.Hive
                            RegKey       = $RegSetting.RegistryDiscoverySource.Key
                            RegValue     = $RegSetting.RegistryDiscoverySource.ValueName
                            Reg64Bit     = $RegSetting.RegistryDiscoverySource.Is64Bit
                            RegMethod    = $Expression.Operands.SettingReference.Method
                            RegData      = $Expression.Operands.ConstantValue.Value
                            RegDataList  = $Expression.Operands.ConstantValueList.ConstantValue.Value
                            RegDataType  = $Expression.Operands.SettingReference.DataType
                            RegOperator  = $Expression.Operator
                        }
                    }
                }
                'File'{
                    $FileSetting = $EnhansedDetectionMethods.EnhancedDetectionMethod.Settings.File | Where-Object {$_.LogicalName -eq "$SettingLogicalName"}
                    $DMSummary.Settings += @{'FileSetting' = [PSCustomObject]@{
                            ParentFolder             = $FileSetting.Path
                            FileName                 = $FileSetting.Filter
                            File64Bit                = $FileSetting.Is64Bit
                            FileOperator             = $Expression.Operator
                            FileMethod               = $Expression.Operands.SettingReference.Method
                            FileValueList            = $Expression.Operands.ConstantValueList.ConstantValue.Value
                            FileValue                = $Expression.Operands.ConstantValue.Value
                            FilePropertyName         = $Expression.Operands.SettingReference.PropertyPath
                            FilePropertyNameDataType = $Expression.Operands.SettingReference.DataType
                        }
                    }
                }
                'Folder'{
                    $FolderSetting = $EnhansedDetectionMethods.EnhancedDetectionMethod.Settings.Folder | Where-Object {$_.LogicalName -eq "$SettingLogicalName"}
                    $DMSummary.Settings += @{'FolderSetting' = [PSCustomObject]@{
                            ParentFolder               = $FolderSetting.Path
                            FolderName                 = $FolderSetting.Filter
                            Folder64Bit                = $FolderSetting.Is64Bit
                            FolderOperator             = $Expression.Operator
                            FolderMethod               = $Expression.Operands.SettingReference.Method
                            FolderValueList            = $Expression.Operands.ConstantValueList.ConstantValue.Value
                            FolderValue                = $Expression.Operands.ConstantValue.Value
                            FolderPropertyName         = $Expression.Operands.SettingReference.PropertyPath
                            FolderPropertyNameDataType = $Expression.Operands.SettingReference.DataType
                        }
                    }
                }
                'MSI'{
                    $MSISetting = $EnhansedDetectionMethods.EnhancedDetectionMethod.Settings.MSI | Where-Object {$_.LogicalName -eq "$SettingLogicalName"}
                    if ($Expression.Operands.SettingReference.DataType -eq 'Int64'){ #Existensile detection
                        Write-Verbose "MSI Exists on System"
                        #$MSIDetection = "MSI Exists on System"
                    } elseif ($Expression.Operands.SettingReference.DataType -eq 'Version') { #Exists plus is a specific version of MSI
                        Write-Verbose "MSI Version is..."
                        #$MSIOperator = "The MSI $MSIDataType is $(Convert-Operator $Expression.Operator) [$MSIVersion]."
                    } Else {
                        Write-Verbose "Unknown MSI Configuration for product code."
                    }
                    $DMSummary.Settings += @{'MsiSetting' = [PSCustomObject]@{
                            MSIProductCode     = $MSISetting.ProductCode
                            MSIDataType        = $Expression.Operands.SettingReference.DataType
                            MSIMethod          = $Expression.Operands.SettingReference.Method
                            MSIDataValue       = $Expression.Operands.ConstantValue.Value
                            MSIPropertyName    = $Expression.Operands.SettingReference.PropertyPath
                            MSIOperator        = $Expression.Operator
                        }
                    }
                }
            }
        }
    }
    Return $DMSummary
}


Function Convert-Operator {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Operator,
        [Parameter(Mandatory=$false)]
        [string]
        $Data,
        [Parameter(Mandatory=$false)]
        [string[]]
        $DataList
    )
    switch ($Operator) {
        Equals { $OperationText = "is equals to: $Data" }
        NotEquals { $OperationText = "is not equal to: $Data" }
        GreaterEquals { $OperationText = "is greater than or equal to: $Data" }
        LessThan { $OperationText = "is less than: $Data" }
        LessEquals { $OperationText = "is less than or equal to: $Data" }
        GreaterThan { $OperationText = "is greater than: $Data" }
        Between { $OperationText = "is between: $($DataList[0]) and $($DataList[1])" }
        OneOf { $OperationText = "is One Of the following: $($DataList -join '; ')" }
        NoneOf { $OperationText = "is None Of the following: $($DataList -join '; ')" }
        default { $OperationText = "Unknown operation for data." }
    }
    Write-Output $OperationText
}

function Write-EDMs{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$EDMHash
    )
    foreach ($key in $EDMHash.keys){
        if ($key -like 'Or'){
            Write-Output '<div class="EdmOr">'
            Write-Output 'Or'
            Write-Verbose "One of these conditions (Or)"
            Write-EDMs $EDMHash.$key
            Write-Output '</div>'
            Write-Verbose "End (Or)"
        }
        elseif ($key -like 'And'){
            Write-Output '<div class="EdmAnd">'
            Write-Output 'And'
            Write-Verbose "All of these conditions (And)"
            Write-EDMs $EDMHash.$key
            Write-Output '</div>'
            Write-Verbose "End (And)"
        }
        elseif ($Key -like 'Settings') {
            foreach ($Setting in $EDMHash.Settings){
                Write-output '<div class="EdmSetting">'
                $Values = $setting.Values[0].ForEach({$_})[0]
                switch ($($Setting.Keys.ForEach({$_})[0])) {
                    'RegSetting' {
                        Write-output "Registry Key: $($values.RegHive)\$($values.RegKey)<br />"
                        Write-output "&bull; Registry Value Name: $($values.RegValue)<br />"
                        If ($Values.RegDataType -eq "Boolean"){
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; Reg key/value exists.<br />"
                        } else {
                            Write-Output "&bull; Value Data Type: $($Values.RegDataType)<br />"
                            $Operation = Convert-Operator -Operator $Values.RegOperator -Data "$($Values.RegData)" -DataList $($Values.RegDataList)
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; $($Operation)<br />"
                        }
                        If ($Values.Reg64Bit -eq "false"){
                            Write-output "&#9745; Reg Key Associated with 32-bit App on 64-bit system" ##Checked
                        } else {
                            Write-output "&#9744; Reg Key Associated with 32-bit App on 64-bit system" ##Unchecked
                        }
                    }
                    'FileSetting' {
                        Write-output "File System File: $($values.ParentFolder)\$($values.FileName)<br />"
                        If ($Values.FilePropertyNameDataType -eq "Int64" -and $Values.FileMethod -eq "Count" -and $Values.FileValue -eq 0){
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; File exists.<br />"
                        } else {
                            Write-output "&bull; File Property Name: $($values.FilePropertyName)<br />"
                            Write-Output "&bull; File Property Data Type: $($Values.FilePropertyNameDataType)<br />"
                            $Operation = Convert-Operator -Operator $Values.FileOperator -Data "$($Values.FileValue)" -DataList $($Values.FileValueList)
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; $($Operation)<br />"
                        }
                        If ($Values.File64Bit -eq "false"){
                            Write-output "&#9745; File/Folder associated with 32-bit App on 64-bit system" ##Checked
                        } else {
                            Write-output "&#9744; File/Folder associated with 32-bit App on 64-bit system" ##Unchecked
                        }
                    }
                    'FolderSetting' {
                        Write-output "File System Folder: $($values.ParentFolder)\$($values.FolderName)<br />"
                        If ($Values.FolderPropertyNameDataType -eq "Int64" -and $Values.FolderMethod -eq "Count" -and $Values.FolderValue -eq 0){
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; Folder exists.<br />"
                        } else {
                            Write-output "&bull; File Property Name: $($values.FolderPropertyName)<br />"
                            Write-Output "&bull; File Property Data Type: $($Values.FolderPropertyNameDataType)<br />"
                            $Operation = Convert-Operator -Operator $Values.FolderOperator -Data "$($Values.FolderValue)" -DataList $($Values.FolderValueList)
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; $($Operation)<br />"
                        }
                        If ($Values.Folder64Bit -eq "false"){
                            Write-output "&#9745; File/Folder associated with 32-bit App on 64-bit system" ##Checked
                        } else {
                            Write-output "&#9744; File/Folder associated with 32-bit App on 64-bit system" ##Unchecked
                        }
                    }
                    'MsiSetting' {
                        Write-output "MSI Product Code: $($values.MSIProductCode)<br />"
                        If ($Values.MSIDataType -eq "Int64" -and $Values.MSIMethod -eq "Count" -and $Values.MSIDataValue -eq 0){
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; MSI exists.<br />"
                        } else {
                            Write-output "&bull; MSI Property Name: $($values.MSIPropertyName)<br />"
                            Write-Output "&bull; MSI Property Data Type: $($Values.MSIDataType)<br />"
                            $Operation = Convert-Operator -Operator $Values.MSIOperator -Data "$($Values.MSIDataValue)"
                            Write-Output "&nbsp;&nbsp;&nbsp;&bull; $($Operation)<br />"
                        }
                    }
                    Default {
                        $SettingDetails = 'Unknown Detection Setting.'
                    }
                }
                Write-Output $SettingDetails
                #Write-output "$($Setting.Keys)"
                Write-output '</div>'
                Write-Verbose "These Settings: $($Setting.Keys)"
            }
        }
    }
}


$EDM = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_f812ada8-6ee5-418d-aae4-c4186abbb22d" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>SOFTWARE\Microsoft\Office\ClickToRun\Configuration</Key><ValueName>ClientVersionToReport</ValueName></RegistryDiscoverySource></SimpleSetting><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_76191ae3-f784-40e3-993d-570b912c64d1" DataType="String"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>SOFTWARE\Microsoft\Office\ClickToRun\Configuration</Key><ValueName>ProductReleaseIds</ValueName></RegistryDiscoverySource></SimpleSetting><MSI xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="MSI_b64575e3-e8b9-4f10-9df7-67f2d608131a" IsPerUser="false"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><ProductCode>{D4DC0E96-E2EA-4BC1-996F-BA346DDD8EA6}</ProductCode></MSI><MSI xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="MSI_31dcdb36-c724-40d0-b1e3-c8e37edf02da" IsPerUser="false"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><ProductCode>{B9274744-8BAE-4874-8E59-2610919CD419}</ProductCode></MSI></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="IsInstalledRule" Severity="None" NonCompliantWhenSettingIsNotFound="false"><Expression><Operator>Or</Operator><Operands><Expression IsGroup="true"><Operator>And</Operator><Operands><Expression><Operator>GreaterEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="1" DataType="Version" SettingLogicalName="RegSetting_f812ada8-6ee5-418d-aae4-c4186abbb22d" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValue Value="16.0.10730.20334" DataType="Version" /></Operands></Expression><Expression><Operator>Contains</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="1" DataType="String" SettingLogicalName="RegSetting_76191ae3-f784-40e3-993d-570b912c64d1" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValue Value="O365ProPlusRetail" DataType="String" /></Operands></Expression><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="46" DataType="Int64" SettingLogicalName="MSI_31dcdb36-c724-40d0-b1e3-c8e37edf02da" SettingSourceType="MSI" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression></Operands></Expression><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="45" DataType="Int64" SettingLogicalName="MSI_b64575e3-e8b9-4f10-9df7-67f2d608131a" SettingSourceType="MSI" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMAndOnly = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_651ec3a4-b0be-413b-934d-3b4b4888d4bc" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>SOFTWARE\Microsoft\Office\ClickToRun\Configuration</Key><ValueName>ClientVersionToReport</ValueName></RegistryDiscoverySource></SimpleSetting><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_a9e86cd8-bd04-4b5e-9d5c-0f1b72cd9b73" DataType="String"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>SOFTWARE\Microsoft\Office\ClickToRun\Configuration</Key><ValueName>ProductReleaseIds</ValueName></RegistryDiscoverySource></SimpleSetting></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="IsInstalledRule" Severity="None" NonCompliantWhenSettingIsNotFound="false"><Expression><Operator>And</Operator><Operands><Expression><Operator>GreaterEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="1" DataType="Version" SettingLogicalName="RegSetting_651ec3a4-b0be-413b-934d-3b4b4888d4bc" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValue Value="16.0.10730.20334" DataType="Version" /></Operands></Expression><Expression><Operator>Contains</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_9c451eb1-d976-4515-b840-0a208848e32d" Version="1" DataType="String" SettingLogicalName="RegSetting_a9e86cd8-bd04-4b5e-9d5c-0f1b72cd9b73" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValue Value="O365ProPlusRetail" DataType="String" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMExists =  [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><File xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>c:\temp</Path><Filter>bob.txt</Filter></File><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_d597c17f-1242-4df1-8f79-fa7f6bc847b7" DataType="String"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>Software\Time</Key><ValueName>Start</ValueName></RegistryDiscoverySource></SimpleSetting></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Int64" SettingLogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad" SettingSourceType="File" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression><Expression><Operator>Equals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Boolean" SettingLogicalName="RegSetting_d597c17f-1242-4df1-8f79-fa7f6bc847b7" SettingSourceType="Registry" Method="Value" PropertyPath="RegistryValueExists" Changeable="false" /><ConstantValue Value="true" DataType="Boolean" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMBetween = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><File xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>c:\temp</Path><Filter>bob.txt</Filter></File><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_de35ddd5-3005-4cb3-8d84-f45162f61526" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>Software\Time</Key><ValueName>Start</ValueName></RegistryDiscoverySource></SimpleSetting></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>Between</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="DateTime" SettingLogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad" SettingSourceType="File" Method="Value" PropertyPath="DateModified" Changeable="false" /><ConstantValueList DataType="DateTimeArray"><ConstantValue Value="2021-03-11T06:00:00Z" DataType="DateTime" /><ConstantValue Value="2021-04-14T05:00:00Z" DataType="DateTime" /></ConstantValueList></Operands></Expression><Expression><Operator>Between</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Version" SettingLogicalName="RegSetting_de35ddd5-3005-4cb3-8d84-f45162f61526" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValueList DataType="VersionArray"><ConstantValue Value="1.2.3.4" DataType="Version" /><ConstantValue Value="2.3.4.5" DataType="Version" /></ConstantValueList></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMWithReg= [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><File xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>c:\temp</Path><Filter>bob.txt</Filter></File><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_b24283ee-8f6a-459d-af2d-2a74f8130271" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="false" CreateMissingPath="true"><Key>software\Google</Key><ValueName>Chrome</ValueName></RegistryDiscoverySource></SimpleSetting><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_5d3b63e9-8d14-4845-bcff-78e3faae26f6" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>Software\Time</Key><ValueName>Start</ValueName></RegistryDiscoverySource></SimpleSetting></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>Between</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="DateTime" SettingLogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad" SettingSourceType="File" Method="Value" PropertyPath="DateModified" Changeable="false" /><ConstantValueList DataType="DateTimeArray"><ConstantValue Value="2021-03-11T06:00:00Z" DataType="DateTime" /><ConstantValue Value="2021-04-14T05:00:00Z" DataType="DateTime" /></ConstantValueList></Operands></Expression><Expression><Operator>OneOf</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Version" SettingLogicalName="RegSetting_5d3b63e9-8d14-4845-bcff-78e3faae26f6" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValueList DataType="VersionArray"><ConstantValue Value="1.2.3.4" DataType="Version" /><ConstantValue Value="2.3.4.5" DataType="Version" /></ConstantValueList></Operands></Expression><Expression><Operator>Equals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="11" DataType="Boolean" SettingLogicalName="RegSetting_b24283ee-8f6a-459d-af2d-2a74f8130271" SettingSourceType="Registry" Method="Value" PropertyPath="RegistryValueExists" Changeable="false" /><ConstantValue Value="true" DataType="Boolean" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMWithFile = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><File xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>c:\temp</Path><Filter>bob.txt</Filter></File><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_b24283ee-8f6a-459d-af2d-2a74f8130271" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="false" CreateMissingPath="true"><Key>software\Google</Key><ValueName>Chrome</ValueName></RegistryDiscoverySource></SimpleSetting><SimpleSetting xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="RegSetting_5d3b63e9-8d14-4845-bcff-78e3faae26f6" DataType="Version"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><RegistryDiscoverySource Hive="HKEY_LOCAL_MACHINE" Depth="Base" Is64Bit="true" CreateMissingPath="true"><Key>Software\Time</Key><ValueName>Start</ValueName></RegistryDiscoverySource></SimpleSetting></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Int64" SettingLogicalName="File_b2e2fa76-5cb3-4ef7-ae3a-228f19f92cad" SettingSourceType="File" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression><Expression><Operator>OneOf</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="9" DataType="Version" SettingLogicalName="RegSetting_5d3b63e9-8d14-4845-bcff-78e3faae26f6" SettingSourceType="Registry" Method="Value" Changeable="false" /><ConstantValueList DataType="VersionArray"><ConstantValue Value="1.2.3.4" DataType="Version" /><ConstantValue Value="2.3.4.5" DataType="Version" /></ConstantValueList></Operands></Expression><Expression><Operator>Equals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="11" DataType="Boolean" SettingLogicalName="RegSetting_b24283ee-8f6a-459d-af2d-2a74f8130271" SettingSourceType="Registry" Method="Value" PropertyPath="RegistryValueExists" Changeable="false" /><ConstantValue Value="true" DataType="Boolean" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMFolderMSI = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Folder xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="Folder_13f9b529-d5e3-4d66-9752-fc92cfac35f1"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>C:\Users</Path><Filter>Public</Filter></Folder><MSI xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="MSI_a760a8e9-b33e-4ca5-bc88-20dc38c2e9b3" IsPerUser="false"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><ProductCode>{23170F69-40C1-2702-1900-000001000000}</ProductCode></MSI></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="14" DataType="Int64" SettingLogicalName="Folder_13f9b529-d5e3-4d66-9752-fc92cfac35f1" SettingSourceType="Folder" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="14" DataType="Int64" SettingLogicalName="MSI_a760a8e9-b33e-4ca5-bc88-20dc38c2e9b3" SettingSourceType="MSI" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'

$EDMFolderMSI = [xml]'<EnhancedDetectionMethod xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Settings xmlns="http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest"><Folder xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" Is64Bit="false" LogicalName="Folder_13f9b529-d5e3-4d66-9752-fc92cfac35f1"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><Path>C:\Users</Path><Filter>Public</Filter></Folder><MSI xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/07/10/DesiredConfiguration" LogicalName="MSI_a760a8e9-b33e-4ca5-bc88-20dc38c2e9b3" IsPerUser="false"><Annotation xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules"><DisplayName Text="" /><Description Text="" /></Annotation><ProductCode>{23170F69-40C1-2702-1900-000001000000}</ProductCode></MSI></Settings><Rule xmlns="http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules" id="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A/DeploymentType_9095159c-06eb-4f9d-886b-8cf01dd72b6c" Severity="Informational" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="" /><Description Text="" /></Annotation><Expression><Operator>And</Operator><Operands><Expression><Operator>NotEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="14" DataType="Int64" SettingLogicalName="Folder_13f9b529-d5e3-4d66-9752-fc92cfac35f1" SettingSourceType="Folder" Method="Count" Changeable="false" /><ConstantValue Value="0" DataType="Int64" /></Operands></Expression><Expression><Operator>GreaterEquals</Operator><Operands><SettingReference AuthoringScopeId="ScopeId_401E747F-ACAE-4042-B5AA-1D32866BBD3A" LogicalName="Application_4aebcd45-6a66-4f62-a00e-7c1b4f0a7c3b" Version="15" DataType="Version" SettingLogicalName="MSI_a760a8e9-b33e-4ca5-bc88-20dc38c2e9b3" SettingSourceType="MSI" Method="Value" PropertyPath="ProductVersion" Changeable="false" /><ConstantValue Value="1.3.5.4" DataType="Version" /></Operands></Expression></Operands></Expression></Rule></EnhancedDetectionMethod>'
