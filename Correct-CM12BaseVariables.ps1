
function Create-Array {
    param ($Key, [string]$Suffix)
    $objArray = New-Object System.Object
    $objArray | Add-Member -type NoteProperty -name Key -value $Key
    $objArray | Add-Member -type NoteProperty -name Suffix -value $Suffix
    $objArray
}
<#
if ($args.Count -eq 1)
    {
        $BaseVariableName = $args[0]
    }
elseif ($args.Count -eq 2)
    {
        $BaseVariableName = $args[0]
        $LengthSuffix = $args[1]
        
    }
#>
$global:BaseVariableList = @{}
$BaseVariableName = "BaseVar"
$LengthSuffix = 2


$objSMSTS = New-Object -ComObject Microsoft.SMS.TSEnvironment

$SMSTSVars = $objSMSTS.GetVariables()

foreach ($Var in $objSMSTS.GetVariables())
    {
        if ( $Var.ToUpper().Substring(0,$var.Length-$LengthSuffix) -eq $BaseVariableName)
            {
                
                #$BaseVariableList += $Var.ToUpper().Substring(0,$var.Length-$LengthSuffix)
                $BaseVariableList.Add($Var,$objSMSTS.Value($Var))
            }
    }
    
$BaseVariableList 


$VarCount = $BaseVariableList.Count
if ($VarCount -gt 1)
    {
        $X = 0
        $arr = @()
        foreach ($Key in $BaseVariableList)
            {
                #$arr.Add($Key, $Key.ToString().Substring($Key.ToString().Length - $LengthSuffix, $LengthSuffix))
                
                Write-Output ""
                $Key
                Write-Output ""
                #$arr.Key = $Key
                #$arr.Suffix = $Key.ToString().Substring($Key.ToString().Length - $LengthSuffix, $LengthSuffix)
                $arr += Create-Array -Key $Key -Suffix $Key.ToString().Substring($Key.ToString().Length - $LengthSuffix, $LengthSuffix)
            }

    }
    Write-Output ""
    $arr
