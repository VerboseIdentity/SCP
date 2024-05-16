function Summary_export_DB{

$Quick_export = [PSCustomobject]@{ Server_Name = $server.Split(",")[0]; Type = $server.Split(",")[1]}
if($Global:SQL_data.OS_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName OS_review -NotePropertyValue $Global:SQL_data.OS_Valuation}
else{ Add-Member -InputObject $Quick_export -NotePropertyName OS_review -NotePropertyValue "PASSED"}

if($Global:SQL_data.RAM_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName RAM_review -NotePropertyValue $Global:SQL_data.RAM_Valuation}
else{ Add-Member -InputObject $Quick_export -NotePropertyName RAM_review -NotePropertyValue "PASSED"}

if($Global:SQL_data.CPU_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName CPU_review -NotePropertyValue $Global:SQL_data.CPU_Valuation}
else{ Add-Member -InputObject $Quick_export -NotePropertyName CPU_review -NotePropertyValue "PASSED"}

if($Global:SQL_data.Drive_Valuation -eq "FAILED"){$failure_omitter = 1; Add-Member -InputObject $Quick_export -NotePropertyName Drive_review -NotePropertyValue $Global:SQL_data.Drive_Valuation}
else{ Add-Member -InputObject $Quick_export -NotePropertyName Drive_review -NotePropertyValue "PASSED"}

if($Global:SQL_data.SQL_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName SQL_review -NotePropertyValue $Global:SQL_data.SQL_Valuation}
else{Add-Member -InputObject $Quick_export -NotePropertyName SQL_review -NotePropertyValue "PASSED"}

if($failure -eq 1 -or $failure_omitter -eq 1){Add-Member -InputObject $Quick_export -NotePropertyName OverAll_review -NotePropertyValue "FAILED"}
else{Add-Member -InputObject $Quick_export -NotePropertyName OverAll_review -NotePropertyValue "PASSED"}

$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK

$Quick_export | Export-Excel -Path ".\Output.xlsx" -AutoSize -Append -ConditionalText $positive_condition,$Negative_condition -BoldTopRow -WorksheetName "Quick_reqference"



<#*************************************************************** Summary *******************************************************************************************************************************#>

$data = @( 
    
    [PSCustomObject]@{ Server_Name = $server.Split(",")[0]; Categories = 'OS_Review'; Current_value = $Global:SQL_data.OS_Available ; Requirement = $Global:SQL_data.OS_Requirement; SES_Valuation = $Global:SQL_data.OS_valuation},
    [PSCustomObject]@{ Categories = 'RAM_Review'; Current_value = $Global:SQL_data.RAM_Available_GB ; Requirement = $Global:SQL_data.RAM_Requirement_GB; SES_Valuation = $Global:SQL_data.RAM_valuation},
    [PSCustomObject]@{ Categories = 'CPU_Review'; Current_value = $Global:SQL_data.CPU_Available ; Requirement = $Global:SQL_data.CPU_Requirement; SES_Valuation = $Global:SQL_data.CPU_valuation},
    [PSCustomObject]@{ Categories = 'SQL_Review'; Current_value = $Global:SQL_data.SQL_Version_Available ; Requirement = $Global:SQL_data.SQL_version_Requirements; SES_Valuation = $Global:SQL_data.SQL_valuation},
    [PSCustomObject]@{ Categories = 'Overall_review'; Current_value = '' ; Requirement = ''; SES_Valuation = if($failure -eq 1){"FAILED"}else{"PASSED"}},
    [PSCustomObject]@{ Categories = ' '; Current_value = ' ' ; Requirement = ' '; SES_Valuation = ' '},
    [PSCustomObject]@{ Categories = ' '; Current_value = ' ' ; Requirement = ' '; SES_Valuation = ' '},
    [PSCustomObject]@{ Server_Name = 'Server_Name'; Categories = 'Categories'; Current_value = 'Current_value' ; Requirement = 'Requirement'; SES_Valuation = 'SES_Valuation'}

)
$Headers_Server_Name = New-ConditionalText "Server_Name" -BackgroundColor Cyan -ConditionalTextColor BLACK
$Headers_Categories = New-ConditionalText "Categories" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_Current_value = New-ConditionalText "Current_value" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_Requirement = New-ConditionalText "Requirement" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_SES_Valuation = New-ConditionalText "SES_Valuation" -BackgroundColor Yellow -ConditionalTextColor BLACK


$Header = [PSCustomObject]@{ Server_Name = 'Server_Name'; Categories = 'Categories'; Current_value = 'Current_value' ; Requirement = 'Requirement'; SES_Valuation = 'SES_Valuation'}
$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK

$data | Export-Excel -Path ".\Output.xlsx" -Append -AutoSize -ConditionalText $positive_condition,$Negative_condition,$Headers,$Headers_Server_Name,$Headers_Categories,$Headers_Current_value,$Headers_Requirement,$Headers_SES_Valuation -BoldTopRow -WorksheetName "Summary"


}


function Summary_export_NonDB{

$Quick_export = [PSCustomobject]@{ Server_Name = $server.Split(",")[0]; Type = $server.Split(",")[1]}
if($Global:myData.OS_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName OS_review -NotePropertyValue $Global:myData.OS_Valuation}
else{Add-Member -InputObject $Quick_export -NotePropertyName OS_review -NotePropertyValue "PASSED"}

if($Global:myData.RAM_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName RAM_review -NotePropertyValue $Global:myData.RAM_Valuation}
else{Add-Member -InputObject $Quick_export -NotePropertyName RAM_review -NotePropertyValue "PASSED"}

if($Global:myData.CPU_Valuation -eq "FAILED"){$failure = 1; Add-Member -InputObject $Quick_export -NotePropertyName CPU_review -NotePropertyValue $Global:myData.CPU_Valuation}
else{Add-Member -InputObject $Quick_export -NotePropertyName CPU_review -NotePropertyValue "PASSED"}

if($Global:myData.Drive_Valuation -eq "FAILED"){$failure_omitter = 1; Add-Member -InputObject $Quick_export -NotePropertyName Drive_review -NotePropertyValue $Global:myData.Drive_Valuation}
else{Add-Member -InputObject $Quick_export -NotePropertyName Drive_review -NotePropertyValue "PASSED"}

if($failure -eq 1 -or $failure_omitter -eq 1){Add-Member -InputObject $Quick_export -NotePropertyName OverAll_review -NotePropertyValue "FAILED"}
else{Add-Member -InputObject $Quick_export -NotePropertyName OverAll_review -NotePropertyValue "PASSED"}


$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK

$Quick_export | Export-Excel -Path ".\Output.xlsx" -AutoSize -Append -ConditionalText $positive_condition,$Negative_condition -BoldTopRow -WorksheetName "Quick_reference"


<#*************************************************************** Summary *******************************************************************************************************************************#>

$data = @( 
    
    [PSCustomObject]@{ Server_Name = $server.Split(",")[0]; Categories = 'OS_Review'; Current_value = $Global:myData.OS_Available ; Requirement = $Global:myData.OS_Requirement; SES_Valuation = $Global:myData.OS_valuation},
    [PSCustomObject]@{ Categories = 'RAM_Review'; Current_value = $Global:myData.RAM_Available_GB ; Requirement = $Global:myData.RAM_Requirement_GB; SES_Valuation = $Global:myData.RAM_valuation},
    [PSCustomObject]@{ Categories = 'CPU_Review'; Current_value = $Global:myData.CPU_Available ; Requirement = $Global:myData.CPU_Requirement; SES_Valuation = $Global:myData.CPU_valuation},
    [PSCustomObject]@{ Categories = 'Overall_review'; Current_value = '' ; Requirement = ''; SES_Valuation = if($failure -eq 1){"FAILED"}else{"PASSED"}},
    [PSCustomObject]@{ Categories = ' '; Current_value = ' ' ; Requirement = ' '; SES_Valuation = ' '},
    [PSCustomObject]@{ Categories = ' '; Current_value = ' ' ; Requirement = ' '; SES_Valuation = ' '},
    [PSCustomObject]@{ Server_Name = 'Server_Name'; Categories = 'Categories'; Current_value = 'Current_value' ; Requirement = 'Requirement'; SES_Valuation = 'SES_Valuation'}

)
$Headers_Server_Name = New-ConditionalText "Server_Name" -BackgroundColor Cyan -ConditionalTextColor BLACK
$Headers_Categories = New-ConditionalText "Categories" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_Current_value = New-ConditionalText "Current_value" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_Requirement = New-ConditionalText "Requirement" -BackgroundColor Yellow -ConditionalTextColor BLACK
$Headers_SES_Valuation = New-ConditionalText "SES_Valuation" -BackgroundColor Yellow -ConditionalTextColor BLACK


$Header = [PSCustomObject]@{ Server_Name = 'Server_Name'; Categories = 'Categories'; Current_value = 'Current_value' ; Requirement = 'Requirement'; SES_Valuation = 'SES_Valuation'}
$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK


if($failure -eq 1){

$data | Export-Excel -Path ".\Output.xlsx" -Append -AutoSize -ConditionalText $positive_condition,$Negative_condition,$Headers,$Headers_Server_Name,$Headers_Categories,$Headers_Current_value,$Headers_Requirement,$Headers_SES_Valuation -BoldTopRow -WorksheetName "Failure_Summary"

}

$data | Export-Excel -Path ".\Output.xlsx" -Append -AutoSize -ConditionalText $positive_condition,$Negative_condition,$Headers,$Headers_Server_Name,$Headers_Categories,$Headers_Current_value,$Headers_Requirement,$Headers_SES_Valuation -BoldTopRow -WorksheetName "Summary"


}

