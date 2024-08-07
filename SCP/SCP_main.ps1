﻿cls

#SES_requirements
#################

$Web = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 28
Proc = 8}

$Msg = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 4
Proc = 2}

$Fax = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 8
Proc = 2}

$Prt = @{ OS  = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 4
Proc = 2}

$Acdm = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 4
Proc = 2}

$Scan = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 4
Proc = 2}

$Unity = @{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
RAM = 8
Proc = 2}

$DB =@{ OS = "Windows Server 2019 Standard","Windows Server 2019 Datacenter","Windows Server 2022 Standard","Windows Server 2022 Datacenter","Windows Server 2022 Enterprise","Windows Server 2019 Enterprise"
}


################################# SQL Connection and Query ####################################################

function DB_connection {

    $serverInstance = Read-Host "`nEnter the SQL Listener or Instance Name"
    $databaseName = "Works"
    $queries = @{
        SQLVersion = "SELECT @@VERSION"
        Active_Users = "SELECT COUNT(*) FROM works..IDX_User WHERE IsInactiveFLAG = 'N'"
        Encounters = "SELECT COUNT(*) AS Last360Days FROM Works..Encounter WITH (NOLOCK) WHERE dttm >= DATEADD(day, -360, GETDATE()) AND dttm <= GETDATE()"
        Appointments = "SELECT COUNT(*) AS Last360Days FROM Works..Appointment WITH (NOLOCK) WHERE Startdttm >= DATEADD(day, -360, GETDATE()) AND Startdttm <= GETDATE() AND AppointmentStatusDE NOT IN (3, 6, 9, 5, 7)"
        Works_Size = "USE Works; SELECT SUM(CAST(size AS bigint) * 8 / 1024) AS TotalSizeMB FROM sys.master_files WHERE database_id = DB_ID('Works');"
        SQL_Valuation = "DECLARE @sqlstatement NVARCHAR(MAX)
        SET @sqlstatement = '
        DECLARE @ProductMajorVersion INT,
                @ProductName VARCHAR(255),
                @ProductLevel VARCHAR(10),
                @ProductEdition VARCHAR(20),
                @ProductUpdateLevel VARCHAR(10),
                @ProductCUNumber INT
        
        SELECT @ProductMajorVersion = TRY_CONVERT(INT, SERVERPROPERTY(''ProductMajorVersion'')),
            @ProductLevel = LEFT(CAST(SERVERPROPERTY(''ProductLevel'') AS VARCHAR), 10),
            @ProductUpdateLevel = CAST(SERVERPROPERTY(''ProductUpdateLevel'') AS VARCHAR)
        
        SELECT @ProductName = CASE @ProductMajorVersion
                                WHEN 11 THEN ''SQL Server 2012''
                                WHEN 12 THEN ''SQL Server 2014''
                                WHEN 13 THEN ''SQL Server 2016''
                                WHEN 14 THEN ''SQL Server 2017''
                                WHEN 15 THEN ''SQL Server 2019''
                                WHEN 16 THEN ''SQL Server 2022''
                            END,
            @ProductEdition = CASE CAST(SERVERPROPERTY(''EngineEdition'') AS INT)
                                    WHEN 2 THEN ''Standard''
                                    WHEN 3 THEN ''Enterprise''
                                    WHEN 8 THEN ''Managed Instance''
                                    ELSE ''Other''
                                END
        
        SET @ProductCUNumber = ISNULL(TRY_CONVERT(INT, SUBSTRING(@ProductUpdateLevel, 3, 3)), 0)
        
        IF @ProductEdition = ''Managed Instance''
            SET @ProductName = ''SQL Server Managed Instance''
        ELSE
            SET @ProductName = CONCAT(@ProductName, '' '', @ProductLevel, '' '', @ProductUpdateLevel, '' '', @ProductEdition + '' Edition'')
        
        DECLARE @ErrorMessage VARCHAR(1000),
                @CRLF VARCHAR(2) = CHAR(13) + CHAR(10),
                @Msg VARCHAR(100)
        
        SET @Msg = ''You are on a supported SQL platform: SQL Server 2019 CU5+ or SQL Server 2022 CU11+''
        
        IF (@ProductMajorVersion = 15 AND @ProductCUNumber >= 5 AND @ProductEdition IN (''Enterprise'', ''Standard''))
        OR (@ProductMajorVersion = 16 AND @ProductCUNumber >= 11 AND @ProductEdition IN (''Enterprise'', ''Standard''))
        OR (@ProductMajorVersion > 16 AND @ProductEdition IN (''Enterprise'', ''Standard''))
        OR @ProductEdition = ''Managed Instance''
            SELECT @Msg AS Msg
        ELSE
            BEGIN
                SET @ErrorMessage = ''Invalid SQL Server version or edition: '' + @ProductName
                SELECT @ErrorMessage AS Msg
            END'
        
        EXEC sp_executesql @sqlstatement
        "
        
        Total_Spooler_Count = "SELECT COUNT(DISTINCT SPOOLER_NM) AS 'Total Spoolers' FROM Works..Css_Job_Queue WHERE SPOOLER_NM <> '' AND CONVERT(DATE, DATETIME_REC_CREATE) = CONVERT(DATE, GETDATE()) AND Job_Status_CD = 3 AND JOB_TYPE_CD IN ('Print', 'Fax')"
        JobCount = "SELECT CONVERT(DATE, DATETIME_REC_CREATE) AS 'Current Date', COUNT(*) AS 'count' FROM Works..Css_Job_Queue WHERE SPOOLER_NM <> '' AND CONVERT(DATE, DATETIME_REC_CREATE) = CONVERT(DATE, GETDATE()) AND Job_Status_CD = 3 AND JOB_TYPE_CD IN ('Print', 'Fax') GROUP BY CONVERT(DATE, DATETIME_REC_CREATE) ORDER BY COUNT(*) DESC"
    }

    $connectionString = "Server=$serverInstance;Database=$databaseName;Integrated Security=True;"

    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        $global:SQL_reslts = @{}
        foreach ($key in $queries.Keys) {
            $command = $connection.CreateCommand()
            $command.CommandText = $queries[$key]
            $global:SQL_reslts[$key] = $command.ExecuteScalar()
        }
    }
    catch {
        Write-Host "An error occured:`n$_"
    }
    finally{
        $connection.Close()
    }

}

################################ Comparision and output ########################################################

function display{

param($current_value, $actual_value)

Write-Host "`n*********************" ($Server.split(",")[0]) "*************************"
if ($current_value.OS -notin $actual_value.OS){Write-Host "OS-FAILED :"$current_value.OS "is Incompatible" -ForegroundColor Red}else{Write-Host "OS-PASSED :",$current_value.OS,"is compatible" -ForegroundColor Green}
if ([int]$current_value.RAM -ge [int]$actual_value.RAM){Write-Host "RAM-PASSED : "$current_value.RAM "GB available |",$actual_value.RAM ,"RAM required" -ForegroundColor Green}else{Write-Host "RAM-FAILED : "$current_value.RAM "GB available |",$actual_value.RAM ,"GB required" -ForegroundColor Red}
if ([int]$current_value.Proc -ge [int]$actual_value.Proc){Write-Host "CPU-PASSED : "$current_value.Proc "Cores available |",$actual_value.Proc ,"cores required" -ForegroundColor Green}else{Write-Host "CPU-FAILED : "$current_value.Proc "Cores available |",$actual_value.Proc ,"cores required" -ForegroundColor Red}
}

################################ SQL_DB Comparision and output #################################################

function Sql_display{

param($current_value, $OS, $RAM, $Proc)

#$global:SQL_Version

Write-Host "`n*********************" ($Server.split(",")[0]) "*************************"
if ($current_value.OS -notin $OS.OS){Write-Host "OS-FAILED :"$current_value.OS "is Incompatible" -ForegroundColor Red}else{Write-Host "OS-PASSED :",$current_value.OS,"is compatible" -ForegroundColor Green}
if ([int]$current_value.RAM -ge [int]$RAM){Write-Host "RAM-PASSED : "$current_value.RAM "GB available |",$RAM,"GB RAM required for",$Works_size,"TB of Works_DB" -ForegroundColor Green}else{Write-Host "RAM-FAILED : "$current_value.RAM "GB available |",$RAM,"GB required for",$Works_size,"TB of Works_DB" -ForegroundColor Red}
if ([int]$current_value.Proc -ge [int]$Proc){Write-Host "CPU-PASSED : "$current_value.Proc "Cores available |",$Proc,"cores required for",$Encounters,"encounters per year" -ForegroundColor Green}else{Write-Host "CPU-FAILED : "$current_value.Proc "Cores available |",$Proc ,"cores required for",$Encounters,"encounters per year" -ForegroundColor Red}
if ([string]$global:SQL_reslts.SQL_Valuation -eq "You are on a supported SQL platform: SQL Server 2019 CU5+ or SQL Server 2022 CU11+"){Write-Host "SQL_PASSED : ",$global:SQL_reslts.SQL_Valuation -ForegroundColor Green}else{Write-Host "SQL-FAILED :",$global:SQL_reslts.SQL_Valuation -ForegroundColor Red}; 
Write-Host "Total encounters : ",$global:SQL_reslts.Encounters
Write-Host "Total appoinments : ",$global:SQL_reslts.Appointments
Write-Host "Total Active Users : ",$global:SQL_reslts.Active_users
Write-Host "Works_DB_size : ",$Works_size "TB"
}


<#**************************** Excel_Module_Installation ********************************#>

Install-Module -Name ImportExcel -RequiredVersion 7.8.4


<#*********************************** Non_DB_Export *************************************#>

function Non_DB_export{

param($SES, $Current_value)

$Global:myData = @( 
    [PSCustomObject]@{ Server_Name = $server.Split(",")[0]; Type = $server.Split(",")[1]; OS_Requirement = "Windows Server 2019 or 2022 (Standanrd, Datacenter or Enterprise)";
    OS_Available = $Current_value.OS; OS_Valuation = $os_valuation = if($Current_value.OS -notin $SES.OS){"FAILED"}else{"PASSED"}; 
    RAM_Requirement_GB = $SES.RAM; RAM_Available_GB = $Current_value.RAM; RAM_Valuation = $RAM_valuation = if([int]$Current_value.RAM -ge [int]$SES.RAM){"PASSED"}else{"FAILED"};
    CPU_requirement = $SES.Proc; CPU_Available = $Current_value.Proc; CPU_Valuation = if([int]$Current_value.Proc -ge [int]$SES.Proc){"PASSED"}else{"FAILED"};
    Total_space_C_drive_GB = [int]$Current_value.Total_space; Free_space_C_drive_GB = [int]$Current_value.Free_space; Drive_Valuation = if([int]$Current_value.Free_space -ge 20){"PASSED"}else{"FAILED"};

    }
)

$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK

$Global:myData | Export-Excel -Path ".\Output.xlsx" -AutoSize -Append -ConditionalText $positive_condition,$Negative_condition -BoldTopRow -FreezeTopRow -WorksheetName "Analysis"

$Summary = ".\Summary.ps1"
."$Summary"
Summary_export_NonDB
}


<#*********************************** DB_Export *************************************#>

function DB_export{

param($OS, $Current_value)

$Global:SQL_data = @( 
    [PSCustomObject]@{ Server_Name = $server.Split(",")[0]; Type = $server.Split(",")[1]; OS_Requirement = "Windows Server 2019 or 2022 (Standanrd, Datacenter or Enterprise)";
    OS_Available = $Current_value.OS; OS_Valuation = if($Current_value.OS -notin $OS.OS){"FAILED"}else{"PASSED"};
    RAM_Requirement_GB = $RAM_Required; RAM_Available_GB = $Current_value.RAM; RAM_Valuation = $RAM_valuation = if([int]$Current_value.RAM -ge [int]$RAM_Required){"PASSED"}else{"FAILED"};
    CPU_requirement = $Cores; CPU_Available = $Current_value.Proc; CPU_Valuation = if([int]$Current_value.Proc -ge [int]$Cores){"PASSED"}else{"FAILED"};
    Total_space_C_drive_GB = $Current_value.Total_space; Free_space_C_drive_GB = $Current_value.Free_space; Drive_Valuation = if($Current_value.Free_space -ge 20){"PASSED"}else{"FAILED"};
    SQL_version_Requirements = "Microsoft SQL Server 2019 (CU5) or Above"; SQL_Version_Available = $global:SQL_reslts.SQLVersion; SQL_Valuation = if($global:SQL_reslts.SQL_Valuation -eq "You are on a supported SQL platform: SQL Server 2019 CU5+ or SQL Server 2022 CU11+"){"PASSED"}else{"FAILED"};
    Total_Encounters = $Global:SQL_reslts.Encounters; Total_Appointments = $global:SQL_reslts.Appointments;
    Total_Active_users = $global:SQL_reslts.Active_users; 'Works_DB_size(TB)' = $Works_size

    }

)

$positive_condition = New-ConditionalText PASSED -BackgroundColor GREEN -ConditionalTextColor BLACK
$Negative_condition = New-ConditionalText FAILED -BackgroundColor RED -ConditionalTextColor BLACK

$SQL_data | Export-Excel -Path ".\Output.xlsx" -AutoSize -Append -ConditionalText $positive_condition,$Negative_condition -BoldTopRow -FreezeTopRow -WorksheetName "Analysis"

$Summary = ".\Summary.ps1"
."$Summary"
Summary_export_DB
}


############################# Data Extraction ##############################################################


Write-Host "Initialising SCP valuation ..." -ForegroundColor Yellow
$Servers = Get-Content -Path ".\Servers.txt"
Remove-Item -Path ".\Output.xlsx" -ErrorAction SilentlyContinue
Copy-Item -Path ".\.Temp\Output.xlsx" -Destination . -Force

Foreach($Server in $Servers){

if ($Server.Split(",")[1] -eq "WEB" -or $Server.Split(",")[1] -eq "AIO"){
$Web_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}
$Web_count += 1

#Print_console
display -current_value $Web_Server_currentvalue -actual_value $Web

#Non_DB_Export
Non_DB_export -SES $Web -Current_value $Web_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "MSG"){
$Msg_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $Msg_Server_currentvalue -actual_value $Msg

#Non_DB_Export
Non_DB_export -SES $Msg -Current_value $Msg_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "Fax"){
$Fax_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $Fax_Server_currentvalue -actual_value $Fax

#Non_DB_Export
Non_DB_export -SES $Fax -Current_value $Fax_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "PRT"){
$Prt_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $Prt_Server_currentvalue -actual_value $Prt

#Non_DB_Export
Non_DB_export -SES $Prt -Current_value $Prt_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "ACDM"){
$ACDM_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $ACDM_Server_currentvalue -actual_value $Acdm

#Non_DB_Export
Non_DB_export -SES $Acdm -Current_value $ACDM_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "SCAN"){
$Scan_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $Scan_Server_currentvalue -actual_value $SCAN

#Non_DB_Export
Non_DB_export -SES $Scan -Current_value $Scan_Server_currentvalue

}
elseif ($Server.Split(",")[1] -eq "DB"){

DB_connection

$DB_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

<#************************** DB RAM_requirement **********************************#>

$Works_size = $Global:SQL_reslts.Works_size / 1048576
if($Works_size -le 1.0 ){$RAM_Required = 64}elseif($Works_size -le 1.5){$RAM_Required = 96}elseif($Works_size -le 2.0){$RAM_Required = 128}elseif($Works_size -le 2.5){$RAM_Required = 160}elseif($Works_size -le 3.0){$RAM_Required = 192}elseif($Works_size -le 3.5){$RAM_Required = 224
}elseif($Works_size -le 4.0){$RAM_Required = 256}elseif($Works_size -le 4.5){$RAM_Required = 288}elseif($Works_size -le 5.0){$RAM_Required = 320}else{$RAM_Required = $Works_size * 320/5}

<#************************** DB processor_requiment ******************************#>

$Encounters = $Global:SQL_reslts.Encounters
if($Encounters -le 500000 ){$Cores = 4}
elseif($Encounters -le 1000000 ){$Cores = 6}
elseif($Encounters -le 2000000 ){$Cores = 8}
elseif($Encounters -le 3000000 ){$Cores = 10}
elseif($Encounters -le 4000000 ){$Cores = 12}
elseif($Encounters -le 5000000 ){$Cores = 16}
elseif($Encounters -le 6000000 ){$Cores = 20}
elseif($Encounters -le 7000000 ){$Cores = 24}
elseif($Encounters -le 8000000 ){$Cores = 28}
elseif($Encounters -le 9000000 ){$Cores = 30}
elseif($Encounters -le 10000000 ){$Cores = 34}
elseif($Encounters -le 11000000 ){$Cores = 38}
elseif($Encounters -le 12000000 ){$Cores = 40}
elseif($Encounters -le 13000000 ){$Cores = 44}
elseif($Encounters -le 14000000 ){$Cores = 48}
elseif($Encounters -le 15000000 ){$Cores = 50}
elseif($Encounters -le 16000000 ){$Cores = 54}
elseif($Encounters -le 17000000 ){$Cores = 58}
elseif($Encounters -le 18000000 ){$Cores = 60}
elseif($Encounters -le 19000000 ){$Cores = 64}
elseif($Encounters -le 20000000 ){$Cores = 68}
elseif($Encounters -le 21000000 ){$Cores = 70}
elseif($Encounters -le 22000000 ){$Cores = 74}
elseif($Encounters -le 23000000 ){$Cores = 78}
elseif($Encounters -le 24000000 ){$Cores = 80}
else{$Cores = $Encounters * 80 / 24000000}

$Appointments = $Global:SQL_reslts.Appointments
if($Appointments -le 80000 ){$Appointment_Cores = 4}
elseif($Appointments -le 170000 ){$Appointment_Cores = 6}
elseif($Appointments -le 330000 ){$Appointment_Cores = 8}
elseif($Appointments -le 500000 ){$Appointment_Cores = 10}
elseif($Appointments -le 670000 ){$Appointment_Cores = 12}
elseif($Appointments -le 830000 ){$Appointment_Cores = 16}
elseif($Appointments -le 1000000 ){$Appointment_Cores = 20}
elseif($Appointments -le 1170000 ){$Appointment_Cores = 24}
elseif($Appointments -le 1330000 ){$Appointment_Cores = 28}
elseif($Appointments -le 1500000 ){$Appointment_Cores = 30}
elseif($Appointments -le 1670000 ){$Appointment_Cores = 34}
elseif($Appointments -le 1830000 ){$Appointment_Cores = 38}
elseif($Appointments -le 2000000 ){$Appointment_Cores = 40}
elseif($Appointments -le 2170000 ){$Appointment_Cores = 44}
elseif($Appointments -le 2330000 ){$Appointment_Cores = 48}
elseif($Appointments -le 2500000 ){$Appointment_Cores = 50}
elseif($Appointments -le 2670000 ){$Appointment_Cores = 54}
elseif($Appointments -le 2830000 ){$Appointment_Cores = 58}
elseif($Appointments -le 3000000 ){$Appointment_Cores = 60}
elseif($Appointments -le 3170000 ){$Appointment_Cores = 64}
elseif($Appointments -le 3330000 ){$Appointment_Cores = 68}
elseif($Appointments -le 3500000 ){$Appointment_Cores = 70}
elseif($Appointments -le 3670000 ){$Appointment_Cores = 74}
elseif($Appointments -le 3830000 ){$Appointment_Cores = 78}
elseif($Appointments -le 4000000 ){$Appointment_Cores = 80}
else{$Appointment_Cores = $Appointments * 80 / 4000000}

if($Appointment_Cores -gt $Cores){$Cores = $Appointment_Cores}


Sql_display -current_value $DB_Server_currentvalue -OS $DB -RAM $RAM_Required -Proc $Cores


<#************************** DB_Export_call **************************************#>

DB_export -OS $DB -Current_value $DB_Server_currentvalue

}
elseif($Server.Split(",")[1] -eq "UNITY"){
$Unity_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $Unity_Server_currentvalue -actual_value $Unity

#Non_DB_Export
Non_DB_export -SES $Unity -Current_value $Unity_Server_currentvalue

}

elseif($Server.Split(",")[1] -eq "FHIR"){
$FHIR_Server_currentvalue = Invoke-Command -ComputerName $Server.Split(",")[0] -ScriptBlock{ @{ OS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum / 1024 / 1024 / 1024
Proc = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment").NUMBER_OF_PROCESSORS
IPAddress = (Get-NetIPAddress -AddressFamily IPv4).IPAddress[0]
Total_space = (Get-Volume C).Size/1gb
Free_space = (Get-Volume C).SizeRemaining/1gb}}

#Print_console
display -current_value $FHIR_Server_currentvalue -actual_value $Unity

#Non_DB_Export
Non_DB_export -SES $Unity -Current_value $FHIR_Server_currentvalue

}

else{Write-Host "`nThis is another server"
continue}
}

$WebServers_requirement = [Math]::Ceiling((($global:SQL_reslts.Active_users)/175))

if($Web_count -ge ($global:SQL_reslts.Active_users)/175){Write-Host "`nWeb servers available : " $Web_count "| Required : "$Web_count -ForegroundColor Green}
elseif($Web_count -ge $WebServers_requirement){
Write-Host "`nWeb servers available : "$Web_count" | Required :",$WebServers_requirement -ForegroundColor Green
}else{
Write-Host "`nWeb servers available : "$Web_count" | Required :", $WebServers_requirement -ForegroundColor Red
}

try {
    $Message_Requirement = [int]$Global:SQL_reslts.Jobcount / 5000
    if ($Global:SQL_reslts.Total_Spooler_Count -ge $Message_Requirement) {
        Write-Host "Sufficient Message\Print Servers available" -ForegroundColor Green
        $Total_Spoolers = $Global:SQL_reslts.Total_Spooler_Count
    }
    else {
        Write-Host "Insufficient Message\Print Servers available : $Total_Spoolers' | Require : $Message_Requirement"
    }
}
catch {

}


########################## Pycode Execution #######################################
Start-Process -FilePath ".\Py_binaries\SCP_Master_Formating.exe"
echo ""
powershell -noexit