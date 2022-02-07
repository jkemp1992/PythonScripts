[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
$jsonserial = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
$jsonserial.MaxJsonLength = [int]::MaxValue

# Script loops through a given date range, pulling data from Service Now one day at a time
# Pulls data from Service Now open APIs (public knowledge)
# https://docs.servicenow.com/bundle/quebec-application-development/page/integrate/inbound-rest/concept/c_RESTAPI.html

# Get date parameter
$t = (get-date).AddDays(1)

# Get each date part seperately
$y = [String]$t.Year
$m = [String]$t.Month
$d = [String]$t.Day

# If the preceding 0's are not there, add them
if ($m.Length -eq 1) {$m = "0" + $m}
if ($d.Length -eq 1) {$d = "0" + $d}

# Construct the final string variable for API call
[string]$maxDate = $y + "-" + $m + "-" + $d + " 00:00:00"

$DateStart = "2022-01-20 00:00:00"
$DateEnd = "2022-01-21 00:00:00"

# Column Fields
$fields = "all, fields, to, include"


DO {

    Write-Host "Started loop!"
    Write-Host $DateStart " - " $DateEnd

    # ------------------------Service Now URLs-----------------------

    # Date Range
    $url = "https://myInstance.service-now.com/api/now/table/my_table?sysparm_limit=5000000&sysparm_fields=" + $fields + "&sysparm_exclude_reference_link=true&sysparm_display_value=true&sysparm_query=sys_created_onBETWEEN" + $DateStart + "@" + $DateEnd


    $username = "username"
    $password = "password"

    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $username, $password)))

    # Set proper headers
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
    #$headers.Add('Accept','application/json')
    #$headers.Add('Content-Type','application/json')

    # Specify HTTP method
    $method = "Get"

    # Send HTTP request
    $response = Invoke-WebRequest -Headers $headers -Method $method -Uri $url -ContentType "application/json" -UseBasicParsing

    # Convert the json to a PSObject
    #$tickets = ($response | ConvertFrom-Json).result
    $content = $response.Content
    $data = $jsonserial.DeserializeObject($content)
    $results = $data.result

    # Save to File
    ($response | ConvertFrom-Json).result | Export-csv -delimiter ',' -Path .\myFile.csv -NoTypeInformation

    # SQL Connection Properties
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=My-Server;Database=MyDB;Integrated Security=True"

    $SQLCmdString = "dbo.spi_SQL_Stored_Procedure"
    $SQLCmdTimeout = 0

    write-host ("Record Count: " + $results.Count)
    write-host " "

    # Loop through each ticket in object, save variables, and upload to SQL
    For ($i=0; $i -le $results.Count-1; $i++){
        $details = $results.Item($i)

        $sys_created_on = $details.sys_created_on
        $u_host  = $details.u_host
        $u_service_check  = $details."u_service_check.u_display_name"
        # .... continued

        # Start SQL Connection
        $SQLCmd = New-Object System.Data.SqlClient.SqlCommand
        $SQLCmd.CommandType = [System.Data.CommandType]::StoredProcedure
        $SQLCmd.CommandText = $SQLCmdString
        $SQLCmd.CommandTimeout = $SQLCmdTimeout
        $SQLCmd.Connection = $SqlConnection

        # Add each Parameter
        $SQLCmd.Parameters.Add("@sys_created_on",[system.data.SqlDbType]::date) | out-Null
        $SQLCmd.Parameters['@sys_created_on'].Direction = [system.data.ParameterDirection]::Input
        $SQLCmd.Parameters['@sys_created_on'].value = $sys_created_on
        # .... continued

        try{
            # Try to connect to SQL and execute command
            $SqlConnection.Open() 
            $SQLCmd.ExecuteNonQuery() | out-null -Verbose
            $SQLConnection.Close()
        }
        catch{

            Read-Host "There was an error exception that was caught. Press any key to continue..."

        }

    }

    # Convert string variable to datetime
    $start = [datetime]::ParseExact($DateStart, "yyyy-MM-dd hh:mm:ss", $null)
    $end = [datetime]::ParseExact($DateEnd, "yyyy-MM-dd hh:mm:ss", $null)

    # Add 1 day to the datetime variables
    $start = [datetime]$start.AddDays(1)
    $end = [datetime]$end.AddDays(1)

    # Get each date part seperately
    $y = [String]$start.Year
    $m = [String]$start.Month
    $d = [String]$start.Day

    # If the preceding 0's are not there, add them
    if ($m.Length -eq 1) {$m = "0" + $m}
    if ($d.Length -eq 1) {$d = "0" + $d}

    # Construct the final string variable for API call
    [string]$DateStart = $y + "-" + $m + "-" + $d + " 00:00:00"

    # Get each date part seperately
    $y = [String]$end.Year
    $m = [String]$end.Month
    $d = [String]$end.Day

    # If the preceding 0's are not there, add them
    if ($m.Length -eq 1) {$m = "0" + $m}
    if ($d.Length -eq 1) {$d = "0" + $d}

    # Construct the final string variable for API call
    [string]$DateEnd = $y + "-" + $m + "-" + $d + " 00:00:00"

} While ($DateEnd -ne $maxDate)
