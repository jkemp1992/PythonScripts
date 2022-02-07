# Script: Status Report Email Generation

#Checks to ensure port to Mail-Relay is open
#New-Object System.Net.Sockets.TcpClient()

#Sets Mail Server Address
$PSEmailServer = ""

#Variables
$from = ""
$recipients = ""
[string[]]$to = $recipients.Split(',')
#$ccrecipients = ""
#[string[]]$cc = $ccrecipients.Split(',')
$Bccrecipients = ""
[string[]]$Bcc = $Bccrecipients.Split(',')
$Error_To_Peeps = ""
[string[]]$Error_To = $Error_To_Peeps.Split(',')

$file = "C:\Scripts\Automated_Status_Report_From_Excel_and_SQL.xlsm"

Function SendErrorEmail([String]$FailedItem, [String]$ErrorMessage)
{
    Send-MailMessage -From blah -To $Error_To -Subject "Status Report Error" -Body "There was an error running report $FailedItem. The error message was $ErrorMessage" -SmtpServer $PSEmailServer -Port "25"
}


Try{
#Create and get my Excel Obj
$x1 = New-Object -comobject Excel.Application
$WorkBook = $x1.Workbooks.Open($file)
$x1.Visible = $true

#Run Macro
Start-Sleep -Seconds 5
$app = $x1.Application
$app.Run("export_data")

#Select Sheet and Save data to list
$sheetname = "ExcelSheet"
$Worksheet = $WorkBook.Sheets.Item($sheetname)

# Create arrays to loop through Excel sheets with
$_Array = @()
$rowESD, $colESD = 3,2

#Get number of rows in table
$row = $Worksheet.Range("P1").Value2
$rowMax = [int]$row

#Add companies to Array
for ($i=0; $i -ne $rowMax; $i++)
{
    $company = $Worksheet.Cells.Item($rowESD+$i,$colESD).value2
    $ESD_Array += $company
}


#Select Sheet
$sheetname = "Pivot"
$Worksheet = $WorkBook.Sheets.Item($sheetname)

#Get number of rows in table
$row = $Worksheet.Range("Z1").Value2
$rowMax = [int]$row


#Starting position for each column
$rowB,$colB = 6,2
$rowM,$colM = 6,13
$rowE,$colE = 6,5
$rowF,$colF = 6,6
$rowG,$colG = 6,7
$rowH,$colH = 6,8

#Get table headers
$header1 = $Worksheet.Range("B5").Value2
$header2 = $Worksheet.Range("M5").Value2
$header3 = $Worksheet.Range("E5").Value2
$header4 = $Worksheet.Range("F5").Value2
$header5 = $Worksheet.Range("G5").Value2
$header6 = $Worksheet.Range("H5").Value2

#Construct table using HTML in Body
$body = @"
	<b>Daytime Status Report</b>
	<br><br>
	
	<style>
	table, th, td {
  				border: 1px solid black;
				border-collapse: collapse;
	}
	th {
				color: white;
				background-color: blue;
	}
	</style>
	<table>
		<tr>
			<th align="left">$header1</th>
			<th align="left">$header2</th>
			<th align="left">$header3</th>
			<th align="left">$header4</th>
			<th align="left">$header5</th>
			<th align="left">$header6</th>
		</tr>
"@



#Iterate over each row in Excel and save variables
for ($x=0; $x -ne $ESD_Array.Count; $x++)
{
    $ESD_Company = $ESD_Array[$x]
    $count = 0

    for ($i=0; $i -ne $rowMax; $i++)
    {
        $company = $Worksheet.Cells.Item($rowB+$i,$colB).value2

        If($ESD_Company -eq $company)
        {

            $issue = $Worksheet.Cells.Item($rowM+$i,$colM).value2
            $ticket = $Worksheet.Cells.Item($rowE+$i,$colE).value2
            $team = $Worksheet.Cells.Item($rowF+$i,$colF).value2
            $lastQueue = $Worksheet.Cells.Item($rowG+$i,$colG).value2
            $country = $Worksheet.Cells.Item($rowH+$i,$colH).value2
            $count += 1
        }
        
        If(($ESD_Company -ne $company) -and ($i -eq $rowMax-1) -and ($count -eq 0))
        {
            $issue = "(No new issues)"
            $ticket = ""
            $team = ""
            $lastQueue = ""
            $country = ""
        }
        

        If(
                (($i -eq $rowMax-1) -and ($count -eq 0)) -or 
                (($ESD_Company -eq $company) -and ($count -eq 1))
          )
        {

        #Add a row to the HTML table
        $body = $body + @"
		        <tr>
              <td valign="top"><b>$ESD_Company</b></td>
			        <td valign="top">$issue</td>
			        <td valign="top">$ticket</td>
				      <td valign="top">$team</td>
			        <td valign="top">$lastQueue</td>
			        <td valign="top">$country</td>
		        </tr>
"@
        }

        If(($ESD_Company -eq $company) -and ($count -gt 1))
        {
        #Add a row to the HTML table
        $body = $body + @"
		        <tr>
              <td valign="top"></td>
			        <td valign="top">$issue</td>
			        <td valign="top">$ticket</td>
				      <td valign="top">$team</td>
			        <td valign="top">$lastQueue</td>
			        <td valign="top">$country</td>
		        </tr>
"@
        }
    }    
}


#Finish the HTML table
    $body = $body + @"
		    </table>
"@



#Send Email with timeout error catch
$Exit = 0
	Do{
		Try{
            Send-MailMessage -From $from -To $to -Bcc $Bcc -Subject "Daytime Status Report" -SmtpServer $PSEmailServer -Port "25" -BodyAsHtml $body
            $Exit = 4
		}
		Catch{
			$Exit++
			$ErrorMessage = "Unable to send email because: $($Error[0]) Attempt # $Exit"
			Write-Verbose "$ErrorMessage"
			If($Exit -eq 4)
			{
				$FailedItem = "Mail Relay"
				$ErrorMessage = "Unable to send email because: $($Error[0]) - attempted 4 times"
				Write-Verbose "$ErrorMessage"
				SendErrorEmail $FailedItem $ErrorMessage
			}
		}
	}Until($Exit -eq 4)
}

Catch{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    SendErrorEmail $FailedItem $ErrorMessage
}
Finally{
    
}


$sheetname = "Pivot_Non-ESD"

Try{
#Select Sheet
$Worksheet = $WorkBook.Sheets.Item($sheetname)

#Get number of rows in table
$row = $Worksheet.Range("Z1").Value2
$rowMax = [int]$row

#Array of time slots, Edit as needed
$B = "8:30AM","9:30AM","10:30AM","11:30AM","12:30PM","1:30PM","2:30PM","3:30PM","4:30PM","5:30PM","6:30PM","7:30PM","8:30PM","9:30PM"

#Starting position for each column
$rowN,$colN = 6,14
$rowM,$colM = 6,13
$rowE,$colE = 6,5
$rowF,$colF = 6,6
$rowG,$colG = 6,7
$rowH,$colH = 6,8
$rowI,$colI = 6,9

#Get table headers
$header1 = $Worksheet.Range("B5").Value2
$header2 = $Worksheet.Range("M5").Value2
$header3 = $Worksheet.Range("E5").Value2
$header4 = $Worksheet.Range("F5").Value2
$header5 = $Worksheet.Range("G5").Value2
$header6 = $Worksheet.Range("H5").Value2
$header7 = $Worksheet.Range("I5").Value2

#Construct table using HTML in Body
$body = @"
	<b>Daytime Status Report</b>
	<br><br>
	
	<style>
	table, th, td {
  				border: 1px solid black;
				border-collapse: collapse;
	}
	th {
				color: white;
				background-color: blue;
	}
    p {
	            border: 1px solid black;
                color: white;
                background-color: blue;
                border-collapse: collapse;
    }
	</style>
	<table>
		<tr>
			<th align="left">$header1</th>
			<th align="left">$header2</th>
			<th align="left">$header3</th>
			<th align="left">$header4</th>
			<th align="left">$header5</th>
			<th align="left">$header6</th>
      <th align="left">$header7</th>
		</tr>
"@

$count = 0
$start = 0

#Iterate over each row in Excel and save variables
for ($x=0; $x -ne 13; $x++)
{
    $time = $B[$x]

    for ($i=$start; $i -ne $rowMax; $i++)
    {
        $time_Excel = $Worksheet.Cells.Item($rowN+$i,$colN).value2

        If($time -eq $time_Excel)
        {
            $issue = $Worksheet.Cells.Item($rowM+$i,$colM).value2
            $ticket = $Worksheet.Cells.Item($rowE+$i,$colE).value2
            $company = $Worksheet.Cells.Item($rowF+$i,$colF).value2
            $team = $Worksheet.Cells.Item($rowG+$i,$colG).value2
            $lastQueue = $Worksheet.Cells.Item($rowH+$i,$colH).value2
            $country = $Worksheet.Cells.Item($rowI+$i,$colI).value2
            $count = $count + 1
        }
        ElseIf($i -eq $start)
        {
            $issue = "(No new issues)"
            $ticket = ""
            $company = ""
            $team = ""
            $lastQueue = ""
            $country = ""
        }
        

        If($i -eq $start)
        {

        #Add a row to the HTML table
        $body = $body + @"
		        <tr>
              <td valign="top"><b>$time</b></td>
			        <td valign="top">$issue</td>
			        <td valign="top">$ticket</td>
				      <td valign="top">$company</td>
			        <td valign="top">$team</td>
			        <td valign="top">$lastQueue</td>
			        <td valign="top">$country</td>
		        </tr>
"@
        }
        Elseif($time -eq $time_Excel)
        {
        #Add a row to the HTML table
        $body = $body + @"
		        <tr>
              <td valign="top"></td>
			        <td valign="top">$issue</td>
			        <td valign="top">$ticket</td>
				      <td valign="top">$company</td>
			        <td valign="top">$team</td>
			        <td valign="top">$lastQueue</td>
			        <td valign="top">$country</td>
		        </tr>
"@
        }

    }
    $start = $count
}

#Finish the HTML table
    $body = $body + @"
		    </table>
"@




#Send Email with timeout error catch
$Exit = 0
	Do{
		Try{
            Send-MailMessage -From $from -To $to -Bcc $Bcc -Subject "Status Report" -SmtpServer $PSEmailServer -Port "25" -BodyAsHtml $body
            $Exit = 4
		}
		Catch{
			$Exit++
			$ErrorMessage = "Unable to send email because: $($Error[0]) Attempt # $Exit"
			If($Exit -eq 4)
			{
				$FailedItem = "Mail Relay"
				$ErrorMessage = "Unable to send email because: $($Error[0]) - attempted 4 times"
				SendErrorEmail $FailedItem $ErrorMessage
			}
		}
	}Until($Exit -eq 4)
}

Catch{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    SendErrorEmail $FailedItem $ErrorMessage
}
Finally{
    
}

#Close Workbook and Excel
$WorkBook.close($true)
$x1.Quit()

