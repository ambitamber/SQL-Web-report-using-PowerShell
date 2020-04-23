#clearing the errors
$Error.Clear()

<#This parameter will use for table when coverting the table into html.#>
$table = @{
    head = @" 
<style> 
    table {font-family: verdana, arial, sans-serif;font-size: 11px;color: #333333;border-width: 1px;border-color: #3A3A3A;border-collapse: collapse;}
    th{border-width: 1px;padding: 8px;border-style: solid;border-color: #517994;background-color: #B2CFD8;}
    tr:hover td{background-color: #DFEBF1;}
    td{border-width: 1px;padding: 8px;border-style: solid;border-color: #517994;background-color: #ffffff;}
</style>
"@}

<# Global parameter for select. This will use when selecting columns from $auditreportdata #>
    $select = "Loadstart","datafile","RecordsRead","ErrorsGenerated","errorfile"

<#Global parementer for email. To add more people, just add coma and quotes. Then add email address inside of quotes.
I have not declear subject or body. This will happen under each facility. #>
    $emailSMTPServer = ""
    $emailRecipients = ""
    $emailcc = ""
    $emailFrom = ""

#Server Parameter
$servername = ""<#Enter the name of the SQL Server#>
$datebasename = ""<#Enter the name of the Database Name#>
$datebasename1 = ""<#Enter the name of the Database Name for second query#>


<#This is a do loop. It will keep asking user for input until it matches the parameter.
These are the parameters: Can only be a number and it needs to be higher than 0. 
when user inputs a number it passes to date. #>
do {
    $num = read-host "`nHow many days do you want to go back?"
    $a = ""
    if([int32]::TryParse( $num , [ref]$a) -and ($a -ge 1)){
        <#This will declear the date and minus days that user enters. Then it will format the date to 'yyyyMMdd' and store in $date #>
        $gettingdate = (get-date).AddDays(-$a)
        $date = Get-Date $gettingdate -Format MM/dd/yyyy
        $datehr = $date +" 12:00"
    }
    else {
        write-host "`nOnly integers and greater than 0 please.`n"
    }
} 
until ( [int32]::TryParse( $num , [ref]$a ) -and ($a -ge 1))

<#This is Sql connection which connects to phssql2139 grabs information from T_LOAD_EXECUTION_LOG and stores into $auditreportdata.
    ------------------------------  First SQL connection begins -------------------------------- 
    The try function will try to connect to Server. If it is successfull, it will go into finally function to close the connection and continue to next step. if it's fail to connect, it go into catch function
    and catches any errors. It will send a email with error.#>
try {
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$servername;Database=$datebasename;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    #Sql qurey to get data from Load execution log table using where the LoadStartDate greater then $date. 
    $SqlCmd.CommandText = $( " Enter the SQL Query Here ")
    $SqlCmd.Connection = $SqlConnection
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $reportDataSet = New-Object System.Data.DataSet
    $SqlAdapter.fill($reportDataSet)
}
catch [System.Management.Automation.MethodInvocationException]{
    $errormes = "There was an error connecting to $servername, and to database $datebasename. here is the error: $error[0]"
    Write-Host $errormes 
    Send-MailMessage -To $emailRecipients -From $emailFrom -Cc $emailcc -Subject "Error running Audit Report Script" -SmtpServer $emailSMTPServer -Body $errormes 
}
finally{
    $SqlConnection.Close()
    $auditreportdata = $reportDataSet.Tables[0]
}
<#---------------------------  First SQL connection Ends ---------------------------- #>

<#This is Sql connection which connects to phssql2139 grabs information from TaskLogMessages and TaskEngineLogMessages under PMDataSutdio and stores into $auditmessagedata
------------------------  Second SQL connection begins ----------------------------  
The try function will try to connect to Server. If it is successfull, it will go into finally function to close the connection and continue to next step. if it's fail to connect, it go into catch function
    and catches any errors. It will send a email with error.#>
try {
    $SqlConnection.ConnectionString = "Server=$servername;Database=$datebasename1;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $( " Enter the SQL Query Here ")
    $SqlCmd.Connection = $SqlConnection
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $reportDataSet = New-Object System.Data.DataSet
    $SqlAdapter.fill($reportDataSet)
}
catch [System.Management.Automation.MethodInvocationException]{
    $errormes = "There was an error connecting to Server $servername, and to database $datebasename1. here is the error: $error[0]"
    Write-Host $errormes 
    Send-MailMessage -To $emailRecipients -From $emailFrom -Cc $emailcc -Subject "Error running 'script name' Script" -SmtpServer $emailSMTPServer -Body $errormes 
}
finally{
    $SqlConnection.Close()
    $auditmessagedata = $reportdataset.Tables[0]
}
<#------------------------  Second SQL connection Ends ------------------------------ #>


#This will tell the user how many day the script going back to pull the data when they input a value in do loop
write-host ""`n"Going back $num day(s) to the following date: $datehr`n"

 
<# After the SQL connection, we use Facilty_Code column in $auditreportdata where equal 1200 to get data for MGH. Also using $select to select columns that we want. Then convert the table into html for email body.#>
    $report = $auditreportdata | Where-Object {$_.Facilty_Code -eq "Please enter a value to be equals to "} | Select-Object $select| ConvertTo-Html @table
    <# After the SQL connection, we use Facilty_Code column in $auditmessagedata where equal 1200 to get data for MGH. Then convert the table into CSV file. Then export the file to the location.
    The reason that we export the file becasue of the layout. As you see secound line, we use import function to import the file back into the script and convert into html.#>
    $messages = $auditmessagedata | Where-Object {$_.Facilty_Code -eq "Please enter a value to be equals to"} | Select-Object "Message"| ConvertTo-Html @table
    $Subject = ""
    $bodyreport = "<p></p>" 
    $bodymessage = "<p><p>" 
    $Body = "$bodyreport" + "$report" + "$bodymessage" + "$messages"
    <#The IF statement will check the file and count rows. If there is more then 0 rows, it will send a email using the paramenter. Else statement will let the user know that there is no data for the facility. #>
    if(($auditreportdata).count -ge 1){ Write-Host "`nFound data. Sending email."
    Send-MailMessage -From $emailFrom -To $emailRecipients -SmtpServer $emailSMTPServer -subject $Subject -Body $Body -BodyAsHtml -Cc $emailcc } else {Write-Output "`nChecked, there is no data."}

        
if($Error.Count -ge 1){
    Write-Output "There is error running this script. This is the error: $error"
}
else {
    #this will tell the user if the script has ran without no issue. 
    Write-Output "`nThe scipt ran without an error. The process has ended.`n"
}

PAUSE