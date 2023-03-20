function Export-Excel {
<#
.SYNOPSIS
Module for out PowersSell Object to Excel table
.DESCRIPTION
Example:
$Service = Get-Service
Export-Excel $Service -Path "$home\Desktop\out.xlsx" # path default
.LINK
https://github.com/Lifailon
#>
Param (
[Parameter(Mandatory = $True)] $Object,
$Path = "$home\Desktop\out.xlsx"
)
$TempCSV = "$env:TEMP\ConvertTo-Excel.csv"
$Delimiter=","
$Object | Export-Csv $TempCSV -Append -Encoding Default -Delimiter $Delimiter
$temp = (cat $TempCSV)[1..10000]
$temp > $TempCSV
$Excel = New-Object -ComObject excel.application
$WorkBook = $Excel.WorkBooks.Add(1)
$WorkSheet = $WorkBook.WorkSheets.Item(1)
$TxtConnector = ("TEXT;" + $TempCSV)
$Connector = $WorkSheet.QueryTables.add($TxtConnector,$WorkSheet.Range("A1"))
$QueryTables = $WorkSheet.QueryTables.item($Connector.name)
$QueryTables.TextFileOtherDelimiter = $Delimiter
$QueryTables.Refresh()
$QueryTables.Delete()
$WorkBook.SaveAs($Path)
$Excel.Quit()
Remove-Item $TempCSV
}

$CredFile = ".\Cred-Email.xml"
try {
$Cred = Import-Clixml -path $credFile
}
catch {
$Cred = Get-Credential -Message "Enter credential"
if ($Cred -ne $null) {
$Cred | Export-CliXml -Path $credFile
}
else {return}
}

### Var
$emailSenderAddr = $Cred.UserName
$emailTo = "login@domain.ru"
$emailSmtpServer = "mail.domain.ru"
$TriggerDays = 7
$PathLog = "C:\Veeam-Job-Stat-Log"
###

$jobs = @()
$stat = Veeam-Job-Stat
$date = Get-Date
foreach ($vstat in $stat) {
[DateTime]$vdate = $vstat.TimeLastCompletion
[int32]$days=($date-$vdate).Days
if ($days -gt $TriggerDays) {
#Write-Host Job: $vstat.JobName Time Last Completion $days days ago -ForegroundColor Red
$jobs += $vstat.JobName
} elseif ($days -le $TriggerDays) {
#Write-Host Job: $vstat.JobName Time Last Completion $days days ago -ForegroundColor Green
}
}
$Count = $jobs.Count
$CountAll = $stat.Count
$JobsJoin = $jobs -join "; "
$Message = "$Count out of $CountAll jobs have not been performed for more than $TriggerDays days:
$JobsJoin"

if (!(Test-Path $PathLog)) {
New-Item -Path $PathLog -ItemType "Directory" -Force
}
$date_out = Get-Date -f "dd/MM/yyyy"
Export-Excel -object $stat -path "$PathLog\Veeam-Job-Stat-$date_out.xlsx"
sleep 1

Send-MailMessage -From $emailSenderAddr -To $emailTo -Subject "Veeam Jobs Statisctics" `
-Body $Message –SmtpServer $emailSmtpServer -Encoding "UTF8" -Credential $Cred `
-Attachment "$PathLog\Veeam-Job-Stat-$date_out.xlsx"