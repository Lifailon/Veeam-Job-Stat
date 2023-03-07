function ConvertFrom-XLSX {
Param (
[Parameter(Mandatory = $True, HelpMessage = "Source path to file csv")] $SourceCsv,
[Parameter(Mandatory = $True, HelpMessage = "Destination path to file xlsx")] $DestXlsx,
[Parameter(HelpMessage="Delimiter")] $Delimiter=","
)
$Excel = New-Object -ComObject excel.application
$WorkBook = $Excel.WorkBooks.Add(1)
$WorkSheet = $WorkBook.WorkSheets.Item(1)
$TxtConnector = ("TEXT;" + $SourceCsv)
$Connector = $WorkSheet.QueryTables.add($TxtConnector,$WorkSheet.Range("A1"))
$QueryTables = $WorkSheet.QueryTables.item($Connector.name)
$QueryTables.TextFileOtherDelimiter = $Delimiter
$QueryTables.Refresh()
$QueryTables.Delete()
$WorkBook.SaveAs($DestXlsx)
$Excel.Quit()
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

$emailSenderAddr = $Cred.UserName
$emailTo = "login@domain.ru"
$emailSmtpServer = "mail.domain.ru"

$TriggerDays = 7
$PathLog = "C:\Veeam-Job-Stat-Log"

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
$stat | Export-Csv "$PathLog\Veeam-Job-Stat.csv" -Append -Encoding Default
ConvertFrom-XLSX "$PathLog\Veeam-Job-Stat.csv" "$PathLog\Veeam-Job-Stat-$date_out.xlsx"
rm "$PathLog\Veeam-Job-Stat.csv"
#rm "$PathLog\Veeam-Job-Stat-$date_out.xlsx"
sleep 1

Send-MailMessage -From $emailSenderAddr -To $emailTo -Subject "Veeam Jobs Statisctics" `
-Body $Message –SmtpServer $emailSmtpServer -Encoding "UTF8" -Credential $Cred `
-Attachment "$PathLog\Veeam-Job-Stat-$date_out.xlsx"