<#
 open excel files, refresh odata link, save, close move to 'tobeimported' directory

 written by rh
 written 2/27/15
 
 Version0.9 3/5/15
   working script ready for production

    Future enhancements:
        formalize Powershell cmdlet

  revised 3/3/15
    added logging logic
    added move files
  3/4/15
    cleaned up logging  output
  3/20/15
    fixed excel row count
  8/3/15
	eBuilder added indexes to the tables so updates are much faster.  There is a limit of 10 pulls per 15 minutes and since
	the pulls are faster we are hitting the API limit.  start-sleep stalls added to the script to stay below the API limit.


Manual refresh of each file is very slow and cumbersome.  each excel file must be opened and refreshed to 
capture new data from eBuilder via the API.  There are performance and network issues if we try to refresh
more than 1 file at a time. so this process has to be "babysat".  Launch refresh and wait, save result and
launch next file.  Other work can be done but NOT in excel.  Most of the files are relatively fast, 2-5 minutes
to refresh.  But the current 4 files over 1MB and the invoice update take from 30 min to 2 hours for invoice
incremental pull.Invoice update has usually be run overnight.  I'll start when i leave for the day and save
the next morning.  Script allows that is run unattended.

Man hours required to complete this manually 1 hour for smaller files, 4 hours for larger remaining file.  
cut that in half since other work can be done (less productive due to network and CPU issues) 
SPS 120 minutes per script run.
5/15 eBuilder added indexes to the tables so download times have been greatly improved.  Total download time is now about an hour with the automated system or 2 hours if performed manually
SPS 60 minutes per script run.
#>

$logpath = "c:\rich\importlog.log"
$gd = get-date
$inputstringSTART =  "$gd : refresh script started"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringSTART -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 10

$connsource = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\sql source files\toberefreshed\"
$conntarget = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\sql source files\tobeimported\"
$d = get-date -displayhint date


$files = Get-ChildItem -path $connsource -Filter *xlsx |
    where-object {$_.LastWriteTime -le ($d.adddays(0))} |
    Select-object -ExpandProperty name

$filescount = ($files | Measure-Object)
$fc = $filescount.count

$gd = get-date
$inputstring =  "$gd : $fc files to be refreshed"
Add-Content -Path $logpath -Value $inputstring -Force
start-sleep -s 10

foreach($f in $files)
{

$a = New-Object -ComObject Excel.Application
$a.Visible = $True
start-sleep -s 10
$b = $a.workbooks.Open("$connsource$f");
start-sleep -s 10
#log refresh start time
$gd = get-date
$inputstring1 =  "$gd : $f refreshed from eB started"
Add-Content -Path $logpath -Value $inputstring1 -Force
start-sleep -s 50

$b.refreshall()
$a.CalculateUntilAsyncQueriesDone()
start-sleep -s 5
$a.activeworkbook.save()
start-sleep -s 50
$b.save()
start-sleep -s 5
$b.saveas("$connsource$f.csv",6)
#log refresh finish time
$gd = get-date
$inputstring2 =  "$gd : $f refreshed from eB finished"
Add-Content -Path $logpath -Value $inputstring2 -Force
start-sleep -s 5

$a.application.displayalerts = $false
$a.application.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($a)

$count = import-csv "$connsource$f.csv" | Measure-Object
$cc = $count.count

$f
$cc

$gd = get-date
$inputstring3 =  "$gd : $f contains $cc lines"
Add-Content -Path $logpath -Value $inputstring3 -Force
start-sleep -s 20
$count = $null
$cc = $null
$b = $null
$a = $null
}   

start-sleep -s 15

Get-ChildItem -path $connsource -Filter "*.csv" | Remove-Item

Get-ChildItem -Path $connsource | move-Item -Destination $conntarget

$gd = get-date
$inputstring4 =  "$gd : refreshed $fc files and moved to tobeimported directory"
Add-Content -Path $logpath -Value $inputstring4 -Force
start-sleep -s 5

$gd = get-date
$inputstringEND =  "$gd : refresh script complete"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringEND -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 5

