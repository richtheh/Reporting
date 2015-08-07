<#

 import script to import excel spreadsheet to our eBdatabase

 written by rh
 written 2/19/15

 Version 0.9  3/5/15
  working script ready for production

  future enhancement - formalize powershell formatting. commandbinding()
      -create function tool - 
      -building error handling
      -revise move function and move into foreach loop to better see process status
      -refine comment - will accomplish with commandbinding()
      -log how often script runs for successfactors

  this upload process is much faster than doing each action manually
  approx 5 min per file to get from email, save to shared drive, upload to access
  currently 22 file
  saving per script run 60-100 minutes
  SPS 80 minutes
#>


$acImport = 0 
$acSpreadsheetTypeExcel12 = 9 
$conndb = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\database\eBAPI_be.accdb"
$connsource = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\sql source files\tobeimported"
$connDone = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\sql source files\done\"
$connRefresh = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\sql source files\toberefreshed"

$d = get-date -displayhint date
$logpath = "c:\rich\importlog.log"

$gd = get-date
$inputstringSTART =  "$gd : import script started"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringSTART -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 5

#identify file and table name to be cleared and updated
$h = Get-ChildItem -path $connsource/*xlsx |
     Select-object -property name, lastwritetime |
     where-object {$_.LastWriteTime -ge ($d.adddays(-30))} | 
     Select-object -ExpandProperty name

$table = $h.replace(".xlsx","")

$hcount = ($h | Measure-Object)
$hc = $hcount.count

#start Access
$a = New-Object -Comobject Access.Application 
$a.OpenCurrentDatabase($conndb) 


ForEach ($t in $table)
{
    $sql = "DELETE * FROM $t;"  # needs to be in foreach loop since it is dependent on $t

    # deletes old data from tables
    $a.DoCmd.runsql($sql,$False)

    # imports new data from spreadhsheets
    $a.DoCmd.TransferSpreadsheet($acImport, $acSpreadsheetTypeExcel12, $t, "$connsource\$t.xlsx", $True) 

    # count rows added to access table and add to log fil
    $count = (get-content $connsource\$t.xlsx | Measure-Object)
    $cc = $count.count
    $gd = get-date
    $inputstring3 =  "$gd : spreadsheet $t.xlsx imported $cc lines to database table $t"
    Add-Content -Path $logpath -Value $inputstring3 -Force
}

$a.Quit()

#copy imported files to 'done' directory

Get-ChildItem "$connsource" | ForEach-Object {          
    copy-Item $_.FullName "$connDone$($_.BaseName -replace " ", "_" -replace '\.([^\.]+)$')-$(Get-Date -Format "MMddyyyy-HHmmss").xlsx"
}


#move files to 'toberefreshed' directory
Get-ChildItem -Path $connsource | Move-Item -Destination $connrefresh


$gd = get-date
$inputstring4 =  "$gd : imported $hc files and moved to done directory"
Add-Content -Path $logpath -Value $inputstring4 -Force
start-sleep -s 5




$gd = get-date
$inputstringEND =  "$gd : Import script END"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringEND -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 5
