<#

 Run Access update and select queries to generate report datasets

 written by rh
 written 3/10/15

 Version 0.9  3/10/15
  working script ready for production

 Version 0.91 5/6/15
  added Change order rate pdf report

 Version 0.92 8/7/15
  cleaned code around change order rate pdf.  was not updating

  future enhancement - formalize powershell formatting. commandbinding()
      -create function tool - 
      -building error handling
      -refine comment - will accomplish with commandbinding()
      -automate access report - prequal report, save PDF
      -log how often script runs for successfactors

  this script save time on a lot of small quick steps
  saving per script run 30 minutes
  SPS 30 minutes
#>

$conndb = "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\database\eBAPI_be.accdb"

$logpath = "c:\rich\importlog.log"

$gd = get-date
$inputstringSTART =  "$gd : query script started"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringSTART -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 5


#start Access
$a = New-Object -Comobject Access.Application 
$a.OpenCurrentDatabase($conndb) 


$a.DoCmd.SetWarnings($false)

#update vr selection - clean up Media 5 vs M5
$a.DoCmd.OpenQuery("update vr select")

#update vr selected pidn - clean up consolidated projects, updating old project id to new consolidated project id
$a.DoCmd.OpenQuery("update vr selected PIDN")

$a.DoCmd.OpenQuery("update invoice Leo date received 01")

$a.DoCmd.OpenQuery("update invoice Leo date received 02")

$a.DoCmd.OpenQuery("update invoice Cardno date received 03")

$a.DoCmd.OpenQuery("QRY append INV_update to Inv")

$a.DoCmd.OpenQuery("CCChangeRateDetails")

$a.DoCmd.OpenQuery("CCApproved")

$a.DoCmd.OpenQuery("CICount")

$a.DoCmd.OpenQuery("QRY change orders")

$a.DoCmd.OpenQuery("MakeAgingCA")

$a.DoCmd.OpenQuery("MakeAgingChanges")

$a.DoCmd.OpenQuery("MakeAgingInvoice")

$a.DoCmd.OpenQuery("PreQualUtilization")

$a.DoCmd.OutputTo(3, "PQ selected with Budget and Contracted with Value", "PDF Format (*.pdf)", "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\PreQualified Consultant Utilization FY1516.pdf" )

$a.DoCmd.OutputTo(3, "ChangeOrderRate", "PDF Format (*.pdf)", "\\kpfs1\Public\shares\cfo\fds\CFM\CFM Shared Files\Reports\ChangeOrderRate.pdf" )

$a.DoCmd.SetWarnings($True)

$a.Quit()


$gd = get-date
$inputstringEND =  "$gd : Query script END"
Add-Content -Path $logpath -Value "********************" -Force
Add-Content -Path $logpath -Value $inputstringEND -Force
Add-Content -Path $logpath -Value "********************" -Force
start-sleep -s 5

