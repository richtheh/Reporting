<#
Name:  FAAADJScript
Author: rcsh
Date: 4/11/19
Version: 0.9

Purpose:  read in custom report adj&midyr to collect individual std data, parse to individual files, update award adjustment template

TODO
Parameterize: source file name
Parameterize: save to directory based on award adj code
Create function

#>

$pathxls = 'C:\reporting\Adjustment_Template.xlsx'
$savedir = 'C:\reporting\new\'

# convert PF Custom report to individual std csv files by std name

$file = 'C:\reporting\CustomReport_ADJ&MIDYR.csv'

$i = 0

ForEach ($line in (import-csv $file)) {
    $FN = $line | select -ExpandProperty FullName
    $N = $FN.Replace('.csv', '')
    $AA = $line | select -ExpandProperty 'Award ADJ'

    $line | export-csv "C:\reporting\new\$FN.csv"
    $FN
    $i++

    #create Excel object
    $Excel = New-Object -ComObject Excel.Application

    $Excel.Visible = $False

    #Open std data workbook
    $WorkBook = $Excel.Workbooks.Open("C:\reporting\new\$FN.csv")
    $WorkSheet = $WorkBook.worksheets.item(1)
    
    $Worksheet.Activate()

    #select & copy range
    $Range2 = $WorkSheet.Range("A1:T3").currentregion
    $Range2.Copy() | out-null

    #open Template workbook
    $Template = $Excel.Workbooks.Open($pathxls)
    $TemplateWorksheet = $Template.worksheets.item('Sheet1')
    $TemplateWorksheet.activate()

    #paste data in template sheet1
    $Range1 = $TemplateWorksheet.Range("A1:T3")
    $TemplateWorksheet.Paste($Range1)

    #activate Template sheet
    $TWS = $Template.worksheets.item('Template')
    $TWS.activate()


    #save files

    # TO DO parameterize the directory selection
    If ($AA -eq '5') {
        $Template.SaveAs("c:\reporting\MidYearGrad\$N.xlsx",51)
    } ElseIf ($AA -eq '7') {
        $Template.SaveAs("c:\reporting\PartTime\$N.xlsx",51)
    } ElseIf ($AA -eq '9') {
        $Template.SaveAs("c:\reporting\ChangeOfProgramMidYear\$N.xlsx",51)
    } ElseIf ($AA -eq '11') {
        $Template.SaveAs("c:\reporting\ChangeOfSchool\$N.xlsx",51)
    } ElseIf ($AA -eq '13') {
        $Template.SaveAs("c:\reporting\MidYearTransfer\$N.xlsx",51)
    } ElseIf ($AA -eq '15') {
        $Template.SaveAs("c:\reporting\LateStart\$N.xlsx",51)
    } ElseIf ($AA -eq '23') {
        $Template.SaveAs("c:\reporting\Multiple\$N.xlsx",51)
    } Else {  $Template.SaveAs("c:\reporting\Other\$N.xlsx",51)
    
    }

    #close workbooks
    $WorkBook.Close($false)
    $Template.Close()

#close Excel
$Excel.Quit()
    


}

#clean up *.csv files
Get-ChildItem -path C:\reporting\new -Filter "*.csv" | Remove-Item

