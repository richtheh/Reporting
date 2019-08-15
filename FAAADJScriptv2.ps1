<#
Name:  FAAADJScript
Author: rcsh
Date: 4/11/19
Version: 1.0
Version Date 8/15/19

Purpose:  read in custom report adj&midyr to collect individual std data, parse to individual files, update award adjustment template

TODO
Add code for each adjustment code/add other templates based on adj code - DONE
Parameterize: source file name
Parameterize: save to directory based on award adj code -DONE
Create function

8/15/19
-updated variables and templates to network drive

add step to concate name in csv/xlsx

future - test parttime award amount against static pool

#>

# 
$connsource = '\\kpfs1\public\shares\teams\FADC\PH\2020 - 2021\1 College Need Based\1 FAA\Award Adj_Static Pool'
$templatesource = '\\kpfs1\public\shares\teams\FADC\PH\2020 - 2021\1 College Need Based\1 FAA\Award Adj_Static Pool\Templates'

# Source CSV file
$SourceCSV = '\\kpfs1\public\shares\teams\FADC\PH\2020 - 2021\1 College Need Based\1 FAA\Award Adj_Static Pool\CustomReport_ADJ&MIDYR.csv'

# Adjustment templates
$TMidYearGrad = "$templatesource\MidYearGrad_Template.xlsx"
#$TMidYearGrad = '\\kpfs1\public\shares\teams\FADC\PH\2020 - 2021\1 College Need Based\1 FAA\Award Adj_Static Pool\Templates\MidYearGrad_Template.xlsx'
$TPartTime = "$templatesource\PartTime_Template.xlsx"
$TChangeProgramMidYear = "$templatesource\ChangeProgramMidYear_Template.xlsx"
$TChangeOfSchool = "$templatesource\ChangeOfSchool_Template.xlsx"
$TMidYearTransfer = "$templatesource\MidYearTransfer_Template.xlsx"
$TLateStart = "$templatesource\LateStart_Template.xlsx"
$TMultiple = "$templatesource\Multiple_Template.xlsx"
$TOther = "$templatesource\Other_Template.xlsx"


# open PF Custom report, run line by line and save to individual std csv files by std name

$i = 0

ForEach ($line in (import-csv $SourceCSV)) {
    $FN = $line | select -ExpandProperty FullName
    $N = $FN.Replace('.csv', '')
    $AA = $line | select -ExpandProperty 'Award ADJ'

    $line | export-csv "$connsource\Holding\$FN.csv"
    $FN
    $i++

    #create Excel object
    $Excel = New-Object -ComObject Excel.Application

    $Excel.Visible = $False

    #Open std data workbook
    $WorkBook = $Excel.Workbooks.Open("$connsource\Holding\$FN.csv")
    $WorkSheet = $WorkBook.worksheets.item(1)
    
    $Worksheet.Activate()

    #select & copy range
    $Range2 = $WorkSheet.UsedRange
    $Range2.Copy() | out-null

    #iterate through records based on Award Adj
    If ($AA -eq '5') {

            $ADJTemplate = $Excel.Workbooks.Open($TMidYearGrad)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory 
            $ADJTemplate.SaveAs("$connsource\05 - Mid-Year Grad\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()

        } ElseIf ($AA -eq '7') {

            $ADJTemplate = $Excel.Workbooks.Open($TPartTime)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\07 - Part-Time\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()

<#	
        } ElseIf ($AA -eq '9') {

            $ADJTemplate = $Excel.Workbooks.Open($TChangeProgramMidYear)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("c:\reporting\ChangeOfProgramMidYear\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()
#>
        } ElseIf ($AA -eq '11') {

            $ADJTemplate = $Excel.Workbooks.Open($TChangeOfSchool)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\11 - Change Of School\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()

        } ElseIf ($AA -eq '13') {

            $ADJTemplate = $Excel.Workbooks.Open($TMidYearTransfer)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\13 - Mid-Year Transfer\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()

        } ElseIf ($AA -eq '15') {

            $ADJTemplate = $Excel.Workbooks.Open($TLateStart)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\15 - Late Start\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()

        } ElseIf ($AA -eq '23') {

            $ADJTemplate = $Excel.Workbooks.Open($TMultiple)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()

        #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\23 - Multiple\$N.xlsx",51)

        #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()
    

        } Else {  
            $ADJTemplate = $Excel.Workbooks.Open($TOther)
            $TemplateWorksheet = $ADJTemplate.worksheets.item('Sheet1')
            $TemplateWorksheet.activate()

        #paste std data in correct template sheet1
            $Range1 = $TemplateWorksheet.Range("A1")
            $TemplateWorksheet.Paste($Range1)

        #activate Template sheet
            $TWS = $ADJTemplate.worksheets.item('Template')
            $TWS.activate()
        
         #save files to correct directory
            $ADJTemplate.SaveAs("$connsource\99 - Other\$N.xlsx",51)

       #close workbooks
            #$WorkBook.Close($false)
            $ADJTemplate.Close()
    
    }
    


    #close workbooks
    $WorkBook.Close($false)
    #$Template.Close()

#close Excel
$Excel.Quit()
    


}


#clean up *.csv files
$count = Get-ChildItem -path $connsource\Holding -Filter *.csv | Measure-Object
$cc = $count.count
$cc
Get-ChildItem -path $connsource\Holding -Filter *.csv | Remove-Item
$count = Get-ChildItem -path $connsource\Holding -Filter *.csv | Measure-Object
$cc = $count.count
$cc
<#


Get-ChildItem -path "$connsource\11 - Change Of School" | Remove-Item
Get-ChildItem -path "$connsource\15 - Late Start" | Remove-Item
Get-ChildItem -path "$connsource\05 - Mid-Year Grad" | Remove-Item
Get-ChildItem -path "$connsource\13 - Mid-Year Transfer" | Remove-Item
Get-ChildItem -path "$connsource\23 - Multiple" | Remove-Item
Get-ChildItem -path "$connsource\99 - Other" | Remove-Item
Get-ChildItem -path "$connsource\07 - Part-Time" | Remove-Item
Get-ChildItem -path "$connsource\Holding -Filter *.csv" | Remove-Item


#>




