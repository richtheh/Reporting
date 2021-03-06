
# Specify the path to the Excel file and WorkSheet Name

$OutFile = "C:\reporting\out.csv"
$FilePath = "C:\reporting\sourcereport_ADJ&MIDYR.xlsx"
$SheetName = "ADJ"

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in Excel
$objExcel.visible = $False

$sl = Set-Location c:\reporting
$sl
Get-childitem

# Open the Excel file an dsave it in $WorkBook
$Workbook = $objExcel.Workbooks.Open($FilePath)

# Load the Worksheet 'ADJ'
$Worksheet = $WorkBook.sheets.item($Sheetname)

$AwardADJ = $Worksheet | select "Award ADJ" | sort-object -Property "Award ADJ" -Unique

$AwardADJ

$Worksheet

<#
$Worksheet.Range("B7").Text

$Output = [pscustomobject]  @{
    SSN = $Worksheet.Range("A7").Text
    LastName = $Worksheet.Range("B7").Text
    FirstName = $Worksheet.Range("C7").Text
    Degree = $Worksheet.Range("E7").Text
    Name_PH = $Worksheet.Range("J7").Text
    Budget = $Worksheet.Range("K7").Text
    EFC = $Worksheet.Range("L7").Text
    Need = $Worksheet.Range("M7").Text
}

$Output | Export-csv $Outfile
#>


# Close Excel
$objExcel.quit()

#$objExcel.WorkBooks | Select-Object -Property name, path, author
#$objExcel.WorkBooks | Get-Member 
#$Workbook | Get-Member -Name *sheet*
#$WorkBook.sheets | Select-Object -Property Name


#$WorkSheet = $Workbook.Sheets.Item(1)
#$Worksheet.Name
#$Found = $WorkSheet.Cells.Find('Fall')
#$BeginAddress = $Found.Address(0,0,1,1)
#$BeginAddress
