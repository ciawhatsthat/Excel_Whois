Attribute VB_Name = "ImportCSV"
Public Sub ImportCSVFile()

Dim xFileName As Variant
Dim TargetSheet As Worksheet
Dim Rg As Range
Dim xAddress As String
Dim TargetRange As Range
Dim wb As Workbook
Dim wsData As Worksheet, wsPT As Worksheet

Set wb = ThisWorkbook

On Error Resume Next

'Create the DATA Sheet to import the csv to
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "DATA"
Set TargetSheet = Sheets("DATA")
TargetSheet.UsedRange.Clear
'open the file picker dialog
xFileName = Application.GetOpenFilename("CSV File (*.csv), *.csv", , , , False)

If xFileName = False Then Exit Sub
 On Error Resume Next
Set Rg = TargetSheet.Range("a1")
On Error GoTo 0
If Rg Is Nothing Then Exit Sub
xAddress = Rg.Address
With ActiveSheet.QueryTables.Add("TEXT;" & xFileName, Range(xAddress))
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .SavePassword = False
    .SaveData = True
    .RefreshPeriod = 0
    .TextFilePlatform = 936
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

End Sub


