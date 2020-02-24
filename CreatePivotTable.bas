Attribute VB_Name = "CreatePivotTable"
'This is a pivot sheet very specific to my needs, see https://www.youtube.com/watch?v=gfnAwpcWv3I
'and https://github.com/DataSolveProblems/Excel-Exercise-Files-and-VBA-Code/tree/master/Create%20Pivot%20Table%20with%20Excel%20VBA

Public Sub Create_Pivot_Table()
    Dim lastRow As Long, LastColumn As Long
    Dim DataRange As Range
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    
    'On Error GoTo errHandle
    
    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("DATA")
    
    '// Delete Pivot Table sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Pivot Table").Delete
    Application.DisplayAlerts = True
    
    '// Create Data Range variable
    With wsData
    lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    
        Set DataRange = .Range(.Cells(1, 1), .Cells(lastRow, LastColumn))
    End With
    
    '// Create Pivot Table Worksheet
    Set wsPT = wb.Worksheets.Add
    wsPT.Name = "Pivot Table"
    
    '// Storing Pivot Table Cache
    Set PTCache = wb.PivotCaches.Create(xlDatabase, DataRange)
    
    '// Create Pivot Table
    Set PT = PTCache.CreatePivotTable(wsPT.Range("A1"), "Spoofed IPs")
    
    
    '// Adding Columns, Rows and Data to pivot table
    With PT
    
        '// Pivot Table Layout
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False 'Optional (Column Grand Total)
        .RowGrand = False 'Optional (Row Grand Total)
        '.Name = "spoofpiv"
        
        .TableStyle2 = "PivotStyleMedium9"
        .HasAutoFormat = False 'Re-Format Pivot Table when refresh
        .SubtotalLocation xlAtTop 'Position SubTotal on the top or bottom
        
       
        ' Row Section (Layer 1)
        With .PivotFields("IP Address")
            .Orientation = xlRowField
            .Position = 1

        End With
    
        With .PivotFields("IP Address")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlCount
            .Caption = "Count"
        End With
        
        With .PivotFields("IP Address")
            .AutoSort xlDescending, "Count"
            
        End With
        
        
    End With
       
    
    With wsPT.Range("C1", "E1")
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
    
    wsPT.Range("C1").Value = "Name"
    wsPT.Range("D1").Value = "CIDR"
    wsPT.Range("E1").Value = "Country"
 
    wsPT.Cells.EntireColumn.AutoFit
    
    
ClearObjects:
    Set PTCache = Nothing
    Set PT = Nothing
    Set DataRange = Nothing
    Set wsData = Nothing
    Set wb = Nothing

Exit Sub

errHandle:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo ClearObjects

End Sub
