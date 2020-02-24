Attribute VB_Name = "Iinstantiate"
Public Sub instantiate()

'Stop screen updating to not see the voodoo
Application.ScreenUpdating = False

'import the CSV
Call ImportCSV.ImportCSVFile

Call Create_Pivot_Table

Application.DisplayAlerts = False
ThisWorkbook.Worksheets("DATA").Delete
Application.DisplayAlerts = True

Worksheets("Pivot Table").Activate

Application.ScreenUpdating = True

Call WhoIs.whoismacro

With ActiveSheet
 .Cells.EntireColumn.AutoFit
 .Cells.EntireRow.AutoFit
End With

End Sub



