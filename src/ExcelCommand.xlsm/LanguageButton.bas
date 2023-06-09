Attribute VB_Name = "LanguageButton"
Option Explicit

Sub ApplyButton_Click()
    Dim table() As Text
    Dim i As Long
    Dim cell As Range
    
    table = GetLangTable()
    For i = 0 To UBound(table)
        Set cell = Worksheets(table(i).sheetname).Cells(table(i).row, table(i).column)
        If cell.value <> table(i).value Then
            cell.value = table(i).value
        End If
    Next i
End Sub
