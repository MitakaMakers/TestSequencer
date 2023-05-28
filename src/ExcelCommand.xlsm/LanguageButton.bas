Attribute VB_Name = "LanguageButton"
Option Explicit

Sub ApplyButton_Click()
    Dim table() As Text
    Dim I As Long
    Dim cell As Range
    
    table = GetLangTable()
    For I = 0 To UBound(table)
        Set cell = Worksheets(table(I).sheetname).Cells(table(I).row, table(I).column)
        If cell.value <> table(I).value Then
            cell.value = table(I).value
        End If
    Next I
End Sub
