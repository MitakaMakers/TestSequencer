Attribute VB_Name = "LanguageSheet"
Option Explicit

Const langSheet    As String = "98_language"
Const startRow     As Long = 9
Const endRow       As Long = 49
Const sheetColumn  As Long = 3
Const rowColumn    As Long = 4
Const columnColumn As Long = 5
Const valueColumn  As Long = 7

Type Text
    sheetname As String
    row       As Long
    column    As Long
    value     As String
End Type

Function GetLangTable() As Text()
    Dim sheet As Worksheet
    Dim column As Long
    Dim i As Long
    Dim table() As Text
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = langSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[language]ÉVÅ[ÉgÇÕÇ†ÇËÇ‹ÇπÇÒ", vbInformation
        Exit Function
    End If
    
    ReDim table(endRow - startRow)
    For i = 0 To endRow - startRow
        table(i).sheetname = CStr(sheet.Cells(startRow + i, sheetColumn).value)
        table(i).row = CLng(sheet.Cells(startRow + i, rowColumn).value)
        table(i).column = CLng(sheet.Cells(startRow + i, columnColumn).value)
        table(i).value = CStr(sheet.Cells(startRow + i, valueColumn).value)
    Next i
    GetLangTable = table
End Function
