Attribute VB_Name = "LanguageSheet"
Option Explicit

Const langSheet    As String = "99_language"
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
    Dim I As Long
    Dim table() As Text
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = langSheet Then
            sheet.name = langSheet
            Exit For
        End If
    Next sheet
    
    If sheet.name <> langSheet Then
        MsgBox "[language]ÉVÅ[ÉgÇÕÇ†ÇËÇ‹ÇπÇÒ", vbInformation
        Exit Function
    End If
    
    ReDim table(endRow - startRow)
    For I = 0 To endRow - startRow
        table(I).sheetname = CStr(sheet.Cells(startRow + I, sheetColumn).value)
        table(I).row       = CLng(sheet.Cells(startRow + I, rowColumn).value)
        table(I).column    = CLng(sheet.Cells(startRow + I, columnColumn).value)
        table(I).value     = CStr(sheet.Cells(startRow + I, valueColumn).value)
    Next I
    GetLangTable = table
End Function

