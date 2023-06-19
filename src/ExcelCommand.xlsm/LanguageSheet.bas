Attribute VB_Name = "LanguageSheet"
'    Excel Commmand: An excel macro file to communicate some measurement insturuments.
'    Copyright (C) 2023 Takatoshi Yamaoka
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as
'    published by the Free Software Foundation, either version 3 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Option Explicit

Const langSheet         As String = "Language"
Const startRow          As Long = 6
Const endRow            As Long = 50
Const sheetColumn       As Long = 3
Const rowColumn         As Long = 4
Const columnColumn      As Long = 5
Const japaneseColumn    As Long = 6
Const englishColumn     As Long = 7

Type Text
    sheetname As String
    row       As Long
    column    As Long
    value     As String
End Type

Function GetLangTable(valueColumn As Long) As Text()
    Dim sheet   As Worksheet
    Dim column  As Long
    Dim i       As Long
    Dim table() As Text
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = langSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[language]ÉVÅ[ÉgÇ™Ç†ÇËÇ‹ÇπÇÒ", vbInformation
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

Sub JapaneseButton_Click()
    Dim table() As Text
    Dim i       As Long
    Dim cell    As Range
    
    table = GetLangTable(japaneseColumn)
    For i = 0 To UBound(table)
        Set cell = Worksheets(table(i).sheetname).Cells(table(i).row, table(i).column)
        If cell.value <> table(i).value Then
            cell.value = table(i).value
        End If
    Next i
End Sub

Sub EnglishButton_Click()
    Dim table() As Text
    Dim i       As Long
    Dim cell    As Range
    
    table = GetLangTable(englishColumn)
    For i = 0 To UBound(table)
        Set cell = Worksheets(table(i).sheetname).Cells(table(i).row, table(i).column)
        If cell.value <> table(i).value Then
            cell.value = table(i).value
        End If
    Next i
End Sub
