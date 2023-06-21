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

Sub ChangeLanguage(valueColumn As Long, lnLo As LanguageLayout)
    Dim bkupSel As Range
    Dim sheet     As Worksheet
    Dim i         As Long
    Dim sheetname As String
    Dim row       As Long
    Dim column    As Long
    Dim value     As String
    Dim cell      As Range
    
    Set bkupSel = Selection
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = lnLo.sheetname Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & lnLo.sheetname & "]ÉVÅ[ÉgÇ™å©Ç¬Ç©ÇËÇ‹ÇπÇÒ", vbInformation
        Exit Sub
    End If
    
    For i = 0 To lnLo.endRow - lnLo.startRow
        sheet.Cells(lnLo.startRow + i, lnLo.sheetColumn).Select
        sheetname = CStr(sheet.Cells(lnLo.startRow + i, lnLo.sheetColumn).value)
        If sheetname = "END" Then
            Exit For
        End If
        sheet.Cells(lnLo.startRow + i, lnLo.rowColumn).Select
        row = CLng(sheet.Cells(lnLo.startRow + i, lnLo.rowColumn).value)
        sheet.Cells(lnLo.startRow + i, lnLo.columnColumn).Select
        column = CLng(sheet.Cells(lnLo.startRow + i, lnLo.columnColumn).value)
        sheet.Cells(lnLo.startRow + i, valueColumn).Select
        value = CStr(sheet.Cells(lnLo.startRow + i, valueColumn).value)
        Set cell = Worksheets(sheetname).Cells(row, column)
        If cell.value <> value Then
            cell.value = value
        End If
    Next i
    
    bkupSel.Select
End Sub

Sub JapaneseButton_Click()
    Dim lnLo    As LanguageLayout
    
    lnLo = GetLangLayout()
    Call ChangeLanguage(lnLo.japaneseColumn, lnLo)
End Sub

Sub EnglishButton_Click()
    Dim lnLo    As LanguageLayout
    
    lnLo = GetLangLayout()
    Call ChangeLanguage(lnLo.englishColumn, lnLo)
End Sub
