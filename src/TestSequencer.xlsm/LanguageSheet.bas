Attribute VB_Name = "LanguageSheet"
'   Test Sequencer: A macro file to communicate measurement insturuments.
'   Copyright 2023 Takatoshi Yamaoka
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

Option Explicit

Sub ChangeLanguage(valueColumn As Long, lnLo As LanguageLayout)
    Dim bkupSel   As Range
    Dim sheet     As Worksheet
    Dim i         As Long
    Dim sheetname As String
    Dim row       As Long
    Dim column    As Long
    Dim value     As String
    Dim cell      As Range
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = lnLo.sheetname Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & lnLo.sheetname & "]ÉVÅ[ÉgÇ™å©Ç¬Ç©ÇËÇ‹ÇπÇÒ", vbInformation
        Exit Sub
    End If
    
    Set bkupSel = Selection
    
    Application.EnableEvents = False
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
    Application.EnableEvents = True
    
    bkupSel.Select
End Sub

Sub JapaneseButton_Click()
    Dim lnLo As LanguageLayout
    lnLo = GetLangLayout()
    Call ChangeLanguage(lnLo.japaneseColumn, lnLo)
End Sub

Sub EnglishButton_Click()
    Dim lnLo As LanguageLayout
    lnLo = GetLangLayout()
    Call ChangeLanguage(lnLo.englishColumn, lnLo)
End Sub
