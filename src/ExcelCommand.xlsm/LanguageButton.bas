Attribute VB_Name = "LanguageButton"
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

Sub ApplyButton_Click()
    Dim table() As Text
    Dim i       As Long
    Dim cell    As Range
    
    table = GetLangTable()
    For i = 0 To UBound(table)
        Set cell = Worksheets(table(i).sheetname).Cells(table(i).row, table(i).column)
        If cell.value <> table(i).value Then
            cell.value = table(i).value
        End If
    Next i
End Sub
