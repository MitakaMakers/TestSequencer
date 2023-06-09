Attribute VB_Name = "SearchButton"
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

Sub SearchButton_Click()
    Dim bkupSel As Range
    Dim cnLo    As ConnectLayout
    Dim ret     As Long
    Dim Id      As Long
    Dim row     As Long
    
    Set bkupSel = Selection
    
    AddDllDirectories (ThisWorkbook.Path)
    
    cnLo = GetCnLayout()
    For row = cnLo.startRow To cnLo.endRow
        Cells(row, cnLo.wireColumn).Select
        Sleep (10)
        Cells(row, cnLo.addressColumn).Select
        Sleep (10)
        Cells(row, cnLo.timeoutColumn).Select
        Sleep (10)
        Cells(row, cnLo.statusColumn).Select
        Sleep (10)
    Next row
    
    bkupSel.Select
End Sub
