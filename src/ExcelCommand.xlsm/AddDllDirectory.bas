Attribute VB_Name = "AddDllDirectory"
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

Private Declare PtrSafe Function SetDefaultDllDirectories Lib "kernel32.dll" (ByVal DirectoryFlags As Long) As Long
Private Declare PtrSafe Function AddDllDirectory Lib "kernel32.dll" (ByVal fileName As String) As LongPtr

Const LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = &H1000

Sub AddDllDirectories(DLLFoldPath As String)
    SetDefaultDllDirectories (LOAD_LIBRARY_SEARCH_DEFAULT_DIRS)
    AddDllDirectory (StrConv(DLLFoldPath, vbUnicode))
End Sub
