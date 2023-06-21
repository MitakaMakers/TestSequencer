Attribute VB_Name = "LoadLibrary"
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

Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal fileName As String) As Long

Sub AddDllDirectories(DLLFoldPath As String)
#If VBA7 And Win64 Then
    LoadLibrary (DLLFoldPath & "\YKMUSB64.dll")
    LoadLibrary (DLLFoldPath & "\tmctl64.dll")
#Else
    LoadLibrary (DLLFoldPath & "\YKMUSB.dll")
    LoadLibrary (DLLFoldPath & "\tmctl.dll")
#End If
End Sub
