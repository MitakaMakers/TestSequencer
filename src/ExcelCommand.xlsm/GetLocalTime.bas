Attribute VB_Name = "GetLocalTime"
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

Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

#If VBA7 And Win64 Then
    Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Function GetLocalTimeStr() As String
    Dim t As SYSTEMTIME
    Dim s As String
    Call GetLocalTime(t)
    s = Format(t.wYear, "0000") & "/"
    s = s & Format(t.wMonth, "00") & "/"
    s = s & Format(t.wDay, "00") & " "
    s = s & Format(t.wHour, "00") & ":"
    s = s & Format(t.wMinute, "00") & ":"
    s = s & Format(t.wSecond, "00") & "."
    s = s & Format(t.wMilliseconds, "000")
    GetLocalTimeStr = """" & s & """"
End Function


