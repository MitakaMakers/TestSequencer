Attribute VB_Name = "GetLocalTime"
'   Test Sequencer: A macro file to communicate measurement insturuments.
'   Copyright 2023 Takatoshi Yamaoka (mitaka.lab@gmail.com)
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


