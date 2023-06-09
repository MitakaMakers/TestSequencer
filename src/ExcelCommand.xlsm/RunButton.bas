Attribute VB_Name = "RunButton"
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

Sub RunButton_Click()
    Dim bkupSel As Range
    Dim cmdLo   As CommandLayout
    Dim excOpt  As ExecOption
    Dim cnLo    As ConnectLayout
    Dim sheet   As Worksheet
    Dim str     As String
    Dim Id()    As Long
    Dim i       As Integer
    Dim value   As Variant
    Dim wire    As Long
    Dim adr     As String
    Dim ret     As Long
    Dim cmd     As String
    Dim rlen    As Long
    Dim row     As Long

    Set bkupSel = Selection
    
    cmdLo = GetCmdLayout()
    excOpt = GetExecOption()
    cnLo = GetCnLayout()
    
    AddDllDirectories (ThisWorkbook.Path)
    
    ReDim Id(cnLo.endRow - cnLo.startRow + 1)
    For row = cnLo.startRow To cnLo.endRow
        i = row - cnLo.startRow
        If Id(i) <> -1 Then
            ret = TmFinish(Id(i))
            Id(i) = -1
            Cells(row, cnLo.statusColumn).value = ""
        End If
    Next row
    
    For row = cnLo.startRow To cnLo.endRow
        Cells(row, cnLo.wireColumn).Select
        value = Cells(row, cnLo.wireColumn).value
        If Not IsEmpty(value) Then
            wire = CInt(Left(value, InStr(value, " ")))
            i = row - cnLo.startRow
            If 0 < wire Then
                Cells(row, cnLo.addressColumn).Select
                value = Cells(row, cnLo.addressColumn).value
                If TypeName(value) = "String" Then
                    adr = CStr(value)
                    ret = TmInitialize(wire, adr, Id(i))
                    If ret = 0 Then
                        Cells(row, cnLo.timeoutColumn).Select
                        Call TmSetTimeout(Id(i), 100)
                        Call TmSetTerm(Id(i), 1, 2)
                        Call TmDeviceClear(Id(i))
                        Cells(row, cnLo.statusColumn).value = "Connected."
                    End If
                 End If
                 ret = TmGetLastError(Id(i))
             End If
        End If
    Next row
    For row = cmdLo.startRow To cmdLo.endRow
        Cells(row, cmdLo.deviceColumn).Select
        value = Cells(row, cmdLo.deviceColumn).value
        If Not IsEmpty(value) Then
            i = CInt(value) - 1
            If 0 < i And Id(i) <> -1 Then
                If 0 < excOpt.interval Then
                    Sleep (excOpt.interval)
                End If
                Cells(row, cmdLo.commandColumn).Select
                value = Cells(row, cmdLo.commandColumn).value
                cmd = CStr(value)
                ret = TmSend(Id(i), cmd)
                If InStr(cmd, "?") Then
                    Cells(row, cmdLo.responseColumn).Select
                    str = String(65536, vbNullChar)
                    ret = TmReceive(Id(i), str, 65536, rlen)
                    str = Left$(str, rlen - 1)
                    Cells(row, cmdLo.responseColumn).value = str
                End If

                Cells(row, cmdLo.statusColumn).Select
                ret = TmGetLastError(Id(i))
                value = GetLocalTimeB()
                Cells(row, cmdLo.statusColumn).value = value
            End If
        End If
    Next row
    For row = cnLo.startRow To cnLo.endRow
        i = row - cnLo.startRow
        If Id(i) <> -1 Then
            ret = TmFinish(Id(i))
            Id(i) = -1
            Cells(row, cnLo.statusColumn).value = ""
        End If
    Next row
    bkupSel.Select
End Sub

Function Sleep(Time As Long)
    Application.Wait [Now()] + Time / 86400000
End Function

Function GetLocalTimeB() As String
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
    GetLocalTimeB = """" & s & """"
End Function
