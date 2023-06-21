Attribute VB_Name = "CommandSheet"
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

Sub SearchButton_Click(ByVal Target As Range)
    Dim bkupSel As Range
    Dim cnLo    As ConnectLayout
    Dim ret     As Long
    Dim Id      As Long
    Dim row     As Long
    Dim devices As DeviceListArray
    Dim num     As Long
    Dim i       As Long
    Dim value   As Variant
    Dim wire    As Long
    
    Set bkupSel = Selection
    cnLo = GetCnLayout()
    If Target.column = cnLo.wireColumn Then
        If cnLo.startRow <= Target.column And Target.column <= cnLo.endRow Then
            AddDllDirectories (ThisWorkbook.Path)
            ret = TmSearchDevices(7, devices, 128, num, "")
            For i = 0 To num
                Cells(row, cnLo.wireColumn).Select
                value = Cells(row, cnLo.wireColumn).value
                wire = GetWireType(value)
                If wire = 7 Then
                     If i < num Then
                         Cells(row, cnLo.addressColumn).Select
                         Cells(row, cnLo.addressColumn).value = devices.list(i).adr
                         i = i + 1
                     End If
                End If
            Next i
        End If
    End If
    bkupSel.Select
End Sub

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
    Dim eos     As Long
    Dim eot     As Long

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
        wire = GetWireType(value)
        If 0 < wire Then
            Cells(row, cnLo.addressColumn).Select
            value = Cells(row, cnLo.addressColumn).value
            adr = CStr(value)
            If Not IsEmpty(value) Then
                i = row - cnLo.startRow
                ret = TmInitializeEx(wire, adr, Id(i), 10)
                If ret = 0 Then
                    value = excOpt.timeout / 100
                    Call TmSetTimeout(Id(i), CInt(value))
                    Cells(row, cnLo.termColumn).Select
                    value = Cells(row, cnLo.termColumn).value
                    eos = GetTermType(value)
                    Call TmSetTerm(Id(i), eos, 1)
                    Cells(row, cnLo.statusColumn).Select
                    Call TmDeviceClear(Id(i))
                    Cells(row, cnLo.statusColumn).value = "Connected."
                Else
                    Cells(row, cnLo.statusColumn).value = ""
                End If
            End If
        End If
    Next row
    
    For row = cmdLo.startRow To cmdLo.endRow
        Cells(row, cmdLo.opColumn).Select
        value = Cells(row, cmdLo.opColumn).value
        If IsEmpty(value) Then
        ElseIf value = "END" Then
            Cells(row, cmdLo.statusColumn).Select
            value = GetLocalTimeStr()
            Cells(row, cmdLo.statusColumn).value = value
            Exit For
        ElseIf value = "PRINT" Then
            Cells(row, cmdLo.arg1Column).Select
            value = CStr(Cells(row, cmdLo.arg1Column))
            MsgBox value, vbOKOnly
            Cells(row, cmdLo.statusColumn).Select
            value = GetLocalTimeStr()
            Cells(row, cmdLo.statusColumn).value = value
        ElseIf value = "WRITE" Then
            Cells(row, cmdLo.arg1Column).Select
            value = CStr(Cells(row, cmdLo.arg1Column))
            i = CInt(value) - 1
            If 0 < i And Id(i) <> -1 Then
                Cells(row, cmdLo.arg2Column).Select
                value = Cells(row, cmdLo.arg2Column).value
                cmd = CStr(value)
                ret = TmSend(Id(i), cmd)
                If InStr(cmd, "?") Then
                    Cells(row, cmdLo.resultColumn).Select
                    str = String(65536, vbNullChar)
                    ret = TmReceive(Id(i), str, 65536, rlen)
                    str = Left$(str, rlen - 1)
                    Cells(row, cmdLo.resultColumn).value = str
                End If
                Cells(row, cmdLo.statusColumn).Select
                value = GetLocalTimeStr()
                Cells(row, cmdLo.statusColumn).value = value
            End If
        ElseIf value = "WAIT" Then
            Cells(row, cmdLo.arg1Column).Select
            value = CInt(Cells(row, cmdLo.arg1Column))
            Sleep (value)
            Cells(row, cmdLo.statusColumn).Select
            value = GetLocalTimeStr()
            Cells(row, cmdLo.statusColumn).value = value
        End If
        Sleep (excOpt.interval)
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

Function Sleep(Second As Long)
    Application.Wait [Now()] + (Second / (24 * 60 * 60))
End Function

Function GetWireType(wire As Variant) As Long
    Dim str As String
    
    If IsEmpty(wire) Then
        GetWireType = 0
    Else
        str = CStr(wire)
        If str = "GP-IB" Then
            GetWireType = 1
        ElseIf str = "RS232C" Then
            GetWireType = 2
        ElseIf str = "USB" Then
            GetWireType = 3
        ElseIf str = "ETHERNET" Then
            GetWireType = 4
        ElseIf str = "USBTMC2" Then
            GetWireType = 7
        ElseIf str = "VXI-11" Then
            GetWireType = 8
        ElseIf str = "VISAUSB" Then
            GetWireType = 10
        ElseIf str = "Socket" Then
            GetWireType = 11
        ElseIf str = "HiSLIP" Then
            GetWireType = 14
        Else
            GetWireType = 0
        End If
    End If
End Function

Function GetTermType(eos As Variant) As Long
    Dim str As String
    
    If IsEmpty(eos) Then
        GetTermType = 0
    Else
        str = CStr(eos)
        If str = "CRLF" Then
            GetTermType = 0
        ElseIf str = "CR" Then
            GetTermType = 1
        ElseIf str = "LF" Then
             GetTermType = 2
        Else
            GetTermType = 3
        End If
    End If
End Function
