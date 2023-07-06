Attribute VB_Name = "CommandSheet"
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

Sub SearchButton_Click(ByVal Target As Range)
    Dim bkupSel    As Range
    Dim cnLo       As ConnectLayout
    Dim ret        As Long
    Dim Id         As Long
    Dim row        As Long
    Dim devices    As DeviceListArray
    Dim num        As Long
    Dim i          As Long
    Dim value      As Variant
    Dim wire       As Long
    Dim validation As validation
    
    Set bkupSel = Selection
    cnLo = GetCnLayout()
    
    If Target.row >= cnLo.startRow And Target.row <= cnLo.endRow Then
        If Target.column = cnLo.wireColumn Then
            wire = GetWireType(Target.value)
            If wire = 7 Or wire = 10 Then
                Application.EnableEvents = False
                AddDllDirectories (ThisWorkbook.Path)
                ret = TmSearchDevices(wire, devices, 8, num, "")
                If 0 < num Then
                    For i = 0 To num
                        Cells(7 + i, cnLo.addressColumn).value = devices.list(i).adr
                    Next i
                    Set validation = Cells(Target.row, Target.column + 1).validation
                    validation.Delete
                    validation.Add _
                        Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Formula1:="=$E$7:$E$15"
                    Cells(Target.row, Target.column + 1).value = devices.list(0).adr
                End If
                Application.EnableEvents = True
            End If
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
    Application.EnableEvents = False
    
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
            If 0 <= i And Id(i) <> -1 Then
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
        ElseIf value = "IMAGE" Then
            Cells(row, cmdLo.arg1Column).Select
            value = CStr(Cells(row, cmdLo.arg1Column))
            i = CInt(value) - 1
            If 0 <= i And Id(i) <> -1 Then
                Cells(row, cmdLo.arg2Column).Select
                value = Cells(row, cmdLo.arg2Column).value
                cmd = CStr(value)
                ret = TmSend(Id(i), cmd)
                If InStr(cmd, "?") Then
                    Cells(row, cmdLo.resultColumn).Select
                    ret = TmReceiveToFile(Id(i), "temp.bmp", rlen)
                    If ret = 0 Then
                        ActiveSheet.Pictures.Insert("temp.bmp").Select
                        Selection.Top = Cells(row, cmdLo.resultColumn).Top
                        Selection.Left = Cells(row, cmdLo.resultColumn).Left
                    End If
                End If
                Cells(row, cmdLo.statusColumn).Select
                value = GetLocalTimeStr()
                Cells(row, cmdLo.statusColumn).value = value
            End If
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
    Application.EnableEvents = True
    bkupSel.Select
End Sub

Function Sleep(Second As Long)
    Application.Wait [Now()] + (Second / (86400))
End Function

Function GetWireType(wire As Variant) As Long
    Dim str As String
    
    If IsEmpty(wire) Then
        GetWireType = 0
    Else
        If TypeName(wire) = "String" Then
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
        If TypeName(eos) = "String" Then
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
        Else
            GetTermType = 0
        End If
    End If
End Function
