Attribute VB_Name = "tmctl"
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

Public Const MaxStationNum = 8

Type DeviceList
        adr As String * 64
End Type

Type DeviceListArray
    list(MaxStationNum) As DeviceList
End Type

#If VBA7 And Win64 Then
Declare PtrSafe Function TmInitialize Lib "tmctl64.dll" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long) As Long
Declare PtrSafe Function TmInitializeEx Lib "tmctl64.dll" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmInitializeExA Lib "tmctl64.dll" Alias "TmInitializeEx" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmDeviceClear Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmDeviceTrigger Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSend Lib "tmctl64.dll" (ByVal Id As Long, ByVal msg As String) As Long
Declare PtrSafe Function TmReceiveBin Lib "tmctl64.dll" Alias "TmReceive" (ByVal Id As Long, ByRef buf As Any, ByVal blen As Long, ByRef rlen As Long) As Long
Declare PtrSafe Function TmReceiveSetup Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmReceiveBlockHeader Lib "tmctl64.dll" (ByVal Id As Long, ByRef rlen As Long) As Long
Declare PtrSafe Function TmReceiveBlockData Lib "tmctl64.dll" (ByVal Id As Long, ByRef buf As Any, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long) As Long
Declare PtrSafe Function TmCheckEnd Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSetRen Lib "tmctl64.dll" (ByVal Id As Long, ByVal flg As Long) As Long
Declare PtrSafe Function TmSetTerm Lib "tmctl64.dll" (ByVal Id As Long, ByVal eos As Long, ByVal eot As Long) As Long
Declare PtrSafe Function TmSetTimeout Lib "tmctl64.dll" (ByVal Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmcGetStatusByte Lib "tmctl64.dll" (ByVal Id As Long, ByRef sts As Byte) As Long
Declare PtrSafe Function TmFinish Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmGetLastError Lib "tmctl64.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSearchDevices Lib "tmctl64.dll" (ByVal wire As Long, list As DeviceListArray, ByVal max As Long, ByRef num As Long, ByVal option1 As String) As Long
Declare PtrSafe Function TmEncodeSerialNumber Lib "tmctl64.dll" (ByVal encode As String, ByVal encodelen As Long, ByVal src As String) As Long
Declare PtrSafe Function TmDecodeSerialNumber Lib "tmctl64.dll" (ByVal decode As String, ByVal decodelen As Long, ByVal src As String) As Long
#Else
Declare PtrSafe Function TmInitialize Lib "tmctl.dll" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long) As Long
Declare PtrSafe Function TmInitializeEx Lib "tmctl.dll" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmInitializeExA Lib "tmctl.dll" Alias "TmInitializeEx" (ByVal wire As Long, ByVal adr As String, ByRef Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmDeviceClear Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmDeviceTrigger Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSend Lib "tmctl.dll" (ByVal Id As Long, ByVal msg As String) As Long
Declare PtrSafe Function TmReceiveBin Lib "tmctl.dll" Alias "TmReceive" (ByVal Id As Long, ByRef buf As Any, ByVal blen As Long, ByRef rlen As Long) As Long
Declare PtrSafe Function TmReceiveSetup Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmReceiveBlockHeader Lib "tmctl.dll" (ByVal Id As Long, ByRef rlen As Long) As Long
Declare PtrSafe Function TmReceiveBlockData Lib "tmctl.dll" (ByVal Id As Long, ByRef buf As Any, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long) As Long
Declare PtrSafe Function TmCheckEnd Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSetRen Lib "tmctl.dll" (ByVal Id As Long, ByVal flg As Long) As Long
Declare PtrSafe Function TmSetTerm Lib "tmctl.dll" (ByVal Id As Long, ByVal eos As Long, ByVal eot As Long) As Long
Declare PtrSafe Function TmSetTimeout Lib "tmctl.dll" (ByVal Id As Long, ByVal tmo As Long) As Long
Declare PtrSafe Function TmGetStatusByte Lib "tmctl.dll" (ByVal Id As Long, ByRef sts As Byte) As Long
Declare PtrSafe Function TmFinish Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmGetLastError Lib "tmctl.dll" (ByVal Id As Long) As Long
Declare PtrSafe Function TmSearchDevices Lib "tmctl.dll" (ByVal wire As Long, list As DeviceListArray, ByVal max As Long, ByRef num As Long, ByVal option1 As String) As Long
Declare PtrSafe Function TmEncodeSerialNumber Lib "tmctl.dll" (ByVal encode As String, ByVal encodelen As Long, ByVal src As String) As Long
Declare PtrSafe Function TmDecodeSerialNumber Lib "tmctl.dll" (ByVal decode As String, ByVal decodelen As Long, ByVal src As String) As Long
#End If

Function TmReceive(ByVal Id As Long, ByRef buf As String, ByVal blen As Long, ByRef rlen As Long)
    TmReceive = TmReceiveBin(Id, ByVal buf, blen, rlen)
End Function

Function TmReceiveBlock(ByVal Id As Long, ByRef buf() As Byte, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long)
    TmReceiveBlock = TmReceiveBlockData(Id, buf(0), blen, rlen, ed)
End Function

Function TmReceiveBlockB(ByVal Id As Long, ByRef buf() As Byte, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long)
    TmReceiveBlockB = TmReceiveBlockData(Id, buf(0), blen, rlen, ed)
End Function

Function TmReceiveToFile(ByVal Id As Long, ByRef buf As String, ByRef rlen As Long)
    Dim buffer() As Byte
    Dim blen     As Long
    Dim rlen2    As Long
    Dim ed       As Long
    Dim ret      As Long

    Open buf For Binary Access Write As #1
    lWritePos = 0
    ret = TmReceiveBlockHeader(Id, rlen)
    Do While ret = 0
        ReDim buffer(65536)
        ret = TmReceiveBlock(Id, buffer, UBound(buffer), rlen2, ed)
        If rlen2 < UBound(buffer) Then
            ReDim Preserve buffer(rlen2)
        End If
        Put #1, , buffer
        rlen = rlen + UBound(buffer)
        If ret = 1 Or ed = 1 Then
            Exit Do
        End If
        ret = TmCheckEnd(Id)
    Loop
    Close #1
End Function
