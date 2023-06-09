Attribute VB_Name = "tmctl"
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

Public Const MaxStationNum = 128

Type DeviceList
        adr As String * 64
End Type

Type DeviceListArray
    list(MaxStationNum - 1) As DeviceList
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

Function TmReceiveBlock(ByVal Id As Long, ByRef buf() As Integer, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long)
    TmReceiveBlock = TmReceiveBlockData(Id, buf(0), blen, rlen, ed)
End Function

Function TmReceiveBlockB(ByVal Id As Long, ByRef buf() As Byte, ByVal blen As Long, ByRef rlen As Long, ByRef ed As Long)
    TmReceiveBlockB = TmReceiveBlockData(Id, buf(0), blen, rlen, ed)
End Function
