Attribute VB_Name = "ConfigSheet"
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

Const cfgSheet As String = "Config"

Type ExecOption
    timeout        As Long
    interval       As Long
End Type

Type ConnectLayout
    startRow       As Long
    endRow         As Long
    wireColumn     As Long
    addressColumn  As Long
    termColumn     As Long
    statusColumn   As Long
End Type

Type CommandLayout
    startRow       As Long
    endRow         As Long
    opColumn       As Long
    arg1Column     As Long
    arg2Column     As Long
    resultColumn   As Long
    statusColumn   As Long
End Type

Type LanguageLayout
    sheetname      As String
    startRow       As Long
    endRow         As Long
    sheetColumn    As Long
    rowColumn      As Long
    columnColumn   As Long
    japaneseColumn As Long
    englishColumn  As Long
    chineseColumn As Long
    koreanColumn  As Long
End Type

Function GetExecOption() As ExecOption
    Dim sheet As Worksheet
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = cfgSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & cfgSheet & "]シートが見つかりません", vbInformation
        Exit Function
    End If
    
    GetExecOption.timeout = CLng(sheet.Range("D5").value)
    GetExecOption.interval = CLng(sheet.Range("D6").value)
End Function

Function GetCnLayout() As ConnectLayout
    Dim sheet As Worksheet
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = cfgSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & cfgSheet & "]シートが見つかりません", vbInformation
        Exit Function
    End If
    
    GetCnLayout.startRow = CLng(sheet.Range("D10").value)
    GetCnLayout.endRow = CLng(sheet.Range("D11").value)
    GetCnLayout.wireColumn = CLng(sheet.Range("D12").value)
    GetCnLayout.addressColumn = CLng(sheet.Range("D13").value)
    GetCnLayout.termColumn = CLng(sheet.Range("D14").value)
    GetCnLayout.statusColumn = CLng(sheet.Range("D15").value)
End Function

Function GetCmdLayout() As CommandLayout
    Dim sheet As Worksheet
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = cfgSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & cfgSheet & "]シートが見つかりません", vbInformation
        Exit Function
    End If
    
    GetCmdLayout.startRow = CLng(sheet.Range("D19").value)
    GetCmdLayout.endRow = CLng(sheet.Range("D20").value)
    GetCmdLayout.opColumn = CLng(sheet.Range("D21").value)
    GetCmdLayout.arg1Column = CLng(sheet.Range("D22").value)
    GetCmdLayout.arg2Column = CLng(sheet.Range("D23").value)
    GetCmdLayout.resultColumn = CLng(sheet.Range("D24").value)
    GetCmdLayout.statusColumn = CLng(sheet.Range("D25").value)
End Function

Function GetLangLayout() As LanguageLayout
    Dim sheet As Worksheet
    
    For Each sheet In Application.ThisWorkbook.Worksheets
        If sheet.name = cfgSheet Then
            Exit For
        End If
    Next sheet
    
    If sheet Is Nothing Then
        MsgBox "[" & cfgSheet & "]シートが見つかりません", vbInformation
        Exit Function
    End If
    
    GetLangLayout.sheetname = CStr(sheet.Range("D29").value)
    GetLangLayout.startRow = CLng(sheet.Range("D30").value)
    GetLangLayout.endRow = CLng(sheet.Range("D31").value)
    GetLangLayout.sheetColumn = CLng(sheet.Range("D32").value)
    GetLangLayout.rowColumn = CLng(sheet.Range("D33").value)
    GetLangLayout.columnColumn = CLng(sheet.Range("D34").value)
    GetLangLayout.japaneseColumn = CLng(sheet.Range("D35").value)
    GetLangLayout.englishColumn = CLng(sheet.Range("D36").value)
    GetLangLayout.chineseColumn = CLng(sheet.Range("D37").value)
    GetLangLayout.koreanColumn = CLng(sheet.Range("D38").value)
End Function
