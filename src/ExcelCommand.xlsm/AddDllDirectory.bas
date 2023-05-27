Attribute VB_Name = "AddDllDirectory"
Option Explicit

Private Declare PtrSafe Function SetDefaultDllDirectories Lib "kernel32.dll" (ByVal DirectoryFlags As Long) As Long
Private Declare PtrSafe Function AddDllDirectory Lib "kernel32.dll" (ByVal fileName As String) As LongPtr

Const LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = &H1000 'ÅF0x00001000

Sub AddDllDirectories(DLLFoldPath As String)
    SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS)
    AddDllDirectory(StrConv(DLLFoldPath, vbUnicode))
End Sub
