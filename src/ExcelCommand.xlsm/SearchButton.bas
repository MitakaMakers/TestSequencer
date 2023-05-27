Attribute VB_Name = "SearchButton"
Option Explicit

Sub SearchButton_Click()
    Dim Ret As Long
    Dim Id  As Long
    
    AddDllDirectories(ThisWorkbook.Path)
    
    Id = 1
    Ret = TmInitialize(1, "a", 1)
    Ret = TmFinish(Id)
    Ret = TmSetTimeout(Id, 100)
    Ret = TmSetTerm(Id, 1, 2)
    Ret = TmSend(Id, "*IDN?")
    Ret = TmCheckEnd(Id)
    Ret = TmGetLastError(Id)
    Ret = TmSetRen(Id, 1)
    Ret = TmDeviceClear(Id)
End Sub
