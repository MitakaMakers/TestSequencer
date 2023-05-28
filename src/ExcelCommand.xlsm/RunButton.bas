Attribute VB_Name = "RunButton"
Option Explicit

Sub RunButton_Click()
    Dim bkupSel As Range
    Dim cmdLo   As CommandLayout
    Dim excOpt  As ExecOption
    Dim cnLo    As ConnectLayout
    Dim sheet   As Worksheet
    
    Set bkupSel = Selection
    
    cmdLo = GetCmdLayout()
    excOpt = GetExecOption()
    cnLo = GetCnLayout()
    
    Dim row As Long
    For row = cnLo.startRow To cnLo.endRow
        Cells(row, cnLo.wireColumn).Select
        Sleep (10)
        Cells(row, cnLo.addressColumn).Select
        Sleep (10)
        Cells(row, cnLo.timeoutColumn).Select
        Sleep (10)
        Cells(row, cnLo.statusColumn).Select
        Sleep (10)
    Next row
    
    For row = cmdLo.startRow To cmdLo.endRow
        Cells(row, cmdLo.deviceColumn).Select
        Sleep (excOpt.interval)
        Cells(row, cmdLo.commandColumn).Select
        Sleep (10)
        Cells(row, cmdLo.responseColumn).Select
        Sleep (10)
        Cells(row, cmdLo.statusColumn).Select
    Next row
    
    bkupSel.Select
End Sub

Function Sleep(Time As Long)
    Application.Wait [Now()] + Time / 86400000
End Function
