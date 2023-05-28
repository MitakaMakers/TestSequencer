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
        MilliSleep (10)
        Cells(row, cnLo.addressColumn).Select
        MilliSleep (10)
        Cells(row, cnLo.timeoutColumn).Select
        MilliSleep (10)
        Cells(row, cnLo.statusColumn).Select
        MilliSleep (10)
    Next row
    
    For row = cmdLo.startRow To cmdLo.endRow
        Cells(row, cmdLo.deviceColumn).Select
        MilliSleep (excOpt.interval)
        Cells(row, cmdLo.commandColumn).Select
        MilliSleep (10)
        Cells(row, cmdLo.responseColumn).Select
        MilliSleep (10)
        Cells(row, cmdLo.statusColumn).Select
    Next row
    
    bkupSel.Select
End Sub
