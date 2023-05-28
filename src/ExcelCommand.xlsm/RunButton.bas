Attribute VB_Name = "RunButton"
Option Explicit

Sub RunButton_Click()
    Dim bkupSht As Worksheet
    Dim bkupSel As Range
    Dim cmdLo   As CommandLayout
    Dim excOpt  As ExecOption
    Dim cnLo    As ConnectLayout
    Dim sheet   As Worksheet
    
    Set bkupSht = ActiveSheet
    Set bkupSel = Selection
    
    cmdLo = GetCmdLayout()
    excOpt = GetExecOption()
    cnLo = GetCnLayout()
    
    Set sheet = ThisWorkbook.Worksheets(cnLo.sheetName)
    sheet.Activate
    
    Dim row As Long
    For row = cnLo.startRow To cnLo.endRow
        Cells(row, cnLo.wireColumn).Select
        MilliSleep (20)
        Cells(row, cnLo.addressColumn).Select
        MilliSleep (20)
        Cells(row, cnLo.timeoutColumn).Select
        MilliSleep (20)
        Cells(row, cnLo.statusColumn).Select
        MilliSleep (20)
    Next row
    
    Set sheet = ThisWorkbook.Worksheets(cmdLo.sheetName)
    sheet.Activate
    
    For row = cmdLo.startRow To cmdLo.endRow
        Cells(row, cmdLo.deviceColumn).Select
        MilliSleep (20)
        Cells(row, cmdLo.commandColumn).Select
        MilliSleep (20)
        Cells(row, cmdLo.responseColumn).Select
        MilliSleep (20)
        Cells(row, cmdLo.statusColumn).Select
        MilliSleep (20)
    Next row
    
    bkupSht.Activate
    bkupSel.Select
End Sub
