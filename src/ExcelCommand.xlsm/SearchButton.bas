Attribute VB_Name = "SearchButton"
Option Explicit

Sub SearchButton_Click()
    Dim bkupSel As Range
    Dim cnLo    As ConnectLayout
    Dim ret     As Long
    Dim Id      As Long
    Dim row     As Long
    
    Set bkupSel = Selection
    
    AddDllDirectories (ThisWorkbook.Path)
    
    cnLo = GetCnLayout()
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
    
    bkupSel.Select
End Sub
