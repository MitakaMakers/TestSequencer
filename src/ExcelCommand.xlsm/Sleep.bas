Attribute VB_Name = "Sleep"
Option Explicit

Function MilliSleep(Time As Long)
    Application.Wait [Now()] + Time / 86400000
End Function
