Attribute VB_Name = "Sleep"
Option Explicit

Function MSleep(Time As Long)
    Application.Wait [Now()] + Time / 86400000
End Function
