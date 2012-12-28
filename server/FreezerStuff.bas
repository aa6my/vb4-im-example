Attribute VB_Name = "FreezerStuff"
Option Explicit
Public exitPause As Boolean


Public Function timedPause(secs As Long)
    Dim secStart As Variant
    Dim secNow As Variant
    Dim secDiff As Variant
    Dim Temp%
    
    exitPause = False 'this is our early way out out of the pause
    
    secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds
    


    Do While secDiff < secs
        If exitPause = True Then Exit Do
        secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and Date at any itteration of the Loop
        secDiff = DateDiff("s", secStart, secNow) 'this compares the start time With the current time
        Temp% = DoEvents
    Loop
End Function
