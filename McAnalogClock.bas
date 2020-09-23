Attribute VB_Name = "McAnalogClock"
'Module name: McAnalogClock ver 1.0 (august 2000)
' Author: Miran Cvenkel
'Purpose: Put pictbox in your form and write this code somewhere
              'McAnalogClock PictBoxName
              'Clock will appear there and will be running
Dim days
Dim hours
Dim minutes
Dim seconds

Public Sub AnalogClock(mycontrol As Control)
Dim Diameter As Integer
Dim a As Long
Dim B As Long
Dim SW As Integer
Dim SH As Integer
Dim Color As Long
Dim Angle As Single


'adjust
mycontrol.Width = mycontrol.Height
'set mode
mycontrol.ScaleMode = 3
'set center point
SW = mycontrol.ScaleWidth / 2
SH = mycontrol.ScaleHeight / 2

'set diameter
Diameter = (mycontrol.ScaleHeight / 2) - 1

'make surronding circle
mycontrol.Circle (SW, SH), Diameter, QBColor(14)

'Draw dots
mycontrol.FillStyle = 0 'if u want filled dots
Dim d As Byte
For i = 0 To 6.3 Step 0.105
DoEvents
Angle = ((d - 30) * 0.105)
    angle1 = (6.3 - Angle)

     z1 = Sin(angle1 - 6.3) * (Diameter - 7)
    z2 = Cos(angle1 - 6.3) * (Diameter - 7)
    
    Select Case d
    Case 15, 30, 45, 60
    Case 59, 1, 0, 14, 16, 44, 46 'around numbers
    Case 5, 10, 20, 25, 35, 40, 50, 55
      mycontrol.FillColor = QBColor(14)
      mycontrol.Circle (z1 + SW, z2 + SH), 2, mycontrol.FillColor
    Case Else
       mycontrol.FillColor = QBColor(1)
      mycontrol.Circle (z1 + SW, z2 + SH), 1, mycontrol.FillColor
    End Select
d = d + 1
Next i
    
'print numbers
Call Print12(mycontrol, (mycontrol.ScaleWidth / 2) - 7.5, 4)
Call Print3(mycontrol, (mycontrol.ScaleWidth) - 12, (mycontrol.ScaleHeight / 2) - 5.5)
Call Print6(mycontrol, (mycontrol.ScaleWidth / 2) - 4, (mycontrol.ScaleHeight) - 14)
Call Print9(mycontrol, 4, (mycontrol.ScaleHeight / 2) - 5.5)
  
Do
DoEvents
For Angle = 0 To 6.3 Step 0.105
        DoEvents

'for seconds
MCTimeMeasurement (Timer) 'what is the time
Angle = ((seconds - 30) * 0.105) 'adjust it to actual seconds
    angle1 = (6.3 - Angle)
    a = Sin(angle1) * Diameter
    B = Cos(angle1) * Diameter
    a1 = Sin(angle1 - 6.3) * (Diameter - Diameter / 6 - 14)
    b1 = Cos(angle1 - 6.3) * (Diameter - Diameter / 6 - 14)
'for minutes

MCTimeMeasurement (Timer) 'what is the time
Angle = ((minutes - 30) * 0.105) 'adjust it to actual minutes
    angle1 = (6.3 - Angle)
    a = Sin(angle1) * Diameter
    B = Cos(angle1) * Diameter
    a2 = Sin(angle1 - 6.3) * (Diameter - Diameter / 7 - 14)
    b2 = Cos(angle1 - 6.3) * (Diameter - Diameter / 7 - 14)
'for hours
h = 6.3 / 24
MCTimeMeasurement (Timer) 'what is the time
Angle = ((((hours + (minutes / 60)) - 6) * 0.2625) * 2) 'adjust it to actual hours
    angle1 = (6.3 - Angle)
    a = Sin(angle1) * Diameter
    B = Cos(angle1) * Diameter
    a3 = Sin(angle1 - 6.3) * (Diameter - Diameter / 8 - 14)
    b3 = Cos(angle1 - 6.3) * (Diameter - Diameter / 8 - 14)


'remember stuff about seconds last position
rem1 = a1 + SW
rem2 = b1 + SH
'remember stuff about minutes last position
rem3 = a2 + SW
rem4 = b2 + SH
'remember stuff about hours last position
rem5 = a3 + SW
rem6 = b3 + SH

'draw lines showing time
mycontrol.Line (SW, SH)-(a3 + SW, b3 + SH), QBColor(13) 'hour
mycontrol.Line (SW, SH)-(a2 + SW, b2 + SH), QBColor(14) 'min
mycontrol.Line (SW, SH)-(a1 + SW, b1 + SH), QBColor(12) 'sec

'delay for 1 sec
Time1 = Timer
Do
DoEvents
Loop Until Timer > Time1 + 1
'-----------------
'erase lines
mycontrol.Line (SW, SH)-(rem1, rem2), mycontrol.BackColor
mycontrol.Line (SW, SH)-(rem3, rem4), mycontrol.BackColor
mycontrol.Line (SW, SH)-(rem5, rem6), mycontrol.BackColor

Next Angle
Loop

Rem --------------------------------------------------------------------------------

End Sub
Sub MCTimeMeasurement(Mresidium As Long)

' Author: Miran Cvenkel
'module Slightly modified from original.

'first reset all to 0
days = 0
hours = 0
minutes = 0
seconds = 0


MainEngine:
Select Case Mresidium ' (which contains seconds elapsed)
Case Is > 86400 'more then one day
GoTo Mdays
Case 3600 To 86400 ' between 1 hour and one day
GoTo Mhours
Case 60 To 3600 ' between 1 minute and one hour
GoTo Mminutes
Case Is < 60 'only seconds - not very likely-but this is used almost every second
GoTo Mseconds
Case Else
End Select

Mdays: 'count days elapsed
days = Int(Mresidium / 86400)
Mresidium = Mresidium - (days * 86400)
GoTo MainEngine

Mhours: 'count hours elapsed
hours = Int(Mresidium / 3600)
Mresidium = Mresidium - (hours * CLng(3600))
GoTo MainEngine

Mminutes: 'count minutes elapsed
minutes = Int(Mresidium / 60)
Mresidium = Mresidium - (minutes * 60)
GoTo MainEngine

Mseconds: 'count seconds elapsed
seconds = Mresidium


'now - compose appropriate string ( FusionTimeString )
'the kind that don't show 0 values - like - it doesn't show 0 sec.
Dim i As Integer
Dim j As Integer
Dim myArray(1, 3) As Variant

myArray(0, 0) = days
myArray(0, 1) = hours
myArray(0, 2) = minutes
myArray(0, 3) = seconds

myArray(1, 0) = ":"
myArray(1, 1) = ":"
myArray(1, 2) = ":"
myArray(1, 3) = ""

fusiontimestring = "" ' set it to zero lenght string

For j = 1 To 3
        'If myArray(i, j) <> 0 Then
        fusiontimestring = fusiontimestring & myArray(i, j) & " " & myArray(i + 1, j) & " "
        'End If
Next j

'Now FusionTimeString is generated and yust waiting to be displayed

End Sub


Sub Print12(TargetCtl As Control, X As Integer, Y As Integer)
Dim ActualTextWidth As Integer
Dim ActualTextHeight As Integer
ActualTextWidth = 15 'Pixels
ActualTextHeight = 11 'Pixels
TargetCtl.Line (X + 3, Y + 0)-(X + 5, Y + 0), 255
TargetCtl.Line (X + 10, Y + 0)-(X + 14, Y + 0), 255
TargetCtl.Line (X + 2, Y + 1)-(X + 5, Y + 1), 255
TargetCtl.Line (X + 9, Y + 1)-(X + 11, Y + 1), 255
TargetCtl.Line (X + 13, Y + 1)-(X + 15, Y + 1), 255
TargetCtl.Line (X + 0, Y + 2)-(X + 5, Y + 2), 255
TargetCtl.Line (X + 8, Y + 2)-(X + 10, Y + 2), 255
TargetCtl.Line (X + 14, Y + 2)-(X + 16, Y + 2), 255
TargetCtl.Line (X + 3, Y + 3)-(X + 5, Y + 3), 255
TargetCtl.Line (X + 8, Y + 3)-(X + 10, Y + 3), 255
TargetCtl.Line (X + 14, Y + 3)-(X + 16, Y + 3), 255
TargetCtl.Line (X + 3, Y + 4)-(X + 5, Y + 4), 255
TargetCtl.Line (X + 14, Y + 4)-(X + 16, Y + 4), 255
TargetCtl.Line (X + 3, Y + 5)-(X + 5, Y + 5), 255
TargetCtl.Line (X + 13, Y + 5)-(X + 15, Y + 5), 255
TargetCtl.Line (X + 3, Y + 6)-(X + 5, Y + 6), 255
TargetCtl.Line (X + 12, Y + 6)-(X + 14, Y + 6), 255
TargetCtl.Line (X + 3, Y + 7)-(X + 5, Y + 7), 255
TargetCtl.Line (X + 11, Y + 7)-(X + 13, Y + 7), 255
TargetCtl.Line (X + 3, Y + 8)-(X + 5, Y + 8), 255
TargetCtl.Line (X + 10, Y + 8)-(X + 12, Y + 8), 255
TargetCtl.Line (X + 3, Y + 9)-(X + 5, Y + 9), 255
TargetCtl.Line (X + 9, Y + 9)-(X + 11, Y + 9), 255
TargetCtl.Line (X + 3, Y + 10)-(X + 5, Y + 10), 255
TargetCtl.Line (X + 8, Y + 10)-(X + 10, Y + 10), 255
TargetCtl.Line (X + 3, Y + 11)-(X + 5, Y + 11), 255
TargetCtl.Line (X + 8, Y + 11)-(X + 16, Y + 11), 255
End Sub

Sub Print3(TargetCtl As Control, X As Integer, Y As Integer)
Dim ActualTextWidth As Integer
Dim ActualTextHeight As Integer
ActualTextWidth = 8 'Pixels
ActualTextHeight = 11 'Pixels
TargetCtl.Line (X + 2, Y + 0)-(X + 7, Y + 0), 255
TargetCtl.Line (X + 1, Y + 1)-(X + 4, Y + 1), 255
TargetCtl.Line (X + 5, Y + 1)-(X + 8, Y + 1), 255
TargetCtl.Line (X + 0, Y + 2)-(X + 3, Y + 2), 255
TargetCtl.Line (X + 6, Y + 2)-(X + 9, Y + 2), 255
TargetCtl.Line (X + 0, Y + 3)-(X + 3, Y + 3), 255
TargetCtl.Line (X + 6, Y + 3)-(X + 9, Y + 3), 255
TargetCtl.Line (X + 5, Y + 4)-(X + 8, Y + 4), 255
TargetCtl.Line (X + 3, Y + 5)-(X + 7, Y + 5), 255
TargetCtl.Line (X + 5, Y + 6)-(X + 8, Y + 6), 255
TargetCtl.Line (X + 6, Y + 7)-(X + 9, Y + 7), 255
TargetCtl.Line (X + 0, Y + 8)-(X + 3, Y + 8), 255
TargetCtl.Line (X + 6, Y + 8)-(X + 9, Y + 8), 255
TargetCtl.Line (X + 0, Y + 9)-(X + 3, Y + 9), 255
TargetCtl.Line (X + 6, Y + 9)-(X + 9, Y + 9), 255
TargetCtl.Line (X + 1, Y + 10)-(X + 4, Y + 10), 255
TargetCtl.Line (X + 5, Y + 10)-(X + 8, Y + 10), 255
TargetCtl.Line (X + 2, Y + 11)-(X + 7, Y + 11), 255
End Sub
Sub Print6(TargetCtl As Control, X As Integer, Y As Integer)

Dim ActualTextWidth As Integer
Dim ActualTextHeight As Integer
ActualTextWidth = 8 'Pixels
ActualTextHeight = 11 'Pixels
TargetCtl.Line (X + 2, Y + 0)-(X + 8, Y + 0), 255
TargetCtl.Line (X + 1, Y + 1)-(X + 4, Y + 1), 255
TargetCtl.Line (X + 6, Y + 1)-(X + 9, Y + 1), 255
TargetCtl.Line (X + 1, Y + 2)-(X + 3, Y + 2), 255
TargetCtl.Line (X + 7, Y + 2)-(X + 9, Y + 2), 255
TargetCtl.Line (X + 0, Y + 3)-(X + 3, Y + 3), 255
TargetCtl.Line (X + 0, Y + 4)-(X + 7, Y + 4), 255
TargetCtl.Line (X + 0, Y + 5)-(X + 4, Y + 5), 255
TargetCtl.Line (X + 5, Y + 5)-(X + 8, Y + 5), 255
TargetCtl.Line (X + 0, Y + 6)-(X + 3, Y + 6), 255
TargetCtl.Line (X + 6, Y + 6)-(X + 9, Y + 6), 255
TargetCtl.Line (X + 0, Y + 7)-(X + 3, Y + 7), 255
TargetCtl.Line (X + 6, Y + 7)-(X + 9, Y + 7), 255
TargetCtl.Line (X + 0, Y + 8)-(X + 3, Y + 8), 255
TargetCtl.Line (X + 6, Y + 8)-(X + 9, Y + 8), 255
TargetCtl.Line (X + 0, Y + 9)-(X + 3, Y + 9), 255
TargetCtl.Line (X + 6, Y + 9)-(X + 9, Y + 9), 255
TargetCtl.Line (X + 1, Y + 10)-(X + 4, Y + 10), 255
TargetCtl.Line (X + 5, Y + 10)-(X + 8, Y + 10), 255
TargetCtl.Line (X + 2, Y + 11)-(X + 7, Y + 11), 255
End Sub
Sub Print9(TargetCtl As Control, X As Integer, Y As Integer)
Dim ActualTextWidth As Integer
Dim ActualTextHeight As Integer
ActualTextWidth = 8 'Pixels
ActualTextHeight = 11 'Pixels
TargetCtl.Line (X + 2, Y + 0)-(X + 7, Y + 0), 255
TargetCtl.Line (X + 1, Y + 1)-(X + 4, Y + 1), 255
TargetCtl.Line (X + 5, Y + 1)-(X + 8, Y + 1), 255
TargetCtl.Line (X + 0, Y + 2)-(X + 3, Y + 2), 255
TargetCtl.Line (X + 6, Y + 2)-(X + 9, Y + 2), 255
TargetCtl.Line (X + 0, Y + 3)-(X + 3, Y + 3), 255
TargetCtl.Line (X + 6, Y + 3)-(X + 9, Y + 3), 255
TargetCtl.Line (X + 0, Y + 4)-(X + 3, Y + 4), 255
TargetCtl.Line (X + 6, Y + 4)-(X + 9, Y + 4), 255
TargetCtl.Line (X + 0, Y + 5)-(X + 3, Y + 5), 255
TargetCtl.Line (X + 6, Y + 5)-(X + 9, Y + 5), 255
TargetCtl.Line (X + 1, Y + 6)-(X + 4, Y + 6), 255
TargetCtl.Line (X + 5, Y + 6)-(X + 9, Y + 6), 255
TargetCtl.Line (X + 2, Y + 7)-(X + 9, Y + 7), 255
TargetCtl.Line (X + 6, Y + 8)-(X + 9, Y + 8), 255
TargetCtl.Line (X + 0, Y + 9)-(X + 2, Y + 9), 255
TargetCtl.Line (X + 6, Y + 9)-(X + 8, Y + 9), 255
TargetCtl.Line (X + 0, Y + 10)-(X + 3, Y + 10), 255
TargetCtl.Line (X + 5, Y + 10)-(X + 8, Y + 10), 255
TargetCtl.Line (X + 1, Y + 11)-(X + 7, Y + 11), 255
End Sub
