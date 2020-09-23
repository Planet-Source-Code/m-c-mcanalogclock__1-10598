VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5040
      Width           =   8415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Start cool stuff"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   5040
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      Visible         =   0   'False
      X1              =   8
      X2              =   160
      Y1              =   80
      Y2              =   16
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   0
      X2              =   152
      Y1              =   64
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim Course As String
Dim X
Dim Y
Dim a
Dim B
Dim back
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Private Sub Command1_Click()

Dim Diameter As Integer
Dim MyColor As Integer
Dim a As Long
Dim B As Long
Dim SW As Integer
Dim SH As Integer
Dim Color As Long
Dim Angle As Single

SW = Me.ScaleWidth / 2
SH = Me.ScaleHeight / 2
Cls
Diameter = 150
MyColor = 255
Color = RGB(MyColor, MyColor, 0)
'make egg
Do
   For Angle = 0 To 6.3 Step 0.01
      a = Sin(Angle) * Diameter
      B = Cos(Angle) * Diameter
      a1 = Sin(20 - Angle) * Diameter
        b1 = Cos(20 - Angle) * Diameter
'SetPixel Me.hdc, A + SW, B + SH, RGB(MyColor, MyColor, 0)
'SetPixel Me.hdc, B + SW, A + SH, RGB(MyColor, MyColor, 0)
      
Line1.X1 = a + SW
Line1.X2 = B + SW
Line1.Y1 = b1 + SH
Line1.Y2 = a1 + SH

Line2.X1 = a + SW
Line2.X2 = B + SW
Line2.Y1 = Sin(B) + (SH)
Line2.Y2 = a + SH


Line3.X1 = Line1.X1
Line3.X2 = Line2.X2
Line3.Y1 = Line1.Y1
Line3.Y2 = Line2.Y2
DoEvents
      'SetPixel Me.hdc, A + SW, B + SH, Color
For i = 0 To 10000
Next i
      
   Next Angle
   Diameter = Diameter + 1
   'MyColor = MyColor - MyColor / Diameter
Loop
End Sub

Private Sub Command2_Click()
End
End Sub








Private Sub Command6_Click()
MsgBox " this will run out of memory" & Chr(10) & _
"i.e. clocks will dissappear" & Chr(10) & _
"as a result of using transparency" & Chr(10) & _
"Same thing happpens if you use Msimg32.dll" & Chr(10) & _
"Any ideas how to avoid this ?"

Timer1.Enabled = True
AnalogClock Picture1
Beep 'this beep want happen ever as it follows above line
End Sub

Private Sub Form_Load()
Course = "downright"
End Sub

Private Sub Timer1_Timer()
 Dim MemStat As MEMORYSTATUS
    'retrieve the memory status
    GlobalMemoryStatus MemStat
    Text1.Text = "You have" + Str$(MemStat.dwTotalPhys / 1024) + " Kb total memory and" + Str$(MemStat.dwAvailPageFile / 1024) + " Kb available PageFile memory."
Form1.Cls



Select Case Course
Case "downright"
X = X + 2
Y = Y + 2
If X > Screen.Width / Screen.TwipsPerPixelX - Picture1.Width Then Course = "downleft"
If Y > Screen.Height / Screen.TwipsPerPixelY - Picture1.Height Then Course = "upright"
'BitBlt Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY
Call TransBitBlt(Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0)
Case "upright"
X = X + 2
Y = Y - 2
If X > Screen.Width / Screen.TwipsPerPixelX - Picture1.Width Then Course = "upleft"
If Y < 0 Then Course = "downright"
'BitBlt Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY
Call TransBitBlt(Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0)
Case "upleft"
X = X - 2
Y = Y - 2
If X < 0 Then Course = "upright"
If Y < 0 Then Course = "downleft"
'BitBlt Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY
Call TransBitBlt(Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0)
Case "downleft"
X = X - 2
Y = Y + 2
If X < 0 Then Course = "downright"
If Y > Screen.Height / Screen.TwipsPerPixelY - Picture1.Height Then Course = "upleft"
'BitBlt Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY
Call TransBitBlt(Me.hdc, X, Y, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0)
Case Else
End Select

'BitBlt Me.hdc, X + 90, Y + 90, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY

Call TransBitBlt(Me.hdc, X + 90, Y + 90, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0)


End Sub
