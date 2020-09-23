VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Basic example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form2.frx":0000
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go to the advanced example of McAnalogClock"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start basic example of McAnalogClock"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   3720
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AnalogClock Picture1
End Sub

Private Sub Command2_Click()
Unload Form2
Form1.Show
End Sub

