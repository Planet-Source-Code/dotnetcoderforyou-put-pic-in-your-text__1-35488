VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   750
      Left            =   315
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   660
      TabIndex        =   1
      Top             =   360
      Width           =   720
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2925
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":2DF3
      Top             =   1680
      Width           =   6750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdShow_Click()

End Sub

Private Sub Text1_GotFocus()
h& = GetFocus&()
b& = Picture1.Picture
Call CreateCaret(h&, b&, 10, 10)
x& = ShowCaret&(h&)
End Sub

