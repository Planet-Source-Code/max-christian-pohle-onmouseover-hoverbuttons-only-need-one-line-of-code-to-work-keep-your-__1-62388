VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Easy Hoover-Buttons by www.coderonline.de"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   90
      TabIndex        =   7
      Top             =   3420
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   2970
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000004&
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   2070
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      ForeColor       =   &H80000007&
      Height          =   375
      Index           =   3
      Left            =   90
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      ForeColor       =   &H80000007&
      Height          =   375
      Index           =   2
      Left            =   90
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   990
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      ForeColor       =   &H80000007&
      Height          =   375
      Index           =   1
      Left            =   90
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   540
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000004&
      Caption         =   "Check1"
      ForeColor       =   &H80000007&
      Height          =   375
      Index           =   0
      Left            =   90
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   90
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is all you need for the hoover-effect...

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HoverButton.MakeHover Check1(Index)
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HoverButton.MakeHover Command1(Index)
End Sub






'this is just a hint for you...

Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Check1(Index).Value = 0 Then Exit Sub
    Check1(Index).Value = 0
    
    Debug.Print "Check1(" & Index & ") was clicked"
    'Now you can use Check-Boxes the same way like command-buttons
    'because they have more properties (as ForeColor) they look better
End Sub

Private Sub Command1_Click(Index As Integer)
    Debug.Print "Command1(" & Index & ") was clicked"
End Sub

