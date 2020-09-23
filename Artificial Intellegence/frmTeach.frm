VERSION 5.00
Begin VB.Form frmTeach 
   Caption         =   "Teach"
   ClientHeight    =   1095
   ClientLeft      =   4275
   ClientTop       =   4965
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4710
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Teach 
      Caption         =   "&Enter"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TeachText 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmTeach.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a character that must be taught"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmTeach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Teach_Click()
    frmMain.Main_TeachText = TeachText.Text
    Unload Me
End Sub
