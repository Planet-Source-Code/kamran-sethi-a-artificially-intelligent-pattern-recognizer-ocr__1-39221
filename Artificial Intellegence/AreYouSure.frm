VERSION 5.00
Begin VB.Form frmConfirmation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmation"
   ClientHeight    =   1020
   ClientLeft      =   4485
   ClientTop       =   4260
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NOButton 
      Caption         =   "&No"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton YESButton 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      Picture         =   "AreYouSure.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Exit character recognition?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public YES As Boolean

'Private Sub Form_Load()
'    Me.Left = (frmMain.Width / 2) - (Me.Width / 2)
'    Me.Top = (frmMain.ScaleHeight / 2) - (Me.Height)
'End Sub

Private Sub NOButton_Click()
    Unload Me
End Sub

Private Sub YESButton_Click()
    YES = True
    Unload Me
End Sub
