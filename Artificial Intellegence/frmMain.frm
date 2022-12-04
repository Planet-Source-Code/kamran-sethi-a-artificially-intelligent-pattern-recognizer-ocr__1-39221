VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Recognition"
   ClientHeight    =   5055
   ClientLeft      =   1755
   ClientTop       =   1410
   ClientWidth     =   8730
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C0FFFF&
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000018&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0442
   ScaleHeight     =   5055
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar progressRecognition 
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton buttonLearnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton buttonLearnConfirm 
      Caption         =   "Enter"
      Height          =   375
      Left            =   360
      Picture         =   "frmMain.frx":074C
      TabIndex        =   6
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox inputLearnCharacter 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton buttonLearn 
      Caption         =   "Learn"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton buttonRecognise 
      Caption         =   "&Recognise"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton buttonClearScreen 
      Caption         =   "&Clear Screen"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox AppTemplateArea 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3120
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.PictureBox userTemplateArea 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6000
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame frameArea1 
      Caption         =   "User's Draw Area"
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2775
      Begin VB.PictureBox userDrawArea 
         BackColor       =   &H80000016&
         DrawStyle       =   2  'Dot
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2595
         ScaleWidth      =   2475
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame frameArea3 
      Caption         =   "Buffer Data Area"
      Height          =   3495
      Left            =   5880
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame frameArea2 
      Caption         =   "Database Area"
      Height          =   3495
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label TeachLabelText 
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
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label textResult 
      Caption         =   "Character Recognition Ready..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp_About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuPopUp_Close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Main_inputLearnCharacter As String

Dim strCaption As String
Dim recordFileExtension As String
Dim currentlyDrawing As Boolean
Dim c As Integer
Dim strData As String
Dim strRECpk As String
Dim arrRawData(100 * 100) As String
Dim arrTagData(100 * 100) As String




Private Sub buttonLearnCancel_Click()
    
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.buttonLearnConfirm.Visible = False
    Me.buttonLearnCancel.Visible = False
    Me.inputLearnCharacter.Visible = False
    
  
    Me.buttonLearn.Visible = True
    Me.buttonRecognise.Visible = True
    Me.buttonClearScreen.Visible = True
  
    
End Sub

Private Sub buttonClearScreen_Click()

    strData = ""

    Me.textResult.Caption = ""
    Me.textResult.ToolTipText = ""
    Me.Refresh

 '  Clear all the drawing windows....
    userDrawArea.Cls
    AppTemplateArea.Cls
    userTemplateArea.Cls
   
'   Enable all buttons
    Me.buttonLearn.Enabled = False
    Me.buttonRecognise.Enabled = False
    Me.buttonClearScreen.Enabled = False
    
End Sub


Private Sub buttonLearnConfirm_Click()

Dim Filename_Database As String
Dim Filename_buttonLearn As String
Dim Buffer_DrawArea As Variant
Dim strbuttonLearnText As String
Dim intCounter As Integer
Dim strBuffer As String
'Dim buttonLearnDialog As New frmbuttonLearn
'Dim oFile As TextStream

'    Set oFile = New TextStream
    
    FileSystem.ChDir (App.Path)
    userTemplateArea.Cls
    strbuttonLearnText = Me.inputLearnCharacter.Text
    
    intCounter = 0
    
    Filename_Database = "DATA" & recordFileExtension
    Filename_buttonLearn = Filename_Database
  
    intCounter = intCounter + 1
    
    Call GraspRawData
    
    If strData = "" Then
        Call buttonClearScreen_Click
        Me.buttonLearn.Enabled = Not Me.buttonLearn.Enabled
        MsgBox "Detect No character was drawn in the Draw Area. buttonLearn operation can not be proceed. ", vbExclamation, "Warning..."
        GoTo buttonLearnConfirm_SkipbuttonLearn
    End If

    Open Filename_buttonLearn For Binary As #1
        strBuffer = Space(5)
        Get #1, , strBuffer
    Close #1
    
    If strBuffer = "recPK" Then
' add buttonLearning character as binary
'**
        strRECpk = ""
        strBuffer = ""
        Open Filename_buttonLearn For Binary As #1
            strBuffer = Space(5)
            Get #1, , strBuffer
            strRECpk = strRECpk & strBuffer
            strBuffer = Space(22)
            While Not EOF(1)
                Get #1, , strBuffer
                strRECpk = strRECpk & strBuffer
            Wend
        Close #1
        strRECpk = Mid(strRECpk, 1, Len(strRECpk) - 22)
        i = 3
        strRECpk = strRECpk & strbuttonLearnText
        strRECpk = strRECpk & ","
        For j = 1 To 10
            strRECpk = strRECpk & Chr(BinToDec(Mid(strData, i - 2 + ((j - 1) * 10), 2)))
            strRECpk = strRECpk & Chr(BinToDec(Mid(strData, i + 0 + ((j - 1) * 10), 8)))
        Next j
        Open Filename_buttonLearn For Binary As #1
            Put #1, , strRECpk
        Close #1

    Else
' add buttonLearning character as string

        Open Filename_buttonLearn For Append As #1
            Write #1, strbuttonLearnText & "," & strData
        Close #1
'**
    End If
    
buttonLearnConfirm_SkipbuttonLearn:
    
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.buttonLearnConfirm.Visible = False
    Me.buttonLearnCancel.Visible = False
    Me.inputLearnCharacter.Visible = False
    
    Me.buttonLearn.Enabled = False

 
    Me.buttonLearn.Visible = True
    Me.buttonRecognise.Visible = True
    Me.buttonClearScreen.Visible = True
   

End Sub

Private Sub Image1_Click()
 frmSplash.Show
 frmSplash.Refresh
End Sub

Private Sub Label1_Click()
frmAbout.Show
frmAbout.Refresh
End Sub





Private Sub userDrawArea_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    currentlyDrawing = True
    
    userDrawArea.DrawWidth = 17
    userDrawArea.PSet (x, y)

    If Not Me.buttonLearn.Enabled Then
     
        Me.buttonLearn.Enabled = True
        Me.buttonRecognise.Enabled = True
        Me.buttonClearScreen.Enabled = True
    End If
    
    If Not Me.buttonRecognise.Enabled Then
        Me.buttonRecognise.Enabled = True
      
    End If

End Sub

Private Sub userDrawArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    userDrawArea.DrawStyle = vbSolid
    userDrawArea.DrawWidth = 17
    If currentlyDrawing Then
        userDrawArea.PSet (x, y)
        
        If Not fMainForm.buttonLearn.Enabled Then
            Me.buttonLearn.Enabled = True
            Me.buttonRecognise.Enabled = True
            Me.buttonClearScreen.Enabled = True
        End If
    End If
End Sub

Private Sub userDrawArea_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    currentlyDrawing = False
    
    Call GraspRawData
End Sub

Private Sub Exit_Click()

'    MsgBox "Are you sure want to Exit?", vbYesNo, "Confirmation"
'    If vbYes Then
'      Unload Me
'    End If

'        frmConfirmation.Left = (frmMain.Width / 2) - (frmConfirmation.Width / 2)
'        frmConfirmation.Top = (frmMain.ScaleHeight / 2) - (frmConfirmation.Height)

        frmConfirmation.Show vbModal
        If frmConfirmation.YES Then
            Unload Me
        End If
End Sub

Private Sub Form_Load()
  
    recordFileExtension = ".rec"
    strCaption = "The software is ready for processing"
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.userTemplateArea.DrawWidth = 2
    Me.AppTemplateArea.DrawWidth = 2
    
    Me.buttonLearn.Enabled = False
    Me.buttonRecognise.Enabled = False
    Me.buttonClearScreen.Enabled = False
    
    Me.inputLearnCharacter.Visible = False
    Me.buttonLearnConfirm.Visible = False
    Me.buttonLearnCancel.Visible = False
    
    
    Me.progressRecognition.Value = 0
    
   
    
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuPopUp
    End If
End Sub



Private Sub mnuPopUp_About_Click()
Dim AboutDialog As New frmAbout
    AboutDialog.Show vbModal
End Sub

Private Sub mnuPopUp_Close_Click()
    Call Exit_Click
End Sub



Private Sub buttonRecognise_Click()
Dim Filename_Database As String
Dim strbuttonRecognised As String
Dim intMatch As Integer
Dim intMaxMatch As Integer
Dim intCounter As Integer
Dim boolFindLastFile As Boolean
Dim boolNoMoreFileLeft As Boolean
Dim buffer As String
Dim Buffer_DatabaseArea As Variant

Dim strBuffer As String

   FileSystem.ChDir (App.Path)
    
    
    
    strbuttonRecognised = ""
    intMaxMatch = 0
    intCounter = 0
    c = 0
    
    On Error GoTo buttonRecognise_FileClose
    Filename_Database = "DATA" & recordFileExtension
    
    If Filename_Database <> "" Then
        Me.AppTemplateArea.Cls
        Open Filename_Database For Binary As #1
            i = 0
            strBuffer = Space(5)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            strBuffer = Space(22)
            While Not EOF(1)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            Wend
            arrRawData(i - 1) = ""
            Close #1
            
            i = 0
            strBuffer = ""
            If arrRawData(0) = "recPK" Then
                i = i + 1
                
                strbuttonRecognised = ""
                intMaxMatch = 0
                intCounter = 0
                c = 0
                
                While arrRawData(i) <> ""
                
                    a = 190
                    b = 190
                    d = 0
                    intMatch = 0
                    Me.AppTemplateArea.Cls
                
                    arrTagData(i - 1) = Mid(arrRawData(i), 1, 1)
'                    strBuffer = strBuffer & Mid(strData(i), 1, 1) & vbCrLf
'                    strBuffer = strBuffer & Mid(strData(0), i, 1) & ","
'                    Me.ListBox_List.AddItem i & ". <" & Mid(strData(i), 1, 1) & ">"
                    
                    arrRawData(i - 1) = ""
                    
                    For j = 1 To 10
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 3 + ((j - 1) * 2), 1)), 2)
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 4 + ((j - 1) * 2), 1)), 8)
                        arrRawData(i - 1) = arrRawData(i - 1) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 3 + ((j - 1) * 2), 1)), 2) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 4 + ((j - 1) * 2), 1)), 8)
'                        strBuffer = strBuffer
                    Next j
                    
'                    strBuffer = strBuffer & vbCrLf
                    i = i + 1

                    For ii = 1 To 10
                    For jj = 1 To 10
                        If Mid(arrRawData(c), d + 1, 1) = vbBlack Then
        '                    userTemplateArea.PSet (i, j)
' j2 - mask - finalise
'                            AppTemplateArea.PSet (a, b)

        '                    userTemplateArea.Circle (a, b), 110
        '                    AppTemplateArea.Line (a - 110, b - 110)-(a + 110, b - 110)
        '                    AppTemplateArea.Line (a + 110, b - 110)-(a + 110, b + 110)
        '                    AppTemplateArea.Line (a + 110, b + 110)-(a - 110, b + 110)
        '                    AppTemplateArea.Line (a - 110, b + 110)-(a - 110, b - 110)
        '                    Debug.Print ""
                            If Mid(strData, d + 1, 1) = vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        Else
                            If Mid(strData, d + 1, 1) <> vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        End If
                        d = d + 1
                        b = b + (Me.AppTemplateArea.Height - 200) / 10
                    Next jj
                    b = 190
                    a = a + (Me.AppTemplateArea.Width - 200) / 10
                    Next ii
                    If intMaxMatch < intMatch Then
                        intMaxMatch = intMatch
                        strbuttonRecognised = arrTagData(c)
                        intCounter = c
                        Me.progressRecognition.Value = intMaxMatch
                        If intMaxMatch > 90 Then
                            GoTo buttonRecognise_FileClose
                        End If
                    End If
                    c = c + 1
                
                Wend
'                Me.RichTextBox_Text.Text = strBuffer
        '        Me.TextBox_Text.ScrollBars
                
                arrTagData(i - 1) = ""
                arrRawData(i - 1) = ""
   
            Else
            
                strbuttonRecognised = ""
                intMaxMatch = 0
                intCounter = 0
                c = 0
                
                Open Filename_Database For Input As #1
                While Not EOF(1)
buttonRecognise_SkipLine:
                a = 190
                b = 190
                d = 0
                intMatch = 0
                Me.AppTemplateArea.Cls
                    If EOF(1) Then
                        GoTo buttonRecognise_FileClose
                    End If
                    Input #1, arrRawData(c)
                    If Len(arrRawData(c)) < 102 Then
                        GoTo buttonRecognise_SkipLine
                    End If
                    arrTagData(c) = Mid(arrRawData(c), 1, 1)
                    arrRawData(c) = Mid(arrRawData(c), 3)
        '            Debug.Print arrRawData(c)
                    
                    For i = 1 To 10
                    For j = 1 To 10
                        If Mid(arrRawData(c), d + 1, 1) = vbBlack Then
        '                    userTemplateArea.PSet (i, j)
                            AppTemplateArea.PSet (a, b)
        '                    userTemplateArea.Circle (a, b), 110
        '                    AppTemplateArea.Line (a - 110, b - 110)-(a + 110, b - 110)
        '                    AppTemplateArea.Line (a + 110, b - 110)-(a + 110, b + 110)
        '                    AppTemplateArea.Line (a + 110, b + 110)-(a - 110, b + 110)
        '                    AppTemplateArea.Line (a - 110, b + 110)-(a - 110, b - 110)
        '                    Debug.Print ""
                            If Mid(strData, d + 1, 1) = vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        Else
                            If Mid(strData, d + 1, 1) <> vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        End If
                        d = d + 1
                        b = b + (Me.AppTemplateArea.Height - 200) / 10
                    Next j
                    b = 190
                    a = a + (Me.AppTemplateArea.Width - 200) / 10
                    Next i
                    If intMaxMatch < intMatch Then
                        intMaxMatch = intMatch
                        strbuttonRecognised = arrTagData(c)
                        intCounter = c
                        Me.progressRecognition.Value = intMaxMatch
                        If intMaxMatch > 90 Then
                            GoTo buttonRecognise_FileClose
                        End If
                    End If
                    c = c + 1
                Wend
buttonRecognise_FileClose:
                Close #1
            End If
    End If
    
    Me.buttonRecognise.Enabled = False
    Me.progressRecognition.Value = 0
    
    If strbuttonRecognised <> "" Then
        AppTemplateArea.Cls
        
        With AppTemplateArea
        .ForeColor = vbWhite
        
        End With
        
        
        a = 190
        b = 190
        d = 0
        For i = 1 To 10
        For j = 1 To 10
            If Mid(arrRawData(intCounter), d + 1, 1) = vbBlack Then
                AppTemplateArea.PSet (a, b)
                AppTemplateArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                AppTemplateArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                AppTemplateArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                AppTemplateArea.Line (a - 110, b + 110)-(a - 110, b - 110)
            End If
            d = d + 1
            b = b + (Me.AppTemplateArea.Height - 200) / 10
        Next j
        b = 190
        a = a + (Me.AppTemplateArea.Width - 200) / 10
        Next i
    End If
    
    If strbuttonRecognised <> "" Then
        'The highest posible of drawn character is buttonRecognised as   'X'
        Me.textResult.Caption = intMaxMatch & "% Match with character '" & strbuttonRecognised & "'"
        Me.textResult.ToolTipText = intMaxMatch & "%"
        '& " , " & intMaxMatch & "%"
        
        Me.DrawWidth = 2
        
  
    End If
    

End Sub

Private Sub buttonLearn_Click()

    Me.TeachLabelText.FontBold = False
    Me.TeachLabelText.Caption = "Enter a character to be buttonLearn"
    
    
    Me.buttonLearn.Visible = False
    Me.buttonRecognise.Visible = False
    Me.buttonClearScreen.Visible = False
  
    
    Me.buttonLearnConfirm.Visible = True
    Me.buttonLearnCancel.Visible = True
    Me.inputLearnCharacter.Visible = True

    Me.inputLearnCharacter.Text = ""
    Me.inputLearnCharacter.SetFocus
    Me.buttonLearnConfirm.Enabled = False
    
End Sub

Private Sub GraspRawData()
Dim bool1stScan As Boolean
Dim ax As Integer
Dim ay As Integer
Dim bx As Integer
Dim by As Integer

bool1stScan = True
strData = ""

c = 0

Me.userTemplateArea.Cls
For i = 1 To userDrawArea.Width Step 100
    For j = 1 To userDrawArea.Height Step 100
        If userDrawArea.Point(i, j) = vbBlack Then
            userTemplateArea.PSet (i, j)
            If Not bool1stScan Then
                If i < ax Then
                    ax = i
                    End If
                If i > bx Then
                    bx = i
                    End If
                If j < ay Then
                    ay = j
                    End If
                If j > by Then
                    by = j
                    End If
            Else
                bool1stScan = False
                ax = i
                bx = i
                ay = j
                by = j
            End If
        End If
Next j, i

'MsgBox ""

If bx - ax <> 0 And by - ay <> 0 Then

    a = 190
    b = 190
    
    Me.userTemplateArea.Cls
    For i = ax To bx - (bx - ax) / 10 Step (bx - ax) / 10
        For j = ay To by - (by - ay) / 10 Step (by - ay) / 10
            If userDrawArea.Point(i, j) = vbBlack Then
                userTemplateArea.PSet (a, b)
    '''            userTemplateArea.Circle (a, b), 110
                userTemplateArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                userTemplateArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                userTemplateArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                userTemplateArea.Line (a - 110, b + 110)-(a - 110, b - 110)
    '                    userTemplateArea.FillStyle = vbSolid
    '                    userTemplateArea.FillColor = vbBlack
    '                    userTemplateArea.fil
    '            RawData(c) = userDrawArea.Point(i, j)
                strData = strData & userDrawArea.Point(i, j)
                c = c + 1
            Else
    '            RawData(c) = 1
                strData = strData & 1
                c = c + 1
            End If
            b = b + (Me.userTemplateArea.Height - 200) / 10
        Next j
        b = 190
        a = a + (Me.userTemplateArea.Width - 200) / 10
    Next i

End If

End Sub


Private Sub inputLearnCharacter_GotFocus()
    Me.inputLearnCharacter.SelStart = 0
    Me.inputLearnCharacter.SelLength = Len(Me.inputLearnCharacter.Text)
End Sub

Private Sub inputLearnCharacter_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Len(Me.inputLearnCharacter.Text) = 1 Then
        Me.buttonLearnConfirm.Enabled = True
        Me.buttonLearnConfirm.SetFocus
    ElseIf Len(Me.inputLearnCharacter.Text) > 1 Then
        Me.inputLearnCharacter.Text = ""
        Me.buttonLearnConfirm.Enabled = False
    Else
        Me.buttonLearnConfirm.Enabled = False
    End If
    
End Sub

Private Function StatusWindow(Optional ByVal strBuffer As String) As String
Dim Status_OpenButton As String
Dim Status_buttonLearnButton As String
Dim Status_buttonLearnConfirmButton As String
Dim Status_buttonLearnCancelButton As String
Dim Status_buttonRecogniseButton As String
Dim Status_buttonClearScreenButton As String
Dim Status_ExitButton As String
Dim Status_DrawArea As String
Dim Status_DatabaseArea As String
Dim Status_DataArea As String
Dim Status_Form As String
    
    Status_OpenButton = "Tips && Help : Open File - Just click the Open button to open FILE then type in your FILENAME - Try it..."
    Status_buttonLearnButton = "Tips && Help : buttonLearn - Just click the buttonLearn button to buttonLearn then type in your buttonLearn CHARACTER - Try it..."
    Status_buttonLearnConfirmButton = "Tips && Help : Confirm the character ENTER in the textbox is match with the drawn character in the Draw Area."
    Status_buttonLearnCancelButton = "Tips && Help : Mispressed or Give up or Do not want to buttonLearn."
    Status_buttonRecogniseButton = "Tips && Help : buttonRecognise Text - Just click the this button to buttonRecognise drawn character then the buttonRecognised character will display at the BOTTOM. Try it..."
    Status_buttonClearScreenButton = "Tips && Help : Clear Screen - Just click the this button to CLEAR screen then continue to draw character. Try it..."
    Status_ExitButton = "Tips && Help : Exit this software."
    Status_DrawArea = "Tips && Help : Your mouse pointer is in the Draw Area, just Click and drag to draw a character..."
    Status_DatabaseArea = "Tips && Help : This is Database Area, which it will display a buttonRecognised character from the database when you click buttonRecognise Button."
    Status_DataArea = "Tips && Help : This Data Area act as a buffer storage of Draw Area when user draw in Draw Area or to retrieve database data when user click Open Button."
    Status_Form = "Tips && Help : Thanks for using this software. Just move the mouse to the Draw Area, then drag the mouse when you want to draw a character to be buttonRecognised. Nice Try ..."
    
    Select Case strBuffer
    Case "OpenButton": strBuffer = Status_OpenButton
    Case "buttonLearnButton": strBuffer = Status_buttonLearnButton
    Case "buttonLearnConfirmButton": strBuffer = Status_buttonLearnConfirmButton
    Case "buttonLearnCancelButton": strBuffer = Status_buttonLearnCancelButton
    Case "buttonRecogniseButton": strBuffer = Status_buttonRecogniseButton
    Case "buttonClearScreenButton": strBuffer = Status_buttonClearScreenButton
    Case "ExitButton": strBuffer = Status_ExitButton
    Case "DrawArea": strBuffer = Status_DrawArea
    Case "DatabaseArea": strBuffer = Status_DatabaseArea
    Case "DataArea": strBuffer = Status_DataArea
    Case "Form": strBuffer = Status_Form
    End Select
    
    StatusWindow = strBuffer

End Function

Private Function BinToDec(strBin As String) As Integer

i = Len(strBin)
While i > 0
    If Mid(strBin, i, 1) = "1" Then
        BinToDec = BinToDec + 2 ^ (Len(strBin) - i)
    End If
    i = i - 1
Wend

End Function

Private Function DecToBin(intDec As Integer, intDigit As Integer) As String
Dim intTemp As Integer

While intDec > 0 And intDigit > 0
    intDigit = intDigit - 1
    intTemp = intDec Mod 2
    If intTemp Then
        DecToBin = "1" & DecToBin
        intDec = (intDec - 1) / 2
    Else
        DecToBin = "0" & DecToBin
        intDec = intDec / 2
    End If
Wend

While intDigit
    intDigit = intDigit - 1
    DecToBin = "0" & DecToBin
Wend

End Function

