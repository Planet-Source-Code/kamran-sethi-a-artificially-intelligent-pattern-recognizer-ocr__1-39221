VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TSOpenMode
   ForAppending = 8
   ForReading = 1
End Enum

Public Enum TSFormat
   TristateUseDefault = -2
   TristateTrue = -1
   TristateFalse = 0
End Enum

Private mintHandle As Integer       ' File Handle
Private mstrFileName As String      ' File Name To Open
Private mintOpenMode As TSOpenMode  ' Open Mode
Private mintFormat As TSFormat      ' Format
Private mboolEOF As Boolean         ' Are we at end of file

Public Function OpenTextFile(strFile As String, _
                             Optional intOpenMode As TSOpenMode = ForReading, _
                             Optional intFormat As TSFormat = TristateTrue) As Boolean
   If strFile <> "" Then
      mstrFileName = strFile
      mboolEOF = False
      mintOpenMode = intOpenMode
      mintFormat = intFormat
      
      If intOpenMode = ForReading Then
         Open mstrFileName For Input As #mintHandle
      Else
         Open mstrFileName For Output As #mintHandle
      End If
      
      OpenTextFile = True
   Else
      OpenTextFile = False
   End If
End Function

Public Function CloseFile() As Boolean
   ' Valid File Handle ?
   If mintHandle <> 0 Then
      Close #mintHandle
      
      mintHandle = 0
      
      CloseFile = True
   Else
      CloseFile = False
   End If
End Function

Public Function ReadLine() As String
   Dim strLine As String

   ' Valid File Handle ?
   If mintHandle <> 0 Then
      ' Are we at End Of File ?
      If EOF(mintHandle) Then
         mboolEOF = True
         ReadLine = ""
      Else
         Line Input #mintHandle, strLine
         ReadLine = strLine
      End If
   Else
      ReadLine = ""
      mboolEOF = True
   End If
End Function

Property Let FileName(ByVal strFile As String)
   mstrFileName = strFile
End Property

Property Get FileName() As String
   FileName = mstrFileName
End Property

Property Get FileTooBig() As Boolean
   If FileLen(mstrFileName) > 32768 Then
      FileTooBig = True
   Else
      FileTooBig = False
   End If
End Property

Public Function WriteLine(Optional strWrite As String = vbCrLf) As Boolean
   If mintHandle <> 0 Then
      Print #mintHandle, strWrite & vbCrLf
      WriteLine = True
   End If
End Function

Private Sub Class_Initialize()
      mintHandle = FreeFile
End Sub

Private Sub Class_Terminate()
   Dim boolRet As Boolean
   
   boolRet = CloseFile()
End Sub

Public Property Get AtEndOfStream() As Boolean
   AtEndOfStream = mboolEOF
End Property

Public Property Let AtEndOfStream(ByVal boolEOF As Boolean)
   mboolEOF = boolEOF
End Property
