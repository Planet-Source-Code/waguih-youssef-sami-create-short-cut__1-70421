VERSION 5.00
Begin VB.Form frmShortCut 
   Caption         =   "Create Short Cut"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir2 
      Height          =   1215
      Left            =   3435
      TabIndex        =   9
      Top             =   1680
      Width           =   2715
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   3435
      TabIndex        =   8
      Top             =   1110
      Width           =   2145
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   285
      TabIndex        =   6
      Top             =   1500
      Width           =   2715
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   285
      TabIndex        =   5
      Top             =   1050
      Width           =   2070
   End
   Begin VB.TextBox txtLinkPath 
      Height          =   375
      Left            =   270
      TabIndex        =   4
      Top             =   510
      Width           =   2715
   End
   Begin VB.TextBox txtDestination 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3435
      TabIndex        =   3
      Top             =   495
      Width           =   2715
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   3435
      TabIndex        =   2
      Top             =   3540
      Width           =   2715
   End
   Begin VB.CommandButton cmdCreateShortCut 
      Caption         =   "Create Short Cut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3765
      TabIndex        =   1
      Top             =   4185
      Width           =   2010
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   285
      TabIndex        =   0
      Top             =   2850
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   660
      Left            =   3615
      Top             =   4050
      Width           =   2280
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Link Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3375
      TabIndex        =   11
      Top             =   3135
      Width           =   2715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   3150
      X2              =   3150
      Y1              =   135
      Y2              =   4600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Destination Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3390
      TabIndex        =   10
      Top             =   135
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Link Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   285
      TabIndex        =   7
      Top             =   135
      Width           =   2715
   End
End
Attribute VB_Name = "frmShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function fCreateShellLink Lib "Vb5stkit.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Sub cmdCreateShortCut_Click()
Dim lReturn As Long
lReturn = fCreateShellLink("..\..\Desktop", txtName.Text, txtLinkPath.Text, "")

Dim ShortCutName
ShortCutName = txtName.Text & ".lnk"
Dim Destination As String, SourceFile As String
Destination = txtDestination.Text & ShortCutName

Dim WinDir As String
WinDir = MyWinDir

SourceFile = WinDir & "\desktop\" & ShortCutName

Dim Msg
If Dir(Destination) <> "" Then
Msg = MsgBox("File Exists Over Write?", vbYesNo)
    If Msg = vbYes Then
    Kill Destination
    Name SourceFile As Destination
    MsgBox " Short Cut For " & SourceFile & " Created", vbOKOnly
    Else
    Exit Sub
    End If
End If

'**********Name Function is for Moving File**********
'*******From Desktop to the other Folder************
Name SourceFile As Destination
'*******No Need For ClipBoard***********************
'Clipboard.Clear
'Clipboard.SetText SourceFile
'SourceFile = Clipboard.GetText
'******************Copy and then delete the orign****************************
'FileCopy SourceFile, Destination
'Kill SourceFile

MsgBox " Short Cut For " & SourceFile & " Created", vbOKOnly
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
txtLinkPath.Text = File1.Path & "\" & File1.filename

End Sub


Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

Private Sub Dir2_Click()
txtDestination.Text = Dir2.Path & "\"
End Sub
Private Function InStrRev(ByVal sText As String, Delimeter As String) As Integer
Dim i As Integer
For i = Len(sText) - Len(Delimeter) To 1 Step -1
If Mid(sText, i, Len(Delimeter)) = Delimeter Then
InStrRev = i
Exit For
End If
Next i

End Function

Private Sub txtLinkPath_Change()
Dim Pos As Integer
Pos = InStrRev(txtLinkPath.Text, "\")
txtName.Text = Right(txtLinkPath.Text, Pos)
End Sub

Private Sub Form_Load()
Dir1.Path = "c:\"
Dir2.Path = "c:\"
End Sub
Public Function FileExists(FSpec$) As Byte
   FileExists = LenB(Dir$(FSpec$))
End Function

Private Function MyWinDir() As String
Dim strWinDir As String
Dim lngSize As Long
    
    strWinDir = Space(255)
    lngSize = 255
    
    Call GetWindowsDirectory(strWinDir, lngSize - 1)

Dim intPosition As Integer
   
    intPosition = InStr(strWinDir, Chr$(0))
    If intPosition Then
       MyWinDir = Left(strWinDir, intPosition - 1)
    Else
    MyWinDir = strWinDir
    End If
End Function

