VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Underground Technologies Picture Viewer"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00A2D544&
      Height          =   345
      Left            =   4200
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Preview"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List Options"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2D544&
      Height          =   3660
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2D544&
      Height          =   1980
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2D544&
      Height          =   1770
      Left            =   2520
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
   Begin VB.ListBox List2 
      ForeColor       =   &H00A2D544&
      Height          =   510
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
Text1.Text = "false"

    Dim Moo
    If Check1.Value = 0 Then
    Unload Form2
    Exit Sub
    End If
    
    If List2.ListCount = 0 Then
        Exit Sub
    Else
    End If
    
    Form2.Show
    For Moo = 0 To List2.ListCount - 1
    If Text1.Text = "true" Then Check1.Value = 0: Unload Form2: Exit For
        List1.Selected(Moo) = True
        List2.Selected(Moo) = True
        Form2.Label1.Caption = List2.List(List2.ListIndex)
        crt List2.List(List2.ListIndex)
        pause 2.5
    Next Moo
pause 2.5
Check1.Value = 0
Unload Form2
'End If

End Sub

Private Sub Dir1_Change()
File1 = Dir1

End Sub

Private Sub Drive1_Change()
On Error GoTo bad_drive
Dir1 = Drive1
Exit Sub
bad_drive:
Drive1 = "C:\"
Dir1 = "C:\"
End Sub


Private Sub File1_DblClick()
If Dir1 = "C:\" Then List2.AddItem "C:\" & File1: List1.AddItem File1
If Dir1 <> "C:\" Then List2.AddItem Dir1 & "\" & File1: List1.AddItem File1
End Sub

Private Sub Form_Load()
Dir1 = "C:\My Documents"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub List1_Click()
List1.ListIndex = (List1.ListIndex)
List2.ListIndex = (List1.ListIndex)
End Sub

Private Sub List1_DblClick()
List2.RemoveItem (List1.ListIndex)
List1.RemoveItem (List1.ListIndex)
End Sub


Private Sub Command1_Click()
Me.PopupMenu Form3.mnu_preview
End Sub
