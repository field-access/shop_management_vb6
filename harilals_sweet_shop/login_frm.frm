VERSION 5.00
Begin VB.Form login_frm 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "enter username"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "login_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Hardcoded credentials
    Const VALID_USERNAME As String = "elon"
    Const VALID_PASSWORD As String = "musk"
    
    ' Get input values
    Dim username As String
    Dim password As String
    
    username = Trim(Text1.Text)
    password = Trim(Text2.Text)
    
    ' Validate input
    If username = "" Or password = "" Then
        MsgBox "Please enter both username and password!", vbExclamation
        Exit Sub
    End If
    
    ' Check credentials
    If username = VALID_USERNAME And password = VALID_PASSWORD Then
        MsgBox "Login successful! Welcome " & username, vbInformation
        MDIForm1.Show
        Unload Me
    Else
        MsgBox "Invalid username or password!", vbCritical
        Text2.Text = "" ' Clear password field
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbYes Then
        End
    End If
End Sub

' Add these events for better user experience
Private Sub Text1_KeyPress(keyascii As Integer)
    If keyascii = 13 Then ' Enter key
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(keyascii As Integer)
    If keyascii = 13 Then ' Enter key
        Command1_Click
    End If
End Sub

