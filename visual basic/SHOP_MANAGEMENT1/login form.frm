VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C0C0&
      Caption         =   "close"
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "login"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "date:"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "user name"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label3.Caption = Format(Now, "dd mm yyyy & hh:mm")
    conn
    Me.WindowState = 2  'Maximized
    frmProduct.Hide
    frmPurchase.Hide
    frmReports.Hide
    frmSales.Hide
    frmSupplier.Hide
End Sub

Private Sub Command1_Click()
If Text1.Text = "sumit" And Text2.Text = "singh" Then
    MDIForm1.Show
    Me.Hide
    
Else
    MsgBox "invalid username or password", vbInformation, "electronic shop management"
    End
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub



