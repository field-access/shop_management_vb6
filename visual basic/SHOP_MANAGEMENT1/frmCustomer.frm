VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmCustomer 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Addnew"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   9240
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8640
      Top             =   8280
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from customer order by custid"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCustomer.frx":0000
      Height          =   4215
      Left            =   8640
      TabIndex        =   23
      Top             =   2520
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7435
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Delete"
      Height          =   615
      Left            =   8040
      TabIndex        =   21
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Update"
      Height          =   615
      Left            =   5880
      TabIndex        =   20
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   3600
      TabIndex        =   19
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   540
      Left            =   4200
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   4200
      TabIndex        =   17
      Top             =   7200
      Width           =   8655
   End
   Begin VB.TextBox Text5 
      Height          =   540
      Left            =   4200
      TabIndex        =   16
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   540
      Left            =   4200
      TabIndex        =   15
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   540
      Left            =   4200
      TabIndex        =   14
      Top             =   4800
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      Height          =   420
      Left            =   6600
      TabIndex        =   13
      Top             =   4080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      Height          =   420
      Left            =   4320
      TabIndex        =   12
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   4200
      TabIndex        =   11
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   540
      Left            =   4200
      TabIndex        =   10
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "aadhar no"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "address"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "pincode"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Email"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ph_number"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gender"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customer name"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1875
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   360
      Width           =   9975
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   0
         Picture         =   "frmCustomer.frx":0015
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Register New Customer"
         Height          =   735
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   4320
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Id"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Adodc1.Refresh
    DataGrid1.Refresh
    
    ' Clear fields
    Text2.Text = ""
    Text3.Text = ""
    Label3.Caption = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""

    sql = "SELECT MAX(custid)FROM customer"
    Set R = C.Execute(sql)
    Text1.Text = R.Fields(0) + 1
    Text1.Locked = True

    cmdSave.Enabled = True
    Text2.SetFocus
End Sub

Private Sub cmdSave_Click()
    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Or Trim(Label3.Caption) = "" Or _
       Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or _
       Trim(Text6.Text) = "" Or Trim(Text7.Text) = "" Then
        MsgBox "Please fill all fields before saving.", vbCritical, "Validation Error"
        Exit Sub
    End If

    sql = "INSERT INTO A customer VALUES(" & Text1.Text & ",'" & Text2.Text & "','" & Label3.Caption & "'," & Text3.Text & ",'" & Text4.Text & "'," & Text5.Text & ",'" & Text6.Text & "'," & Text7.Text & ")"
    Set R = C.Execute(sql)
    MsgBox "RECORD SAVED"
End Sub

Private Sub Form_Load()
    conn
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub option1_click()
Label3.Caption = "M"
End Sub

Private Sub option2_click()
Label3.Caption = "F"
End Sub

Private Sub HandleEnterKey(KeyAscii As Integer, NextControl As Control)
    If KeyAscii = 13 Then
        KeyAscii = 0
        NextControl.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, cmdSave
End Sub



