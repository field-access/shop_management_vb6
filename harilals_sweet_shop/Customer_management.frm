VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Customer_management 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "addnew"
      Height          =   615
      Left            =   10680
      TabIndex        =   16
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "exit"
      Height          =   615
      Left            =   8400
      TabIndex        =   15
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "update"
      Height          =   615
      Left            =   8400
      TabIndex        =   14
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "delete"
      Height          =   615
      Left            =   6240
      TabIndex        =   13
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   6120
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer_management.frx":0000
      Height          =   4455
      Left            =   5880
      TabIndex        =   10
      Top             =   1320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
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
      Connect         =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from customer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   585
      Left            =   3000
      TabIndex        =   7
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "customer details"
      Height          =   345
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "cus_bal:"
      Height          =   345
      Left            =   840
      TabIndex        =   4
      Top             =   5880
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "phone_no:"
      Height          =   345
      Left            =   840
      TabIndex        =   3
      Top             =   4920
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "address:"
      Height          =   345
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "cust_name:"
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "cust_id"
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   1110
   End
End
Attribute VB_Name = "Customer_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sql = "insert into customer values ('" & Text1.Text & " ', ' " & Text2.Text & " ' , '" & Text3.Text & "' , '" & Text4.Text & "' , " & Text5.Text & ")"
Set R = C.Execute(sql)
MsgBox "record saved", vbInformation
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
MsgBox "record deleted"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update
MsgBox "record updated"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
sql = "select max(cus_id) from customer"
Set R = C.Execute(sql)
Do While Not R.EOF
Text1.Text = R.Fields(0) + 1
R.MoveNext
Loop
R.Close
Adodc1.Refresh
DataGrid1.Refresh
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text2.SetFocus
End Sub

Private Sub Form_Load()
conn
Adodc1.Refresh
DataGrid1.Refresh
End Sub

    
