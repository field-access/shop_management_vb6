VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Product_management 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "delete"
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9600
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select * from product order by pro_id"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton command5 
      Caption         =   "exit"
      Height          =   555
      Left            =   7800
      TabIndex        =   17
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "update"
      Height          =   555
      Left            =   7800
      TabIndex        =   16
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add new"
      Height          =   555
      Left            =   5640
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   435
      Left            =   5640
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Product_management.frx":0000
      Height          =   3615
      Left            =   5640
      TabIndex        =   13
      Top             =   1800
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
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
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   435
      Left            =   2280
      TabIndex        =   10
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   2400
      TabIndex        =   8
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   2520
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "product details "
      Height          =   855
      Left            =   3480
      TabIndex        =   12
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "stock_quantity"
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   6720
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "pro_sell_rate"
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   5640
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "mfg_rate"
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "product_type"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "product_name"
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "product_id"
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1230
   End
End
Attribute VB_Name = "Product_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
        MsgBox "Please fill all fields.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    sql = "INSERT INTO product VALUES('" & Text1.Text & "', '" & _
      Text2.Text & "', '" & Text3.Text & "', " & Val(Text4.Text) & ", " & _
      Val(Text5.Text) & ", " & Val(Text6.Text) & ")"
    
    Set R = C.Execute(sql)
    
    MsgBox "Record saved successfully.", vbInformation, "Success"
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh

    Text2.SetFocus
    
    sql = "SELECT max(pro_id) FROM PRODUCT order by pro_id"
    Set R = C.Execute(sql)

    Do While Not R.EOF
    Text1.Text = R.Fields(0) + 1
    R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
End Sub

Private Sub Command3_Click()
    
    sql = "UPDATE product SET pro_id = '" & Text1.Text & "' , pro_name='" & Text2.Text & "', pro_type='" & Text3.Text & "', " & _
          "mfg_rate=" & Val(Text4.Text) & ", pro_sell_rate=" & Val(Text5.Text) & ", " & _
          "stock_qty=" & Val(Text6.Text) & " WHERE pro_id='" & Text1.Text & "'"
    Set R = C.Execute(sql)
    
    MsgBox "record updated", vbInformation, "success"
    
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click() ' Load Selected Record into Textboxes
    If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
        Text1.Text = Adodc1.Recordset.Fields("pro_id")
        Text2.Text = Adodc1.Recordset.Fields("pro_name")
        Text3.Text = Adodc1.Recordset.Fields("pro_type")
        Text4.Text = Adodc1.Recordset.Fields("mfg_rate")
        Text5.Text = Adodc1.Recordset.Fields("pro_sell_rate")
        Text6.Text = Adodc1.Recordset.Fields("stock_qty")
    End If
End Sub

Private Sub Form_Load()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub enterkey(keyascii As Integer, nextcontrol As Control)
    If keyascii = 13 Then
        nextcontrol.SetFocus
        keyascii = 0
    End If
End Sub

Private Sub Text1_KeyPress(keyascii As Integer)
    enterkey keyascii, Text2
End Sub

Private Sub Text2_KeyPress(keyascii As Integer)
    enterkey keyascii, Text3
End Sub

Private Sub Text3_KeyPress(keyascii As Integer)
    enterkey keyascii, Text4
End Sub

Private Sub Text4_KeyPress(keyascii As Integer)
    enterkey keyascii, Text5
End Sub

Private Sub Text5_KeyPress(keyascii As Integer)
    enterkey keyascii, Text6
End Sub

Private Sub Text6_KeyPress(keyascii As Integer)
    enterkey keyascii, Command1
End Sub

Private Sub Command1_KeyPress(keyascii As Integer)
enterkey keyascii, Command2
End Sub








