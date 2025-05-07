VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Raw_material_management 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Fax"
      Size            =   15
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "exit:"
      Height          =   495
      Left            =   11040
      TabIndex        =   22
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "update:"
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete:"
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save:"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "addnew"
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   3840
      TabIndex        =   17
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   3840
      TabIndex        =   16
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   3840
      TabIndex        =   15
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3600
      TabIndex        =   13
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3600
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   585
      Left            =   3600
      TabIndex        =   11
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   1200
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Raw_material_management.frx":0000
      Height          =   5295
      Left            =   7680
      TabIndex        =   1
      Top             =   1080
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9340
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Fax"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Fax"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      Left            =   8640
      Top             =   6840
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      RecordSource    =   "select * from raw_material"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Fax"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "stock_qty:"
      Height          =   345
      Left            =   600
      TabIndex        =   9
      Top             =   7080
      Width           =   1605
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "rate:"
      Height          =   345
      Left            =   600
      TabIndex        =   8
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "size:"
      Height          =   345
      Left            =   600
      TabIndex        =   7
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "unit of measurement:"
      Height          =   345
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   3480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "type:"
      Height          =   345
      Left            =   600
      TabIndex        =   5
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "company_name:"
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   2550
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "name:"
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "raw_id:"
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "raw material management"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   4200
   End
End
Attribute VB_Name = "Raw_material_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click() ' Save Button
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or _
       Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
        MsgBox "Please fill all fields.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    sql = "INSERT INTO raw_material VALUES('" & Text1.Text & "', '" & _
          Text2.Text & "', '" & Text3.Text & "', '" & Text4.Text & "', '" & _
          Text5.Text & "', '" & Text6.Text & "', " & Val(Text7.Text) & ", " & _
          Val(Text8.Text) & ")"
    
    Set R = C.Execute(sql)
    
    MsgBox "Record saved successfully.", vbInformation, "Success"
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command1_Click() ' Add New Button
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
    
    Text2.SetFocus
    
    sql = "SELECT count(raw_id) FROM raw_material"
    Set R = C.Execute(sql)
    
    Do While Not R.EOF
        Text1.Text = R.Fields(0) + 1
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    
    ' Clear all fields
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
End Sub

Private Sub Command3_Click() ' Delete Button
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command4_Click() ' Update Button
    sql = "UPDATE raw_material SET raw_id = '" & Text1.Text & "', name='" & Text2.Text & "', " & _
          "company_name='" & Text3.Text & "', type='" & Text4.Text & "', " & _
          "unit_of_measurement='" & Text5.Text & "', size='" & Text6.Text & "', " & _
          "rate=" & Val(Text7.Text) & ", stock_qty=" & Val(Text8.Text) & " " & _
          "WHERE raw_id='" & Text1.Text & "'"
    
    Set R = C.Execute(sql)
    
    MsgBox "Record updated", vbInformation, "Success"
    
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command5_Click() ' Exit Button
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
        Text1.Text = Adodc1.Recordset.Fields("raw_id")
        Text2.Text = Adodc1.Recordset.Fields("name")
        Text3.Text = Adodc1.Recordset.Fields("company_name")
        Text4.Text = Adodc1.Recordset.Fields("type")
        Text5.Text = Adodc1.Recordset.Fields("unit_of_measurement")
        Text6.Text = Adodc1.Recordset.Fields("siz")
        Text7.Text = Adodc1.Recordset.Fields("rate")
        Text8.Text = Adodc1.Recordset.Fields("stock_qty")
    End If
End Sub

Private Sub Form_Load()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

' Keyboard navigation
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
    enterkey keyascii, Text7
End Sub

Private Sub Text7_KeyPress(keyascii As Integer)
    enterkey keyascii, Text8
End Sub

Private Sub Text8_KeyPress(keyascii As Integer)
    enterkey keyascii, Command2
End Sub

Private Sub Command1_KeyPress(keyascii As Integer)
    enterkey keyascii, Command1
End Sub


