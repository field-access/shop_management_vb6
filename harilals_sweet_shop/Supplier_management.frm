VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Supplier_management 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Console"
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
   Begin VB.CommandButton Command5 
      Caption         =   "exit"
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "add new"
      Height          =   615
      Left            =   8280
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   615
      Left            =   6360
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "update"
      Height          =   615
      Left            =   8280
      TabIndex        =   12
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1440
      Top             =   5760
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
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
      RecordSource    =   "select * from sup_det"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Supplier_management.frx":0000
      Height          =   3855
      Left            =   6360
      TabIndex        =   10
      Top             =   1200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "supplier management"
      Height          =   735
      Left            =   4200
      TabIndex        =   16
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "comp_name:"
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   4800
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "contact no:"
      Height          =   240
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "address:"
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "name:"
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "sup_id:"
      Height          =   240
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1155
   End
End
Attribute VB_Name = "Supplier_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command1_Click() ' Save Button
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
        MsgBox "Please fill all fields.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    ' Validate numeric fields
    If Not IsNumeric(Text4.Text) Then
        MsgBox "Contact number must be numeric.", vbExclamation, "Input Error"
        Text4.SetFocus
        Exit Sub
    End If
    
    sql = "INSERT INTO sup_det VALUES('" & Text1.Text & "', '" & _
          Text2.Text & "', '" & Text3.Text & "', " & _
          Text4.Text & ", '" & Text5.Text & "')"
    
    Set R = C.Execute(sql)
    
    MsgBox "Record saved successfully.", vbInformation, "Success"
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command2_Click() ' Update Button
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
        MsgBox "Please fill all fields.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    ' Validate numeric fields
    If Not IsNumeric(Text4.Text) Then
        MsgBox "Contact number must be numeric.", vbExclamation, "Input Error"
        Text4.SetFocus
        Exit Sub
    End If
    
    sql = "UPDATE sup_det SET name='" & Text2.Text & "', " & _
          "addr='" & Text3.Text & "', cont_no=" & Text4.Text & ", " & _
          "comp_name='" & Text5.Text & "' WHERE sup_id='" & Text1.Text & "'"
    
    Set R = C.Execute(sql)
    
    MsgBox "Record updated successfully.", vbInformation, "Success"
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub Command3_Click() ' Delete Button
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo) = vbYes Then
        Adodc1.Recordset.Delete
        MsgBox "Record deleted successfully.", vbInformation, "Success"
        Adodc1.Refresh
        DataGrid1.Refresh
    End If
End Sub

Private Sub Command4_Click() ' Add New Button
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
    
    sql = "SELECT max(sup_id) FROM sup_det"
    Set R = C.Execute(sql)
    
    Do While Not R.EOF
        Text1.Text = Val(R.Fields(0)) + 1
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    
    ' Clear all fields
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text2.SetFocus
End Sub

Private Sub Command5_Click() ' Exit Button
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
        Text1.Text = Adodc1.Recordset.Fields("sup_id")
        Text2.Text = Adodc1.Recordset.Fields("name")
        Text3.Text = Adodc1.Recordset.Fields("addr")
        Text4.Text = Adodc1.Recordset.Fields("cont_no")
        Text5.Text = Adodc1.Recordset.Fields("comp_name")
    End If
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
    enterkey keyascii, Command1
End Sub

Private Sub Command1_KeyPress(keyascii As Integer)
    enterkey keyascii, Command2
End Sub



    



