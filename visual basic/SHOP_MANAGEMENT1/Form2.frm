VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmProduct 
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Oklahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4080
      TabIndex        =   22
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   4080
      TabIndex        =   21
      Top             =   6000
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "refresh"
      Height          =   615
      Left            =   5760
      TabIndex        =   20
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      DataField       =   "MRP"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      DataField       =   "PRO_WATT"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   18
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      DataField       =   "PRO_COLOR"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      DataField       =   "PRO_SIZE"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      DataField       =   "PNAME"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   6045
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10663
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Oklahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Oklahoma"
         Size            =   11.25
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
   Begin VB.TextBox Text1 
      DataField       =   "PCODE"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   7080
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   8520
      Top             =   8040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   714
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
      RecordSource    =   "SELECT * FROM product ORDER BY pcode"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "update"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "add new"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "MRP"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   15
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "pro_watt"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   14
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "pro_color"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   13
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "pro_size"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   12
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "pro_type"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label pna 
      Caption         =   "pname"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   10
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "product Id"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "p_brand"
      BeginProperty Font 
         Name            =   "Oklahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc1.RecordSource = "SELECT * FROM product ORDER BY pcode"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = Adodc1
   
   End Sub



Private Sub Form_Load()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh
End Sub

Private Sub cmdAdd_Click()
    Adodc1.Refresh
    DataGrid1.Refresh
    
    ' Clear fields
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""

    sql = "SELECT MAX(PCODE)FROM PRODUCT"
    Set R = C.Execute(sql)
    Text1.Text = R.Fields(0) + 1
    Text1.Locked = True

    cmdSave.Enabled = True
    Text2.SetFocus
End Sub

Private Sub cmdSave_Click()
    ' Validate fields
    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Or _
       Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Or _
       Trim(Text7.Text) = "" Or Trim(Text8.Text) = "" Then
        MsgBox "Please fill all fields before saving.", vbCritical, "Validation Error"
        Exit Sub
    End If

    ' Proceed with saving if all fields are filled
    sql = "INSERT INTO product VALUES(" & Text1.Text & ", '" & Text2.Text & "', '" & Text3.Text & "', '" & _
          Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text7.Text & "', " & Text8.Text & ")"
    Set R = C.Execute(sql)
    MsgBox "Record saved successfully.", vbInformation, "Success"

    ' Refresh DataGrid after save
    Command1_Click
    
    ' Clear fields
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    cmdAdd.Enabled = True
    cmdAdd.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    Adodc1.Recordset.Update
    MsgBox "RECORD UPDATED"
    
    ' Refresh DataGrid
    Command1_Click
End Sub



Private Sub cmdDelete_Click()
    ' Delete the current record
    Adodc1.Recordset.Delete
    MsgBox "RECORD DELETED"
    
    ' Force Adodc to requery the database to update the recordset
    Adodc1.RecordSource = "SELECT * FROM product ORDER BY pcode"
    Adodc1.Refresh
    
    ' Rebind DataGrid
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = Adodc1
    
    ' Refresh DataGrid
    Command1_Click
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Adodc1.Recordset.EditMode <> adEditNone Then
        Adodc1.Recordset.CancelUpdate
    End If
End Sub

' Handle Enter Key for field navigation
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
    HandleEnterKey KeyAscii, Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, cmdSave
End Sub


