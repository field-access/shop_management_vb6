VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPurchase 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   6960
      Width           =   3735
      _ExtentX        =   6588
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
      Connect         =   "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sale_details"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   7200
      TabIndex        =   17
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   855
      Left            =   4440
      TabIndex        =   16
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "save"
      Height          =   855
      Left            =   1920
      TabIndex        =   15
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtAmount 
      Height          =   735
      Left            =   11880
      TabIndex        =   14
      Top             =   6720
      Width           =   2775
   End
   Begin VB.TextBox txtGst 
      Height          =   735
      Left            =   11880
      TabIndex        =   13
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox txtQty 
      Height          =   735
      Left            =   11880
      TabIndex        =   12
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtMrp 
      Height          =   735
      Left            =   11880
      TabIndex        =   11
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtPname 
      Height          =   735
      Left            =   11880
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtProId 
      Height          =   735
      Left            =   11880
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtInvNo 
      Height          =   735
      Left            =   11880
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPurchase.frx":0000
      Height          =   6255
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11033
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
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
   Begin VB.Label Label7 
      Caption         =   "Amount:"
      Height          =   735
      Left            =   9600
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "GST:"
      Height          =   735
      Left            =   9720
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity:"
      Height          =   735
      Left            =   9600
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "MRP:"
      Height          =   735
      Left            =   9600
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "pro_name:"
      Height          =   735
      Left            =   9600
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "product_id:"
      Height          =   735
      Left            =   9600
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "invoice no:"
      Height          =   735
      Left            =   9600
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Initialize the ADODC control
    conn
    ' Bind the DataGrid to the ADODC control
    Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub cmdAdd_Click()
    ' Clear fields
    ClearFields
    
    ' Generate new Invoice No
    Dim sql As String
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    R.Open "SELECT MAX(INV_NO) FROM sale_details", C, adOpenStatic, adLockReadOnly
    If Not R.EOF Then
        txtInvNo.Text = R.Fields(0).Value + 1
    Else
        txtInvNo.Text = 1
    End If
    R.Close
    Set R = Nothing
    txtInvNo.Locked = True
    
    cmdSave.Enabled = True
    txtProId.SetFocus
End Sub

Private Sub cmdSave_Click()
    ' Validate fields
    If ValidateFields() Then
        ' Insert new record
        Dim sql As String
        sql = "INSERT INTO sale_details (INV_NO, PRO_ID, PNAME, MRP, QTY, GST, AMOUNT) VALUES (" & _
              txtInvNo.Text & ", " & txtProId.Text & ", '" & txtPname.Text & "', " & txtMrp.Text & ", " & _
              txtQty.Text & ", " & txtGst.Text & ", " & txtAmount.Text & ")"
        C.Execute sql
        MsgBox "Record saved successfully.", vbInformation, "Success"
        
        ' Refresh DataGrid
        Adodc1.Refresh
        
        ' Clear fields
        ClearFields
    End If
End Sub

Private Sub cmdUpdate_Click()
    ' Validate fields
    If ValidateFields() Then
        ' Update existing record
        Dim sql As String
        sql = "UPDATE sale_details SET PRO_ID = " & txtProId.Text & ", PNAME = '" & txtPname.Text & "', " & _
              "MRP = " & txtMrp.Text & ", QTY = " & txtQty.Text & ", GST = " & txtGst.Text & ", " & _
              "AMOUNT = " & txtAmount.Text & " WHERE INV_NO = " & txtInvNo.Text
        C.Execute sql
        MsgBox "Record updated successfully.", vbInformation, "Success"
        
        ' Refresh DataGrid
        Adodc1.Refresh
    End If
End Sub

Private Sub cmdDelete_Click()
    ' Delete the current record
    Dim sql As String
    sql = "DELETE FROM sale_details WHERE INV_NO = " & txtInvNo.Text
    C.Execute sql
    MsgBox "Record deleted successfully.", vbInformation, "Success"
    
    ' Refresh DataGrid
    Adodc1.Refresh
End Sub

Private Sub cmdClear_Click()
    ' Clear the form fields
    ClearFields
End Sub

Private Sub cmdExit_Click()
    ' Close the form
    Unload Me
End Sub

Private Sub ClearFields()
    txtInvNo.Text = ""
    txtProId.Text = ""
    txtPname.Text = ""
    txtMrp.Text = ""
    txtQty.Text = ""
    txtGst.Text = ""
    txtAmount.Text = ""
End Sub

Private Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim(txtInvNo.Text) = "" Or Trim(txtProId.Text) = "" Or Trim(txtPname.Text) = "" Or _
       Trim(txtMrp.Text) = "" Or Trim(txtQty.Text) = "" Or Trim(txtGst.Text) = "" Or _
       Trim(txtAmount.Text) = "" Then
        MsgBox "Please fill all fields before saving.", vbCritical, "Validation Error"
        Exit Function
    End If
    ValidateFields = True
End Function

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

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtProId
End Sub

Private Sub txtProId_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtPname
End Sub

Private Sub txtPname_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtMrp
End Sub

Private Sub txtMrp_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtQty
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtGst
End Sub

Private Sub txtGst_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, txtAmount
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    HandleEnterKey KeyAscii, cmdSave
End Sub

