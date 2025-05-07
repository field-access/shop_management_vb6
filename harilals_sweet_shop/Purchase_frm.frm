VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form purchase_frm 
   Caption         =   "Purchase - Harilal's Sweet Shop"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   8160
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1508
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
      RecordSource    =   "select* from pro_master"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Purchase_frm.frx":0000
      Height          =   5175
      Left            =   8280
      TabIndex        =   17
      Top             =   240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      Begin VB.TextBox txtOrderNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtOrderDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtSuppId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox cboPaymentMethod 
         BeginProperty Font 
            Name            =   "Oklahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtAdvance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtDues 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox txtNetAmt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Method:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   4680
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   "raw material purchase form"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "purchase_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    conn
    ' Set current date in order date field
    txtOrderDate.Text = Format(Date, "dd/mm/yyyy")
    
    ' Add payment methods to combo box
    With cboPaymentMethod
        .AddItem "CASH"
        .AddItem "CHEQUE"
        .AddItem "CARD"
        .AddItem "UPI"
    End With
    
    ' Generate new order number
    GenerateOrderNo

End Sub



Private Sub GenerateOrderNo()
    On Error GoTo ErrorHandler
    
    ' Format: PO followed by YYMMDD and 3 digit sequence number
    Dim prefix As String
    Dim dateStr As String
    Dim seqNum As String
    
    prefix = "PO"
    dateStr = Format(Date, "yymmdd")
    
    ' Get last order number from database and increment
    ' For now, just use a simple counter
    seqNum = Format(1, "000") ' You should get this from the database
    
    txtOrderNo.Text = prefix & dateStr & seqNum
    txtOrderNo.Enabled = False ' Make it read-only
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating order number: " & Err.Description, vbCritical
End Sub

Private Sub cmdSave_Click()
    ' Validate required fields
    If Not ValidateFields Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    sql = "INSERT INTO pro_master (ORDER_NO, ORDER_DATE, SUPP_ID, METHOD_OF_PAYMENT, ADVANCE, DUES, NET_AMT) " & _
          "VALUES ('" & txtOrderNo.Text & "', '" & _
          txtOrderDate.Text & "', '" & _
          txtSuppId.Text & "', '" & _
          cboPaymentMethod.Text & "', " & _
          IIf(Trim(txtAdvance.Text) = "", "NULL", txtAdvance.Text) & ", " & _
          IIf(Trim(txtDues.Text) = "", "NULL", txtDues.Text) & ", " & _
          txtNetAmt.Text & ")"
    
    Set R = C.Execute(sql)
    
    MsgBox "Purchase order saved successfully!", vbInformation
    ClearFields
    GenerateOrderNo
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving purchase order: " & Err.Description, vbCritical
End Sub

Private Function ValidateFields() As Boolean
    ' Validate Order No
    If Trim(txtOrderNo.Text) = "" Then
        MsgBox "Order number is required!", vbExclamation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    ' Validate Supplier ID
    If Trim(txtSuppId.Text) = "" Then
        MsgBox "Please enter Supplier ID!", vbExclamation
        txtSuppId.SetFocus
        Exit Function
    End If
    
    ' Validate Payment Method
    If cboPaymentMethod.ListIndex = -1 Then
        MsgBox "Please select Payment Method!", vbExclamation
        cboPaymentMethod.SetFocus
        Exit Function
    End If
    
    ' Validate Net Amount
    If Not IsNumeric(txtNetAmt.Text) Or Val(txtNetAmt.Text) <= 0 Then
        MsgBox "Please enter valid Net Amount!", vbExclamation
        txtNetAmt.SetFocus
        Exit Function
    End If
    
    ' Validate Advance and Dues if entered
    If Trim(txtAdvance.Text) <> "" Then
        If Not IsNumeric(txtAdvance.Text) Then
            MsgBox "Please enter valid Advance amount!", vbExclamation
            txtAdvance.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtDues.Text) <> "" Then
        If Not IsNumeric(txtDues.Text) Then
            MsgBox "Please enter valid Dues amount!", vbExclamation
            txtDues.SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True
End Function

Private Sub ClearFields()
    txtSuppId.Text = ""
    cboPaymentMethod.ListIndex = -1
    txtAdvance.Text = ""
    txtDues.Text = ""
    txtNetAmt.Text = ""
    txtOrderDate.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub

' Add keyboard navigation
Private Sub txtOrderNo_KeyPress(keyascii As Integer)
    If keyascii = 13 Then txtOrderDate.SetFocus
End Sub

Private Sub txtOrderDate_KeyPress(keyascii As Integer)
    If keyascii = 13 Then txtSuppId.SetFocus
End Sub

Private Sub txtSuppId_KeyPress(keyascii As Integer)
    If keyascii = 13 Then cboPaymentMethod.SetFocus
End Sub

Private Sub cboPaymentMethod_KeyPress(keyascii As Integer)
    If keyascii = 13 Then txtAdvance.SetFocus
End Sub

Private Sub txtAdvance_KeyPress(keyascii As Integer)
    If keyascii = 13 Then txtDues.SetFocus
End Sub

Private Sub txtDues_KeyPress(keyascii As Integer)
    If keyascii = 13 Then txtNetAmt.SetFocus
End Sub

Private Sub txtNetAmt_KeyPress(keyascii As Integer)
    If keyascii = 13 Then cmdSave_Click
End Sub
