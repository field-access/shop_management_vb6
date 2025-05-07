VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form sale_details 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   12795
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   12
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   12795
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   3480
      TabIndex        =   41
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "calculate netamt"
      Height          =   630
      Left            =   7200
      TabIndex        =   39
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   9720
      TabIndex        =   38
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "save order"
      Height          =   630
      Left            =   120
      TabIndex        =   36
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "delete  transaction"
      Height          =   855
      Left            =   4440
      TabIndex        =   35
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "exit"
      Height          =   630
      Left            =   9960
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   34
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   9720
      TabIndex        =   33
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   9720
      TabIndex        =   31
      Top             =   3000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Sales_frm.frx":0000
      Height          =   3495
      Left            =   12840
      TabIndex        =   27
      Top             =   4680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Sales_frm.frx":0015
      Height          =   3135
      Left            =   12840
      TabIndex        =   26
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8640
      Top             =   240
      Visible         =   0   'False
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
      Connect         =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sales"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   240
      Visible         =   0   'False
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
      Connect         =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sales_details"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000014&
      Caption         =   "print invoice"
      Height          =   615
      Left            =   9960
      TabIndex        =   25
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "addnew:"
      Height          =   615
      Left            =   4680
      TabIndex        =   24
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete order"
      Height          =   615
      Left            =   2280
      TabIndex        =   23
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "update All :"
      Height          =   855
      Left            =   2280
      TabIndex        =   22
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save transaction"
      Height          =   615
      Left            =   7200
      TabIndex        =   21
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   9720
      TabIndex        =   19
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   9720
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   390
      Left            =   9720
      TabIndex        =   17
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   390
      Left            =   9720
      TabIndex        =   16
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   510
      Left            =   3480
      TabIndex        =   10
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "pro_name:"
      Height          =   270
      Left            =   720
      TabIndex        =   40
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "dues:"
      Height          =   270
      Left            =   6960
      TabIndex        =   37
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "date:"
      Height          =   270
      Left            =   7080
      TabIndex        =   32
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "total_quantity:"
      Height          =   270
      Left            =   6840
      TabIndex        =   30
      Top             =   3120
      Width           =   1770
   End
   Begin VB.Label Label14 
      Caption         =   "Final_billing:"
      Height          =   375
      Left            =   12960
      TabIndex        =   29
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "products_bought:"
      Height          =   375
      Left            =   12840
      TabIndex        =   28
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label12 
      Caption         =   "Billing Form: sale details:"
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "NetAmount:"
      Height          =   270
      Left            =   6960
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "advance:"
      Height          =   270
      Left            =   6960
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "method_of_payment:"
      Height          =   270
      Left            =   6840
      TabIndex        =   13
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "customer_ID:"
      Height          =   270
      Left            =   7080
      TabIndex        =   12
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "GST:"
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   5280
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "rate:"
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "quantity:"
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   840
      Y2              =   9240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "pro_id:"
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "invoice_no"
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1305
   End
End
Attribute VB_Name = "sale_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Adodc1.Recordset.Update
Adodc2.Recordset.Update
MsgBox "record Updated", vbInformation, "success"
End Sub

Private Sub Command3_Click()
    Adodc1.Recordset.Delete
    MsgBox "Record Deleted", vbInformation, "Success"
    Adodc1.Refresh
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Adodc2.Recordset.Delete
End Sub

Private Sub Command7_Click()
    sql = "INSERT INTO sales_details (inv_no, pro_id, pro_name, qty, rate, gst, amt) VALUES ('" & Trim(Text1.Text) & "', '" & Trim(Combo1.Text) & "', '" & Text11.Text & "', " & IIf(IsNumeric(Text2.Text), Val(Text2.Text), "NULL") & ", " & IIf(IsNumeric(Text3.Text), Val(Text3.Text), "NULL") & ", " & IIf(IsNumeric(Text4.Text), Val(Text4.Text), "NULL") & ", " & IIf(IsNumeric(Text5.Text), Val(Text5.Text), "NULL") & ")"
    C.Execute sql

  ' Update stock quantity in product table
    sql = "UPDATE product SET stock_qty = stock_qty - " & Text2.Text & " WHERE pro_id = '" & Combo1.Text & "'"
    C.Execute sql


    Adodc1.Refresh
    DataGrid1.Refresh
    MsgBox "record updated in sales_details | product-stock ", vbInformation
End Sub



Private Sub Form_Load()
    conn
    Adodc1.Refresh
    DataGrid1.Refresh

    ' Get next invoice number
    sql = "SELECT MAX(inv_no) FROM sales_details"
    Set R = C.Execute(sql)
    If Not R.EOF And Not IsNull(R.Fields(0)) Then
        Text1.Text = Val(R.Fields(0)) + 1
    Else
        Text1.Text = 1
    End If
    R.Close

    ' Load product IDs into Combo1
    Combo1.Clear
    sql = "SELECT pro_id FROM product"
    Set R = C.Execute(sql)
    Do While Not R.EOF
        Combo1.AddItem R.Fields(0)
        R.MoveNext
    Loop
    R.Close

    ' Load customer IDs into Combo2
    Combo2.Clear
    sql = "SELECT cus_id FROM customer"
    Set R = C.Execute(sql)
    Do While Not R.EOF
        Combo2.AddItem R.Fields(0)
        R.MoveNext
    Loop
    R.Close

    ' Load payment methods into Combo3
    Combo3.AddItem "Cash"
    Combo3.AddItem "Online"
    Combo3.AddItem "Card"

    ' Set current date
    Text9.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Command1_Click()

    ' Insert into sales table
    sql = "INSERT INTO sales (inv_no, cus_id, sale_date, method_of_payment, adv, net_amt, total_qty, dues) VALUES ('" & Text1.Text & "', '" & Combo2.Text & "', '" & Text9.Text & "', '" & Combo3.Text & "', " & Text6.Text & ", " & Text8.Text & ", " & Text7.Text & ", " & Text10.Text & ")"
    C.Execute sql

    
    ' Update customer balance (wallet system)
    sql = "UPDATE customer SET cus_bal = cus_bal - " & Text8.Text & " WHERE cus_id = '" & Combo2.Text & "'"
    C.Execute sql

    ' Refresh DataGrid
    Adodc1.Refresh
    DataGrid1.Refresh

    MsgBox "record updated in sales , customer", vbInformation

End Sub

Private Sub Combo2_Click()
    ' Fetch customer balance
    sql = "SELECT cus_bal FROM customer WHERE cus_id = '" & Combo2.Text & "'"
    Set R = C.Execute(sql)

    If Not R.EOF Then
        Text6.Text = R.Fields("cus_bal").Value
    Else
        Text6.Text = "0"
    End If

    R.Close
End Sub

Private Sub Command4_Click()
    ' Clear existing items before adding new ones to prevent duplication
    Combo1.Clear

    ' Reload product IDs (only unique values will be added)
    sql = "SELECT DISTINCT pro_id FROM product"
    Set R = C.Execute(sql)

    Do While Not R.EOF
        Combo1.AddItem R.Fields(0).Value
        R.MoveNext
    Loop
    R.Close  ' Close the recordset
    Set R = Nothing ' Free memory

    ' Clear fields
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text11.Text = ""

    ' Reset ComboBox selections
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1

    ' Restore current date
    Text9.Text = Format(Date, "dd/mm/yyyy")
End Sub


Private Sub Command9_Click()
    DataEnvironment1.Command2 Text1.Text
    DataReport1.Show
    DataEnvironment1.Command1 Text1.Text
    DataReport2.Show
End Sub

Private Sub Combo1_Click()
    ' Fetch product selling rate
    sql = "SELECT pro_sell_rate, pro_name FROM product WHERE pro_id = '" & Combo1.Text & "'"
    Set R = C.Execute(sql)

    If Not R.EOF Then
        Text3.Text = R.Fields("pro_sell_rate").Value
        Text11.Text = R.Fields("pro_name").Value
    Else
        Text3.Text = "0"
    End If

    R.Close
End Sub

Private Sub CalculateNetAmount()
    Dim qty As Double, mrp As Double, gst As Double

    ' Convert text to numeric values safely
    qty = IIf(IsNumeric(Text2.Text), Val(Text2.Text), 0)
    mrp = IIf(IsNumeric(Text3.Text), Val(Text3.Text), 0)
    gst = IIf(IsNumeric(Text4.Text), Val(Text4.Text), 0)

    ' Perform calculation and update net amount as string
    Text5.Text = Format((qty * mrp) + gst, "0.00")
End Sub


Private Sub Text2_Change()
CalculateNetAmount
End Sub

Private Sub Text3_Change()
CalculateNetAmount
End Sub

Private Sub Text4_Change()
CalculateNetAmount
End Sub

Private Sub billingamount()
Dim adv As Double, netamt As Double, qty As Integer, dues As Integer

qty = IIf(IsNumeric(Text7.Text), Val(Text7.Text), 0)
dues = IIf(IsNumeric(Text10.Text), Val(Text10.Text), 0)
adv = IIf(IsNumeric(Text6.Text), Val(Text6.Text), 0)

netamt = Format((dues) - adv, "0.00")
Text8.Text = netamt
End Sub


Private Sub Text6_Change()
billingamount
End Sub

Private Sub Text7_Change()
billingamount
End Sub

Private Sub Text10_Change()
billingamount
End Sub


Private Sub Command8_Click()
Dim qty As Integer
sql = "select sum(qty) from sales_details where inv_no = '" & Text1.Text & "'"
Set R = C.Execute(sql)
Text7.Text = R.Fields(0)
qty = Text7.Text
Set R = Nothing

sql = "select sum(amt) from sales_details where inv_no = '" & Text1.Text & "'"
Set R = C.Execute(sql)
Text10.Text = R.Fields(0)
Set R = Nothing

End Sub


Private Sub DataGrid1_Click()
    ' Load Selected Record from sales_details into Textboxes
    If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
        Text1.Text = Adodc1.Recordset.Fields("inv_no")
        Combo1.Text = Adodc1.Recordset.Fields("pro_id")
        Text11.Text = Adodc1.Recordset.Fields("pro_name")
        Text2.Text = IIf(IsNull(Adodc1.Recordset.Fields("qty")), "", Adodc1.Recordset.Fields("qty"))
        Text3.Text = IIf(IsNull(Adodc1.Recordset.Fields("rate")), "", Adodc1.Recordset.Fields("rate"))
        Text4.Text = IIf(IsNull(Adodc1.Recordset.Fields("gst")), "", Adodc1.Recordset.Fields("gst"))
        Text5.Text = IIf(IsNull(Adodc1.Recordset.Fields("amt")), "", Adodc1.Recordset.Fields("amt"))
    End If
End Sub

Private Sub DataGrid2_Click()
    ' Load Selected Record from sales into Textboxes
    If Not (Adodc2.Recordset.EOF Or Adodc2.Recordset.BOF) Then
        Text1.Text = Adodc2.Recordset.Fields("inv_no")
        Combo2.Text = Adodc2.Recordset.Fields("cus_id")
        Text9.Text = Adodc2.Recordset.Fields("sale_date")
        Combo3.Text = Adodc2.Recordset.Fields("method_of_payment")
        Text6.Text = IIf(IsNull(Adodc2.Recordset.Fields("adv")), "", Adodc2.Recordset.Fields("adv"))
        Text8.Text = IIf(IsNull(Adodc2.Recordset.Fields("net_amt")), "", Adodc2.Recordset.Fields("net_amt"))
        Text7.Text = IIf(IsNull(Adodc2.Recordset.Fields("total_qty")), "", Adodc2.Recordset.Fields("total_qty"))
        Text10.Text = IIf(IsNull(Adodc2.Recordset.Fields("dues")), "", Adodc2.Recordset.Fields("dues"))
    End If
End Sub

Private Sub Text1_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Combo1.SetFocus
    End If
End Sub

Private Sub Combo1_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text4.SetFocus
    End If
End Sub

Private Sub Text4_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text5.SetFocus
    End If
End Sub

Private Sub Text5_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Command7.SetFocus
    End If
End Sub

Private Sub Combo2_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text9.SetFocus
    End If
End Sub

Private Sub Text9_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Combo3.SetFocus
    End If
End Sub

Private Sub Combo3_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text6.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text7.SetFocus
    End If
End Sub

Private Sub Text7_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text8.SetFocus
    End If
End Sub

Private Sub Text8_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Text10.SetFocus
    End If
End Sub

Private Sub Text10_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Command1.SetFocus
    End If
End Sub

