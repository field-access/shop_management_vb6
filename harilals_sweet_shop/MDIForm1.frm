VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000004&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.PictureBox Picture2 
         Height          =   2295
         Left            =   6480
         Picture         =   "MDIForm1.frx":0000
         ScaleHeight     =   2235
         ScaleWidth      =   6195
         TabIndex        =   1
         Top             =   0
         Width           =   6255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "sweet shop management"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   420
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   4650
      End
   End
   Begin VB.Menu CUSTOMER 
      Caption         =   "CUSTOMER"
   End
   Begin VB.Menu PRODUCT 
      Caption         =   "PRODUCT"
   End
   Begin VB.Menu RAW 
      Caption         =   "RAW MATERIAL"
   End
   Begin VB.Menu SUPPLIER 
      Caption         =   "SUPPLIER"
   End
   Begin VB.Menu PURCHASE 
      Caption         =   "PURCHASE ENTRY"
   End
   Begin VB.Menu SALES 
      Caption         =   "SALES ENTRY"
   End
   Begin VB.Menu exits 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CUSTOMER_Click()
Customer_management.Show
End Sub

Private Sub EMP_Click()

End Sub

Private Sub exits_Click()
End
End Sub


Private Sub PRODUCT_Click()
Product_management.Show
End Sub

Private Sub PURCHASE_Click()
purchase_frm.Show
End Sub

Private Sub RAW_Click()
Raw_material_management.Show
End Sub

Private Sub SALES_Click()
sale_details.Show
End Sub


Private Sub SUPPLIER_Click()
Supplier_management.Show
End Sub
