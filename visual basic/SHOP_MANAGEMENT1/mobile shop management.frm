VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.Image Image2 
         Height          =   960
         Left            =   11400
         Picture         =   "mobile shop management.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "electronic shop management"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   0
         Width           =   8205
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   1560
         Picture         =   "mobile shop management.frx":1084A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Menu mnuProducts 
      Caption         =   "PRODUCTS"
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "CUSTOMERS"
   End
   Begin VB.Menu mnuSuppliers 
      Caption         =   "SUPPLIERS"
   End
   Begin VB.Menu mnuSales 
      Caption         =   "SALES"
   End
   Begin VB.Menu mnuPurchase 
      Caption         =   "PURCHASE"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "REPORTS"
   End
   Begin VB.Menu exit 
      Caption         =   "exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
End
End Sub

Private Sub MDIForm_Load()
    Me.WindowState = 2  'Maximized
    Me.Caption = "Electronic Shop Management System"
End Sub

Private Sub mnuProducts_Click()
    frmProduct.Show
End Sub

Private Sub mnuCustomers_Click()
    frmCustomer.Show
End Sub

Private Sub mnuSuppliers_Click()
    frmSupplier.Show
End Sub

Private Sub mnuSales_Click()
    frmSales.Show
End Sub

Private Sub mnuPurchase_Click()
    frmPurchase.Show
End Sub

Private Sub mnuReports_Click()
    frmReports.Show
End Sub
