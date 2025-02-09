VERSION 5.00
Begin VB.Form frmCustomer 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As ADODB.Connection
Dim R As ADODB.Recordset
Dim sql As String
Dim N As String

Private Sub Command1_Click()
sql = "INSERT INTO A VALUES('" + Reg.Text + "','" + Sname.Text + "','" + Photo.Caption + "')"
    Set R = C.Execute(sql)
    MsgBox "RECORD SAVED"
    End Sub
