Attribute VB_Name = "Module1"
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public sql As String

Public Function conn()
    Set C = New ADODB.Connection
    C.Open "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False"
    Set R = New ADODB.Recordset
End Function

Public Sub A()
    RES = MsgBox("DO YOU WANT EXIT", vbQuestion + vbYesNoCancel, "FOR EXIT")
    If RES = vbYes Then
        End
    End If
End Sub

