Attribute VB_Name = "Module1"
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public sql As String

Public Function conn()
    Set C = New ADODB.Connection
    C.Open "Provider=MSDAORA.1;User ID=elon/musk;Persist Security Info=true"
    Set R = New ADODB.Recordset
End Function



