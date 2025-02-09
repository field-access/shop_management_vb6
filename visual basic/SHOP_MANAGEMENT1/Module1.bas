Attribute VB_Name = "Module1"
' Module1.bas (Simplified - SECURITY REMOVED for easy testing)

Public C As ADODB.Connection
Public R As ADODB.Recordset
Public sql As String

Public Sub conn()
    Set C = New ADODB.Connection
    C.Open "Provider=MSDAORA.1;User ID=sumit/singh;Persist Security Info=False" ' HARDCODED CREDENTIALS - **REMOVE FOR PRODUCTION**
    Set R = New ADODB.Recordset
End Sub

Public Sub PromptExitApplication()
    Dim RES As VbMsgBoxResult
    RES = MsgBox("DO YOU WANT TO EXIT?", vbQuestion + vbYesNoCancel, "Exit Application")
    If RES = vbYes Then
        End
    End If
End Sub
