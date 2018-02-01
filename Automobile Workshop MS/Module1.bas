Attribute VB_Name = "Module1"
Public adocn As New ADODB.Connection
Public mech_id As Integer
Public VECH_ID As Integer
Public J_ID As Integer
Public pw As String
Public unam As String

Public Sub main()
adocn.ConnectionString = "Provider=SQLNCLI.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AWSM;Data Source=HP-PC"
adocn.CursorLocation = adUseClient
adocn.Open
frm_login.Show
End Sub
