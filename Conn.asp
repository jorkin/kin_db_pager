<%
Dim oConn
Public Function OpenConn()
	On Error Resume Next
	If oConn<>Empty Then Exit Function
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Kin_Article.mdb") & ";"
	oConn.CommandTimeout = 30
	oConn.ConnectionTimeout = 30
	oConn.Open()
End Function
OpenConn()
%>