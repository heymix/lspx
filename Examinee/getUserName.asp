<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
UserName=request.Cookies("login")("USER_NAME")
If UserName="" Then
	Response.write "异常"
Else
	Response.write UserName
End If

%>