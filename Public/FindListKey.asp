<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
queryStr=Request("queryString")
response.write "<li onClick=""fill();"">模糊查找:["&queryStr&"]</li>"
If queryStr="" Then
	Response.End()
End IF
set rs=server.createobject("ADODB.Recordset")
	sql="Select ID,SCHOOL From T_SCHOOL WHERE SCHOOL like'%"&queryStr&"%' "
	rs.open sql,conn,1,3
	do while not (rs.eof or err)
		Response.write "<li onClick=""fill('"&rs("SCHOOL")&"','"&rs("ID")&"');"">"&rs("SCHOOL")&"</li>"
	rs.movenext
	loop
	rs.close
	set rs= nothing
appEnd()
%>
