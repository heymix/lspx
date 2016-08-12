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
	sql="Select S.ID,S.SCHOOL From T_SCHOOL AS S RIGHT JOIN T_TEST_SCHOOL_RELATION AS R ON S.ID=R.SCHOOL_ID WHERE S.SCHOOL like'%"&queryStr&"%' And R.SCHOOL_ID not in(select SCHOOL_ID From T_EXA_SCHOOL_RELATION where TEST_ID="&TEST_ID&") And R.SCHOOL_ID not in(select SCHOOL_ID from T_EXA where TEST_ID="&TEST_ID&") And TEST_ID="&test_id
	rs.open sql,conn,1,3
	do while not (rs.eof or err)
		Response.write "<li onClick=""fill('"&rs("SCHOOL")&"','"&rs("ID")&"');"">"&rs("SCHOOL")&"</li>"
	rs.movenext
	loop
	rs.close
	set rs= nothing
appEnd()
%>
