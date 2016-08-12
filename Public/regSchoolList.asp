<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<% 
'判断用户在线

'Response.write ToUnixTime(Now(),+8)

	nContent=""
	set rs=server.createobject("adodb.recordset")
	sql="select ID,SCHOOL from T_SCHOOL where checked=1 ORDER BY SCHOOL ASC"
	rs.open sql,conn,1,1
	do while not(rs.eof or err)
		nContent=nContent&"<option value='"&rs("ID")&"'>"&rs("SCHOOL")&"</option>"
	rs.movenext
	loop
	rs.close
	set rs=nothing
	response.write  nContent

appEnd()%>