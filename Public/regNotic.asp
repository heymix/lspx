<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<% 
'判断用户在线

'Response.write ToUnixTime(Now(),+8)

	nContent="none"
	set rs=server.createobject("adodb.recordset")
	sql="select NOTICE from T_NOTICE where IS_TOP=1 order by IS_TOP_TIME desc"
	rs.open sql,conn,1,1
	If not(rs.eof or rs.bof) Then
		nContent=rs("NOTICE")
	End If
	rs.close
	set rs=nothing
	response.write  nContent
%>