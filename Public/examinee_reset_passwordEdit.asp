﻿<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckExamineeOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
USER_NAME	 = Request("USER_NAME")
PASSWORD	 = Request("PASSWORD")
OLD_PASSWORD = Request("OLD_PASSWORD")

'response.write "0|"&PURVIEW&"()"&OLD_PASSWORD&"()"&USER_NAME&"()"&PASSWORD
'Call appEnd() 
If EXAMINEE_USER_ID="" or PASSWORD="" or OLD_PASSWORD="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

'theSame=conn.execute("select count(*) from T_ADMIN Where ID<>"&ID&" AND USER_NAME='"&USER_NAME&"'")(0)
'If theSame>0 Then
'	Response.write "0|操作失败,用户名重名！"
'	Call appEnd()
'End If
	

conn.begintrans
set rs=server.createobject("ADODB.Recordset")
sql="select * From T_EXAMINEE Where  USER_ID='"&ID&"' AND PASSWORD='"&md5(OLD_PASSWORD)&"'"
rs.open sql,conn,1,3
If not(rs.eof or rs.bof) Then
		rs("PASSWORD")=md5(PASSWORD)
	rs.update
	rs.close
	set rs= nothing
	if Err.Number = 0 Then
		operateLog "reset_password.asp","修改："&l_Sort&" 身份证号："&ID
		Response.write "1|修改成功！请重新登录！"
		response.Cookies("login")=""
		conn.CommitTrans 
	Else
		Response.write "0|修改失败！"
		conn.RollbackTrans
	End If
Else
	Response.write  "0|操作失败,密码错误！"
	conn.RollbackTrans
End If
call CloseConn()
%>
