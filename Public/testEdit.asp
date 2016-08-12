<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
operateType	 = Request("operateType")
TITLE	 	 = Request("TITLE")
START_TIME	 = Request("START_TIME")
END_TIME	 = Request("END_TIME")
CHECKED	 	 = Request("CHECKED")
NOTICE_ID	 = Request("NOTICE_ID")
TEST_CATE	 = Request("TEST_CATE")
CERT_TIME	 = Request("CERT_TIME")
CHECK_TIME	 = Request("CHECK_TIME")
CERT_NO	 	 = Request("CERT_NO")

'response.write "0|"&PURVIEW&"()"&operateType&"()"&USER_NAME&"()"&PASSWORD
'Call appEnd() 
If PURVIEW="" or  operateType="" or TITLE="" or START_TIME="" or END_TIME="" or TEST_CATE="0" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

'If sqlInjection(opValue) Then
'	Response.write "0|操作失败,名称包含敏感信息！"
'	Call appEnd()
'End If

If PURVIEW<>"0" Then
	Response.write "0|操作失败,没有权限！"
	Call appEnd()
End If
If CHECKED<>"1" Then CHECKED="0"
If operateType=0 Then

		conn.begintrans
		autoID=getID("T_TEST")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST Where TITLE='"&TITLE&"'"
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID") 	= autoID
				rs("TITLE") = TITLE
				rs("START_TIME") = ToUnixTime(START_TIME&" 0:01",+8)
				rs("END_TIME")	 = ToUnixTime(END_TIME&" 59:59",+8)
				rs("NOTICE_ID")	 = NOTICE_ID
				rs("TEST_CATE")	 = TEST_CATE
				rs("CHECKED")	 = CHECKED
				rs("INFO_TIME")	 = ToUnixTime(now(),+8)
				rs("CERT_TIME")	 = CERT_TIME
				rs("CHECK_TIME") = ToUnixTime(CHECK_TIME&" 0:00",+8)
				rs("CERT_NO")	 = CERT_NO
			rs.update
			rs.close
			conn.execute("insert into T_TEST_EXA_RELATION (TEST_ID,EXA_ID)select "&autoID&",ID from T_EXA where CHECKED=1 and TEST_CATE="&TEST_CATE)
			conn.execute("insert into T_TEST_TOPIC_RELATION (TEST_ID,TOPIC_ID)select "&autoID&",ID from T_TEST_TOPIC where TEST_CATE="&TEST_CATE)
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "test_add.asp","添加：系统编号："&autoID&" 标题："&TITLE
				Response.write "1|添加成功！"
				conn.CommitTrans 
			Else
				Response.write "0|添加失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|添加失败！记录已经存在."
			conn.RollbackTrans
		End If
ElseIf operateType=1 Then
		If not is_Id(ID) Then
			Response.write "0|操作失败,系统编号错误！"
			Call appEnd()
		End If

		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST Where ID="&ID
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("TITLE")=TITLE
				rs("START_TIME")=ToUnixTime(START_TIME&" 0:01",+8)
				rs("END_TIME")=ToUnixTime(END_TIME&" 23:59",+8)
				rs("NOTICE_ID")=NOTICE_ID
				rs("TEST_CATE")=TEST_CATE
				rs("CHECKED")=CHECKED
				rs("CERT_TIME")	 = CERT_TIME
				rs("CHECK_TIME") = ToUnixTime(CHECK_TIME&" 0:00",+8)
				rs("CERT_NO")	 = CERT_NO
			rs.update
			rs.close
		
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "test_add.asp","修改：系统编号："&id&" 标题："&TITLE
				Response.write "1|修改成功！"
				conn.CommitTrans 
			Else
				Response.write "0|修改失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|修改失败！记录不存在."
			conn.RollbackTrans
		End If
Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
