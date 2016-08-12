<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
EXA_DATE	 = Request("EXA_DATE")
EXA_TIME	 = Request("EXA_TIME")


If PURVIEW="" or ID="" or  EXA_DATE="" or EXA_TIME=""  Then 
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
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST_TOPIC_RELATION Where TOPIC_ID="&ID&" and TEST_ID="&test_id
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("EXA_DATE")	= ToUnixTime(EXA_DATE&" 0:01",+8)
				rs("EXA_TIME")	= EXA_TIME
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "current_topic_reEdit.asp","修改：考试："&test_id&","&test_title&"科目："&ID
				Response.write "1|修改成功！"
				conn.CommitTrans 
			Else
				Response.write "0|修改失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|修改失败！未查到记录."
			conn.RollbackTrans
		End If

call CloseConn()
%>
