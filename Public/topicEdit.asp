<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

TNAME = Request("TNAME")
TEST_CATE = Request("TEST_CATE")
IS_PRATICE = Request("IS_PRATICE")
operateType = Request("operateType")
ID = Request("ID")
'response.write "0|"&IS_PRATICE&TEST_CATE&ID
'response.End()
If IS_PRATICE<>"1" Then IS_PRATICE="0"
If PURVIEW="" or  operateType="" or TNAME="" or not is_Id(TEST_CATE) or TEST_CATE="0" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

If sqlInjection(TNAME) Then
	Response.write "0|操作失败,名称包含敏感信息！"
	Call appEnd()
End If

If PURVIEW<>"0" Then
	Response.write "0|操作失败,没有权限！"
	Call appEnd()
End If
If operateType="0" Then
		conn.begintrans
		autoID=getID("T_TEST_TOPIC")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST_TOPIC Where TNAME='"&TNAME&"'"
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID")=autoID
				rs("TNAME")=TNAME
				rs("TEST_CATE")=TEST_CATE
				rs("IS_PRATICE")=IS_PRATICE
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "topic.asp","添加考试科目："&autoID&" "&TNAME
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
ElseIf operateType="1" Then
		If not is_Id(ID) Then
			Response.write "0|修改失败,操作主键丢失"
			Call appEnd()
		End If
		rs_same=conn.Execute("select Count(*) From T_TEST_TOPIC Where TNAME='"&TNAME&"' And ID<>"&ID)(0)
		If rs_same>0 Then
			Response.write "0|修改失败,记录重复！"
			Call appEnd()
		End If
			conn.begintrans
			set rs=server.createobject("ADODB.Recordset")
			sql="select * From T_TEST_TOPIC Where ID="&ID
			rs.open sql,conn,1,3
			If not(rs.eof or rs.bof) Then
				oldName=rs("TNAME")
				rs("TNAME")=TNAME
				rs("TEST_CATE")=TEST_CATE
				rs("IS_PRATICE")=IS_PRATICE
				rs.update
				rs.close
				set rs=nothing
				if Err.Number = 0 Then
					operateLog "topic.asp","修改考试科目："&opID&" "&oldName&"->"&TNAME
					Response.write "1|修改成功！"
					conn.CommitTrans 
				Else
					Response.write "0|修改失败！"
					conn.RollbackTrans
				End If
			Else
				Response.write "0|修改失败！记录记录不存在."
				conn.RollbackTrans
			End If
Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
