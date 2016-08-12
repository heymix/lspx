<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

opValue = Request("opValue")
operateType = Request("operateType")
opID = Request("sID")

If PURVIEW="" or  operateType="" or opValue="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

If sqlInjection(opValue) Then
	Response.write "0|操作失败,名称包含敏感信息！"
	Call appEnd()
End If

If PURVIEW<>"0" Then
	Response.write "0|操作失败,没有权限！"
	Call appEnd()
End If

If operateType=2 Then
		conn.begintrans
		autoID=getID("T_SCHOOL")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_SCHOOL Where School='"&opValue&"'"
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID")=autoID
				rs("SCHOOL")=opValue
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "school.asp","添加学校："&autoID&" "&opValue
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
		If opID="" or not IsNumeric(opID) Then
			Response.write "0|修改失败,操作主键丢失"
			Call appEnd()
		End If
		rs_same=conn.Execute("select Count(*) From T_SCHOOL Where School='"&opValue&"' And ID<>"&opID)(0)
		If rs_same>0 Then
			Response.write "0|修改失败,记录重复！"
			Call appEnd()
		End If
			conn.begintrans
			set rs=server.createobject("ADODB.Recordset")
			sql="select * From T_SCHOOL Where ID="&opID
			rs.open sql,conn,1,3
			If not(rs.eof or rs.bof) Then
				oldName=rs("SCHOOL")
				rs("SCHOOL")=opValue
				rs.update
				rs.close
				set rs=nothing
				if Err.Number = 0 Then
					operateLog "school.asp","修改学校："&opID&" "&oldName&"->"&opValue
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
