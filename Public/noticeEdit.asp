<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
CATEGORY	 = Request("CATEGORY")
TITLE	 	= Request("TITLE")
NOTICE	 	= Request("CONTENT")
CHECKED	 	= Request("CHECKED")
IS_TOP	 	= Request("IS_TOP")
IS_READ	 	= Request("IS_READ")
operateType	= Request("operateType")
'response.write "0|"&PURVIEW&"()"&operateType&"()"&CATEGORY&"()"&NOTICE
'Call appEnd() 
If PURVIEW="" or operateType="" or  CATEGORY="" or NOTICE=""  or TITLE="" Then 
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
If IS_TOP<>"1" Then IS_TOP="0"
If CHECKED<>"1" Then CHECKED="0"
If operateType=0 Then
		conn.begintrans
		autoID=getID("T_NOTICE")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_NOTICE Where 1=2"
		rs.open sql,conn,1,3
			rs.AddNew
				rs("ID")		= autoID
				rs("CATEGORY")	= CATEGORY
				rs("TITLE")		= TITLE
				rs("NOTICE")	= NOTICE
				rs("CHECKED")	= CHECKED
				rs("IS_TOP")	= IS_TOP
				rs("IS_READ")	= IS_READ
				rs("IS_TOP_TIME")= ToUnixTime(now(),+8)
				rs("INFO_TIME")	= ToUnixTime(now(),+8)
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "notice_add.asp","添加：类别为："&getDatakey(CATEGORY)&" 编号："&autoID&" 标题："&TITLE
				Response.write "1|添加成功！"
				conn.CommitTrans 
			Else
				Response.write "0|添加失败！"
				conn.RollbackTrans
			End If
ElseIf operateType=1 Then
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_NOTICE Where ID="&ID
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("CATEGORY")	= CATEGORY
				rs("TITLE")		= TITLE
				rs("NOTICE")	= NOTICE
				rs("CHECKED")	= CHECKED
				rs("IS_TOP")	= IS_TOP
				rs("IS_READ")	= IS_READ
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "notice_add.asp","修改：类别为："&getDatakey(CATEGORY)&" 编号："&ID&" 标题："&TITLE
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

Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
