<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
USER_NAME	 = Request("USER_NAME")
PASSWORD	 = Request("PASSWORD")
operateType	 = Request("operateType")
rMark		 = Request("rMark")
opID		 = Request("sID")
schoolID	 = Request("schoolID")
examID		 = Request("examID")
l_Name		 =""
l_Sort		 ="考试中心管理员"
'response.write "0|"&PURVIEW&"()"&operateType&"()"&USER_NAME&"()"&PASSWORD
'Call appEnd() 
If PURVIEW="" or  operateType="" or USER_NAME="" or PASSWORD="" Then 
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
If rMark=2 Then
	If schoolID="" or Not Isnumeric(schoolID) Then
		Response.write "0|操作失败,学校参数错误！"
		Call appEnd()
	End If
	l_Name=getSchool(schoolID)
	l_Sort="学校管理员"
End If
If rMark=3 Then
	If examID="" or Not Isnumeric(examID) Then
		Response.write "0|操作失败,学校参数错误！"
		Call appEnd()
	End If
	l_Name=getExam(examID)
	l_Sort="考点管理员"	
End If
If CHECKED<>"1" Then CHECKED="0"
If operateType=0 Then
		conn.begintrans
		autoID=getID("T_ADMIN")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_ADMIN Where USER_NAME='"&USER_NAME&"'"
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID")=autoID
				If rMark=2 Then rs("SCHOOL_ID")=schoolID
				If rMark=3 Then rs("EXA_ID")=examID
				rs("USER_NAME")=USER_NAME
				rs("PASSWORD")=md5(PASSWORD)
				rs("PURVIEW")=rMark
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "manager_add.asp","添加："&l_Sort&" 系统编号："&autoID&" 用户名："&USER_NAME&" 名称："&l_Name
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

Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
