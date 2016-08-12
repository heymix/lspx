<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckExamineeOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
	   operateType=trim(Request("operateType"))
		REAL_NAME	=trim(Request("REAL_NAME"))
		EDUCATION	=trim(Request("EDUCATION"))
		SCHOOL_ID	=trim(Request("SCHOOL_ID"))
		TEL			=trim(Request("TEL"))
		MOBILE		=trim(Request("MOBILE"))
		MOBILE_BAK	=trim(Request("MOBILE_BAK"))
		EMAIL		=trim(Request("EMAIL"))
		GRADUATE_SCHOOL=trim(Request("GRADUATE_SCHOOL"))
		GRADUATE_DATE=trim(Request("GRADUATE_DATE"))
		PRO			=trim(Request("PRO"))
		PHOTO		=trim(Request("PHOTO"))

If  operateType="" or EXAMINEE_USER_ID="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

If REAL_NAME="" Then
	Response.write "0|姓名不能为空！"
	Call appEnd()
End If

If EDUCATION="" Then
	Response.write "0|请选择学历！"
	Call appEnd()
End If

If SCHOOL_ID="-1" Then
	Response.write "0|选择所在学校！"
	Call appEnd()
End If

If not(TEL<>""or MOBILE<>"") Then
	Response.write "0|电话不能为空！"
	Call appEnd()
End If
If EMAIL="" Then
	Response.write "0|邮件不能为空！"
	Call appEnd()
End If
If GRADUATE_SCHOOL="" Then
	Response.write "0|毕业学校不能为空！"
	Call appEnd()
End If
If GRADUATE_DATE="" Then
	Response.write "0|毕业日期不能为空！"
	Call appEnd()
End If

If PRO="" Then
	Response.write "0|专业不能为空！"
	Call appEnd()
End If

If PHOTO="" Then
	Response.write "0|照片不能为空！"
	Call appEnd()
End If
If operateType=1 Then
		If not checkIDCard(EXAMINEE_USER_ID) Then
			Response.write "0|操作失败,注册的身份证号码有误！"
			Call appEnd()
		End If
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXAMINEE Where USER_ID='"&EXAMINEE_USER_ID&"'"
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("REAL_NAME")=REAL_NAME
				rs("AGE")=IDentity(EXAMINEE_USER_ID,"age")
				rs("EDUCATION")=EDUCATION
				rs("TEL")=TEL
				rs("SCHOOL_ID")=SCHOOL_ID
				rs("MOBILE")=MOBILE
				rs("MOBILE_BAK")=MOBILE_BAK
				rs("EMAIL")=EMAIL
				rs("GRADUATE_SCHOOL")=GRADUATE_SCHOOL
				rs("GRADUATE_DATE")=ToUnixTime(strToDate(GRADUATE_DATE,1),+8)
				rs("PRO")=PRO
				rs("PHOTO")=PHOTO
			rs.update
			rs.close
		
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "examinee_myset.asp","对资料进行了修改主。 by 用户 身份证号："&EXAMINEE_USER_ID&" ："&REAL_NAME
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
	Response.write "0|失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
