<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
	
	   operateType=trim(Request("operateType"))
	   	USER_ID=uCase(trim(Request("USER_ID")))
		PASSWORD	 = trim(Request("PASSWORD"))
		PASSWORD_R	 = trim(Request("PASSWORD_R"))
		REAL_NAME	=trim(Request("REAL_NAME"))
		EDUCATION	=trim(Request("EDUCATION"))
		SCHOOL_ID	=trim(Request("SCHOOL_ID"))
		TEL			=trim(Request("TEL"))
		MOBILE		=trim(Request("MOBILE"))
		EMAIL		=trim(Request("EMAIL"))
		GRADUATE_SCHOOL=trim(Request("GRADUATE_SCHOOL"))
		GRADUATE_DATE=trim(Request("GRADUATE_DATE"))
		PRO			=trim(Request("PRO"))
		PHOTO		=trim(Request("PHOTO"))

If  operateType="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

If not(checkIDCard(USER_ID)) Then
	Response.write "0|身份证号码有误！"
	Call appEnd()
End If

If(len(PASSWORD)<6 or len(PASSWORD)>12) Then
	Response.write "0|密码6-12位！"
	Call appEnd()
End If

If PASSWORD<>PASSWORD_R Then
	Response.write "0|两次输入的密码不一样！"
	Call appEnd()
End If


If SCHOOL_ID="-1" or SCHOOL_ID=""Then
	Response.write "0|请选择学校！"
	Call appEnd()
End If

If REAL_NAME="" Then
	Response.write "0|姓名不能为空！"
	Call appEnd()
End If

If  len(MOBILE)<>11 Then
	Response.write "0|手机号码不正确！"
	Call appEnd()
End If
If EMAIL="" Then
	Response.write "0|邮件不能为空！"
	Call appEnd()
End If

If operateType=0 Then
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXAMINEE Where USER_ID='"&USER_ID&"'"
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("USER_ID")=USER_ID
				rs("REAL_NAME")=REAL_NAME
				rs("PASSWORD")=md5(PASSWORD)
				rs("SCHOOL_ID")=SCHOOL_ID
				rs("AGE")=IDentity(USER_ID,"age")
				rs("GENDER")=IDentity(USER_ID,"gender")
				rs("BIRTHDAY")=ToUnixTime(IDentity(USER_ID,"birthday")&" 08:00",+8)
				rs("MOBILE")=MOBILE
				rs("EMAIL")=EMAIL
				rs("INFO_TIME")=ToUnixTime(now(),+8)
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "examinee_reg.html","添加：身份证号："&USER_ID&" 姓名："&REAL_NAME
				Response.write "1|添加成功！"
				conn.CommitTrans 
			Else
				Response.write "0|添加失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|添加失败！身份证号已经被注册."
			conn.RollbackTrans
		End If
Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
