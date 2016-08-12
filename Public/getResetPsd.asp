<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
RNDPSD	 		 = Request("RND")
PASSWORD	 = Request("PASSWORD")
PASSWORD_R	 = Request("PASSWORD_R")


'response.write "0|"&PURVIEW&"()"&OLD_PASSWORD&"()"&USER_NAME&"()"&PASSWORD
'Call appEnd() 
If RNDPSD=""  Then 
	Response.write "0|参数丢失！请重新操作！"
	Call appEnd()
End If
If PASSWORD="" Then
	Response.write "0|错误密码不能为空！"
	Call appEnd()
End If
If PASSWORD<>PASSWORD_R Then
	Response.write "0|错误两次输入的密码不一样！"
	Call appEnd()
End If
if CheckLoginTimes(RNDPSD)>5 Then
	Response.Write("0|连续失败5次,锁定用户操作10分钟。")
	Response.End()
End If
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXAMINEE Where RNDPSD='"&RNDPSD&"'"
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				USER_ID=rs("USER_ID")
				rs("PASSWORD")=md5(PASSWORD)
			rs.update
			rs.close
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "getpsd.asp","修改密码：用户："&USER_ID
				Conn.execute("delete from T_LOGIN_CHECK_IP where  LOGIN_IP='"&RNDPSD&"'")
				Response.write "1|修改成功！请重新登录！"
				conn.CommitTrans 
			Else
				Response.write "0|修改失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|修改失败！操作过期，请重新发送邮件！."
			conn.RollbackTrans
		End If
call CloseConn()

'记录登录IP
Function CheckLoginTimes(RNDPSD)
	Conn.execute("delete from T_LOGIN_CHECK_IP where INFO_TIME<"&ToUnixTime(now(),+8)-300)
	Conn.execute("Insert into T_LOGIN_CHECK_IP (INFO_TIME,LOGIN_IP) values("&ToUnixTime(now(),+8)&",'"&RNDPSD&"')")
	getLogCount=Conn.execute("select Count(LOGIN_IP) from T_LOGIN_CHECK_IP where LOGIN_IP='"&RNDPSD&"'")
	CheckLoginTimes=getLogCount(0)
End Function
%>
