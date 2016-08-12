<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/md5.asp"-->
<% 
'判断用户在线

'Response.write ToUnixTime(Now(),+8)



dim sql,rs
dim username,password,CheckCode
UserName=uCase(trim(Request("UserName")))
PassWord=replace(trim(Request("PassWord")),"'","")

if not(checkIDCard(UserName)) then
	Response.Write "0|身份证号，不正确！"
	Response.End()
end if
if Password="" then
	Response.Write "0|密码不能为空！"
	Response.End()
end if


if CheckLoginTimes(UserName)>5 Then
	Response.Write("0|连续失败5次,锁定用户登录10分钟。")
	Response.End()
End If
conn.begintrans
	password=md5(password)
	set rs=server.createobject("adodb.recordset")
	sql="select * from T_EXAMINEE where PASSWORD='"&password&"' and USER_ID='"&UserName&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		
		Msg="0|用户名或密码错误！"
	else
		if password<>rs("PASSWORD") then

			Msg="0|用户名或密码错误！"
		else
			If rs("LOCKED")=1 Then
				Msg="0|异常：用户被锁定！"
			Else
				'rs("LAST_LOGIN_IP")=Request.ServerVariables("REMOTE_ADDR")
				rs("LAST_LOGIN_TIME")=ToUnixTime(now(),+8)
				'rs("LOGIN_TIMES")=rs("LOGIN_TIMES")+1
				rs.update
				
'				session.Timeout=SessionTimeout
'				session("UserID")	=rs("UserID")
'				session("UserName")	=rs("UserName")
'				session("NickName")	=rs("NickName")
'				session("PassWord")	=rs("PassWord")
'				session("Sex")		=rs("Sex")	
'				session("ComeFrom") =rs("ComeFrom")
				'response.Cookies("login")("PURVIEW")	=rs("PURVIEW")
				response.Cookies("login")("EXAMINEE_USER_ID")=rs("USER_ID")
				response.Cookies("login")("SYS_USER_ID")=rs("USER_ID")
				response.Cookies("login")("USER_NAME")	=rs("REAL_NAME")				
				Msg="1|登录成功！"
			End If
			
			Conn.execute("delete from T_LOGIN_CHECK_IP where  LOGIN_IP='"&UserName&"'")
		end if
	end if
	rs.close
	set rs=nothing
	if Err.Number = 0 Then
		operateLog "examinee_login.html","考生登录"&USER_ID&".IP:"&Request.ServerVariables("REMOTE_ADDR")
		conn.CommitTrans 
	Else
		conn.RollbackTrans
	End If
call CloseConn()
Response.Write Msg



'记录登录IP
Function CheckLoginTimes(UserName)
	Conn.execute("delete from T_LOGIN_CHECK_IP where INFO_TIME<"&ToUnixTime(now(),+8)-300)
	Conn.execute("Insert into T_LOGIN_CHECK_IP (INFO_TIME,LOGIN_IP) values("&ToUnixTime(now(),+8)&",'"&UserName&"')")
	getLogCount=Conn.execute("select Count(LOGIN_IP) from T_LOGIN_CHECK_IP where LOGIN_IP='"&UserName&"'")
	CheckLoginTimes=getLogCount(0)
End Function

%>