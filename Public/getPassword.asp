<!--#include file="../inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/Jmail.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%

UserName=uCase(trim(Request("UserName")))
EMAIL=replace(trim(Request("EMAIL")),"'","")

if not(checkIDCard(UserName)) then
	Response.Write "0|身份证号，不正确！"
	Response.End()
end if
If EMAIL="" Then
	Response.write "0|邮件不能为空！"
	Call appEnd()
End If

if CheckLoginTimes(UserName)>5 Then
	Response.Write("0|连续失败5次,锁定用户操作10分钟。")
	Response.End()
End If

	set rs=server.createobject("ADODB.Recordset")
	sql="select * From T_EXAMINEE Where USER_ID='"&UserName&"' And EMAIL='"&EMAIL&"'"
	rs.open sql,conn,1,3
	If rs.eof or rs.bof Then
		response.write "0|错误，身份证号和邮箱不匹配！"
	Else
	
	getCheck=md5(USERNAME&rs("LAST_LOGIN_TIME"))
	rs("RNDPSD")=getCheck
	rs.update
		E_Server = "smtp.163.com" '发件服务器
		E_ServerUser = "testcenter71@163.com" '登录用户名
		E_ServerPass = "yue123456" '登录密码
		E_SendManMail = "testcenter71@163.com" '发件人邮件地址
		E_SendManName = "考试中心" '发件人姓名
		E_Title = "辽宁省高等学校师资培训考试-密码取回"
		E_Content ="您好！<br /><br />取回密码请，<a target=""_blank"" href=""http://lngspx.lnnu.edu.cn/getpsd.asp?pd="&getCheck&""">请点击这里</a><br /><br />如果无法点击上面的链接，请复制以下网址，并粘帖在浏览器的地址栏中访问。<br /><a target=""_blank"" http://lngspx.lnnu.edu.cn/getpsd.asp?pd="&getCheck&""">http://lngspx.lnnu.edu.cn/getpsd.asp?pd="&getCheck&"</a><br /><br />祝您使用愉快！<br /><br />--辽宁省高等学校师资培训考试中心"
		SendEmail EMAIL,E_Title,E_Content,"Jmail"
		Conn.execute("delete from T_LOGIN_CHECK_IP where  LOGIN_IP='"&UserName&"'")
		response.write "1|邮件发送成功！请去邮箱查收并修改密码！"
		
	End If
	rs.close
	set rs=nothing

	
'记录登录IP
Function CheckLoginTimes(UserName)

	Conn.execute("delete from T_LOGIN_CHECK_IP where INFO_TIME<"&ToUnixTime(now(),+8)-300)
	Conn.execute("Insert into T_LOGIN_CHECK_IP (INFO_TIME,LOGIN_IP) values("&ToUnixTime(now(),+8)&",'"&UserName&"')")
	getLogCount=Conn.execute("select Count(LOGIN_IP) from T_LOGIN_CHECK_IP where LOGIN_IP='"&UserName&"'")
	CheckLoginTimes=getLogCount(0)
End Function
appEnd()
%>