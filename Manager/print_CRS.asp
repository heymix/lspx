<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
window.document.body.innerHTML=bdhtml; 
         }
</script>
<link href="../Css/print_crs.css" rel="stylesheet" type="text/css" />
</head>

<body>
<%
subBox=request("subBox")
If subBox="" Then
	Response.write errStr("没有考生,点击确定关闭！","close")
	appEnd()
End If
%>
<table width="98%" border="0" cellpadding="0" cellspacing="2" align="center">
  <tr> 
    <td height="21" align="center"><input type="button" onClick="preview();//window.close()" value="打印">
    </td>
  </tr>
</table>
<!--startprint-->
<%set rs=conn.execute("select * from T_TEST where ID="&test_id)
	If not(rs.eof or err) Then
		CERT_TIME	= rs("CERT_TIME")
		CHECK_TIME	= Format_Time(FromUnixTime(rs("CHECK_TIME"),+8),4)
	else
		Response.Write "考试错误"
		Response.End()
	End If
	rs.close
	set rs=nothing
%>
<%
user_arr=split(subBox,",")
user_id_arr=""
for i=0 to Ubound(user_arr)
	if user_id_arr="" then
		user_id_arr=user_id_arr&"'"&trim(user_arr(i))&"'"
	else
		user_id_arr=user_id_arr&",'"&trim(user_arr(i))&"'"
	end if
	set rs=conn.execute("select * from V_EXAMINEE_TEST where USER_ID='"&trim(user_arr(i))&"' and TEST_ID="&test_id)
	If rs.eof or err Then
		Response.Write("出错")
		Response.End()
	Else
		REAL_NAME=rs("REAL_NAME")
		GENDER=rs("GENDER")
		CERT_TIME=rs("CERT_TIME")
		CHECK_TIME=FromUnixTime(rs("CHECK_TIME"),+8)
		CRS_NO=rs("CRS_NO")
		PHOTO=rs("PHOTO")
	End If
	rs.close
	set rs=nothing
	set rs=conn.execute("select max(RESULT) as RESULT,TOPIC_ID from V_SCORE where TEST_CATE=12 and RESULT>=60 and USER_ID='"&trim(user_arr(i))&"' group by TOPIC_ID,USER_ID")
	gs=0
	gxl=0
	gfg=0
	gxy=0
	do while not(rs.eof or err)
	If rs("TOPIC_ID")="1" Then
		gs=rs("RESULT")
	End If
	If rs("TOPIC_ID")=2 Then
		gxl=rs("RESULT")
	End If
	If rs("TOPIC_ID")=3 Then
		gfg=rs("RESULT")
	End If
	If rs("TOPIC_ID")=4 Then
		gxy=rs("RESULT")
	End If
	
	rs.movenext
	loop
	rs.close
	set rs=nothing
	%>
<div id="container">
	<div id="user_photo"><img src="..//UpLoadFiles/userPhoto/<%=PHOTO%>" height="160" width="120"></div>
	<div id="p_name"><%=REAL_NAME%></div>
    <div id="gender"><%=getGender(GENDER)%></div>
    <div id="year"><%=CERT_TIME%></div>
    <div id="idcard"><%=trim(user_arr(i))%></div>
    <div id="crsno"><%=CRS_NO%></div>
    <div id="gs"><%=gs%></div>
    <div id="gxl"><%=gxl%></div>
    <div id="gfg"><%=gfg%></div>
    <div id="gxy"><%=gxy%></div>
    <div id="jsj">合格</div>
    <div id="c_year"><%=year(CHECK_TIME)%></div>
    <div id="c_month"><%=month(CHECK_TIME)%></div>
    <div id="c_day"><%=day(CHECK_TIME)%></div>
</div>
<div class="clearPage"></div>
	
<%
	next
%>

<!--endprint-->
</body>
</html>