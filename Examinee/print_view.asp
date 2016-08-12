<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckExamineeOL.asp"-->
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
<link href="../Css/print.css" rel="stylesheet" type="text/css" />
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
<%
user_arr=split(subBox,",")
user_id_arr=""
for i=0 to Ubound(user_arr)
	if user_id_arr="" then
		user_id_arr=user_id_arr&"'"&trim(user_arr(i))&"'"
	else
		user_id_arr=user_id_arr&",'"&trim(user_arr(i))&"'"
	end if
next
%>
<div id="container">
<%
'取得考试科目 并按时间排序
max_date=0
min_date=0
topicStr=""
date_temp=0
set rs_topic=conn.execute("SELECT TEST_ID,TOPIC_ID,EXA_DATE,EXA_TIME,TNAME FROM V_CURRENT_TOPIC where IS_PRATICE=0 and TEST_ID="&test_id&" order by EXA_DATE asc,EXA_TIME asc")
	do while not (rs_topic.eof or err)
	 If max_date=0 Then max_date=rs_topic("EXA_DATE")
	 If min_date=0 Then min_date=rs_topic("EXA_DATE")
	 If max_date<rs_topic("EXA_DATE") Then max_date=rs_topic("EXA_DATE")
	 If min_date>rs_topic("EXA_DATE") Then min_date=rs_topic("EXA_DATE")
	 If date_temp<>rs_topic("EXA_DATE") Then 
	 	topicStr=topicStr&"考试日期:<font>"&Format_Time(FromUnixTime(rs_topic("EXA_DATE"),+8),4)&"</font><br>"
		date_temp=rs_topic("EXA_DATE")
	End If
	 topicStr=topicStr&rs_topic("TNAME")&":"&rs_topic("EXA_TIME")&"<br>"
	rs_topic.movenext
	loop
	rs_topic.close
	set rs_topic=nothing
	If max_date<>min_date Then
		exam_date=Format_Time(FromUnixTime(min_date,+8),4)&"-"&mid(Format_Time(FromUnixTime(max_date,+8),4),9,10)
	Else
		exam_date=Format_Time(FromUnixTime(max_date,+8),4)
	End If
	

set rs=conn.execute("SELECT USER_ID,SCHOOL_ID,SCHOOL,GENDER,REAL_NAME,AGE,EDUCATION,TEL,MOBILE,EMAIL,LAST_LOGIN_TIME,CHECKED,LOCKED,IS_MAKEUP,TRCHECKED,PHOTO from V_EXAMINEE_TEST where USER_ID in("&user_id_arr&") and test_id="&test_id&" order by USER_ID desc")
i=0
do while not(rs.eof or err)
	set rs_exam=conn.execute("SELECT ROOM_ID,T_ROOM,EXA_ID,EXA_NAME,TEST_ID,USER_ID,SEAT_ID,T_ROOM_ORDER FROM V_ROOM_SEAT where USER_ID='"&rs("USER_ID")&"' and  test_id="&test_id&"")
	If not(rs_exam.eof or rs_exam.bof) Then
		T_ROOM=rs_exam("T_ROOM")
		EXA_ID=rs_exam("EXA_ID")
		EXA_NAME=rs_exam("EXA_NAME")
		SEAT_ID=rs_exam("SEAT_ID")
		T_ROOM_ORDER=rs_exam("T_ROOM_ORDER")
	End If
	rs_exam.close
	set rs_exam=nothing
	PHOTO=rs("PHOTO")	
		If PHOTO="" or isNULL(PHOTO) Then
			PHOTO_show="../Images/nonepic.jpg"
		Else
			PHOTO_show="../UpLoadFiles/userPhoto/"&PHOTO
		End If
		
		'取得注意事项
	set rs_sThing=conn.execute("select NOTICE From T_NOTICE where id in(select NOTICE_ID from T_TEST where id="&test_id&" )")
	If not (rs_sThing.eof or rs_sThing.bof) Then
		sThing=rs_sThing("NOTICE")
	End If
	rs_sThing.close
	set rs_sThing=nothing
%>
<div class="leftDiv">
    <h1><span><%=test_title%></span></h1>
    <h2>准考证</h2>
    <div class="leftPhoto"><img src="<%=PHOTO_show%>" width="120" height="167" /></div>
    <div class="rightDesc">
    姓名： <%=rs("REAL_NAME")%><br>
    性别： <%=getGender(rs("GENDER"))%><br>
    年龄： <%=rs("AGE")%><br>
    </div>
    <div class="bottomMain">
    	身份证号：<%=rs("USER_ID")%><br>
        准考证号：<%=getTestNo(rs("USER_ID"))%><br>
        主考学校：<%=numberCover(EXA_ID)%>  &nbsp;<%=EXA_NAME%><br>
        考试日期：<%=exam_date%><br>
        考试地点：<%=T_ROOM%><br>
        考场号 ： <%=numberCover(T_ROOM_ORDER)%>  &nbsp;&nbsp;座位号： <%=numberCover(SEAT_ID)%><br>
        所在学校：<%=rs("SCHOOL")%>
   	</div>
    <h3>辽宁省高等学校师资培训中心监制</h3>
</div>
<div class="rightDiv">
<h4>考试注意事项</h4><br>
<div><%=sThing%></div>
<h5>考试科目与时间</h5>
<div class="topicTime">
<%=topicStr%></div>
</div>
<%
i=i+1
If i Mod 2=0 Then
	Response.write "<div class=""clearPage""></div>"
Else
	Response.write "<div class=""clearLine""><br><hr /><br></div>"
End If%>
<%

rs.movenext
loop
rs.close
set rs=nothing%>

</div>
<!--endprint-->
</body>
</html>