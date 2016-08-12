<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
BrowserString = Request.ServerVariables("HTTP_USER_AGENT")
BrowserString = Lcase(BrowserString)

fileName=test_title&"考试对应单"
response.ContentType ="application/vnd.ms-excel"
If inStr(BrowserString,"firefox")<>0 Then
	Response.AddHeader "content-disposition","attachment;filename="&fileName&".xls" 
Else
	Response.AddHeader "content-disposition","attachment;filename="&Server.UrlEncode(fileName)&".xls" 
End If

id=Request("id")
If not is_Id(id) Then
	appEnd()
End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="pragma" content="no-cache"> 
<meta http-equiv="cache-control" content="no-cache"> 
<meta http-equiv="expires" content="0"> 
<title>导出样式</title>
<style>
body,form,div,img,table,a,font,radio,td,th{
	margin: 0px;
	font-size: 12px;
	color: #1F4A65;
	padding: 0px;
	border:0px;
}
table{
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #000;
	border-bottom-color: #000;
	}
td,th{
	text-height: 80px;
	text-align: center;
	empty-cells: show;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: none;
	border-bottom-style: none;
	border-left-style: solid;
	border-top-color: #000;
	border-left-color: #000;
	white-space: nowrap ;
}
th{ text-align:center;
height:80px;}
#title{
	size: 29px;
	font-size: 24px;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	color: #333;	
	}
</style>
</head>
<body>
<table width="1500px;" cellpadding="0" cellspacing="0" border="1">
 <tr>
    <td colspan="9" id="title" style="height:60px;"><%=test_title&"考试对应单"%></td>
  </tr>
  <tr>
            <th  width="57"   bgcolor="#006633" style="color:#CCC; height:40px;">座位号</th>
            <th  width="235"   bgcolor="#006633" style="color:#CCC; height:40px;">准考证号</th>
            <th  width="146"   bgcolor="#006633" style="color:#CCC; height:40px;">身份证号</th>
            <th  width="146"   bgcolor="#006633" style="color:#CCC; height:40px;">姓名</th>
            <th  width="146"   bgcolor="#006633" style="color:#CCC; height:40px;">性别</th>
            <th  width="152"   bgcolor="#006633" style="color:#CCC; height:40px;">学校</th>
            <th  width="141"   bgcolor="#006633" style="color:#CCC; height:40px;">教室号</th>
            <th  width="146"   bgcolor="#006633" style="color:#CCC; height:40px;">考场号</th>
            <th  width="146"   bgcolor="#006633" style="color:#CCC; height:40px;">是否补考</th>
            <th  width="163"   bgcolor="#006633" style="color:#CCC; height:40px;">补考科目</th>
   
    
  </tr>
  <%sql="SELECT [ROOM_ID],[USER_ID],[SEAT_ID],[REAL_NAME],[TEST_NO],[T_ROOM],[T_ROOM_ORDER],[GENDER],[AGE],[MOBILE],[MOBILE_BAK],[TEL],[PHOTO],[SCHOOL],[IS_MAKEUP],[TEST_YEAR]FROM V_EXA_ORDER where ROOM_ID="&ID&" and (TEST_ID="&test_id&" or TEST_ID is null)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	If(rs.eof or rs.bof) Then%>
  <tr>
    <td colspan="9">没有查到数据</td>
  </tr>

<%
Else
	do while not (rs.eof or err)
	If rs("USER_ID")="" or isNull(rs("USER_ID")) Then
		IS_MAKEUP=""
	Else
		IS_MAKEUP=getTrue(rs("IS_MAKEUP"))
	End If

	%>    
          <tr>
            <td height="18" >'<%=numberCover(rs("SEAT_ID"))%>'</td>
            <td height="18" ><%=rs("TEST_NO")%></td>
            <td height="18" >'<%=rs("USER_ID")%>'</td>
            <td height="18" ><%=rs("REAL_NAME")%></td>
            <td height="18" ><%=getGender(rs("GENDER"))%></td>
            <td height="18" ><%=rs("SCHOOL")%></td>
            <td height="18" ><%=rs("T_ROOM")%></td>
            <td height="18" ><%=rs("T_ROOM_ORDER")%></td>
            <td height="18" ><%=IS_MAKEUP%></td>
            <td height="18" ><%If not (rs("USER_ID")="" or isNull(rs("USER_ID"))) and rs("IS_MAKEUP")=1 Then%><ul class="sList">
            <%set rs_topic=server.createobject("ADODB.Recordset")
			 sql="SELECT TT.ID,TT.TNAME,TT.IS_PRATICE FROM [T_TOPIC_RESULT] AS TR Left join T_TEST_TOPIC AS TT ON TR.TOPIC_ID=TT.ID where TR.TEST_ID="&test_id&" and TR.USER_ID='"&rs("USER_ID")&"'"
			 rs_topic.open sql,conn,1,3
			 	delStr="&nbsp;"
			  	If rs_topic.eof or rs_topic.bof Then
			 		Response.write "<li>没有记录</li>"
				Else
					do while not (rs_topic.eof or err)
						If rs_topic("IS_PRATICE")=0 Then
							Response.write "<li><span  class=""slMiddle"">"&rs_topic("TNAME")&"</span></li>"
						End If
				   	rs_topic.movenext
				   	loop
			  	End If
			   rs_topic.close
			   set rs_topic=nothing
			   %>
            </ul>
			<%End If%></td>
            </tr>
  <%
  rs.movenext
  loop
End If
  %>  
</table>

</body>
</html>
