<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
BrowserString = Request.ServerVariables("HTTP_USER_AGENT")
BrowserString = Lcase(BrowserString)

fileName=test_title&"成绩信息"
response.ContentType ="application/vnd.ms-excel"
If inStr(BrowserString,"firefox")<>0 Then
	Response.AddHeader "content-disposition","attachment;filename="&fileName&".xls" 
Else
	Response.AddHeader "content-disposition","attachment;filename="&Server.UrlEncode(fileName)&".xls" 
End If

REAL_NAME	 	= Request("REAL_NAME")
If REAL_NAME<>"" and not sqlInjection(REAL_NAME) Then
	findKey=findKey&" and REAL_NAME ='"&REAL_NAME&"'"
End If
USER_ID	 	= Request("USER_ID")
If USER_ID<>"" and not sqlInjection(USER_ID) Then
	findKey=findKey&" and USER_ID ='"&USER_ID&"'"
End If

TEL	 	= Request("TEL")
If TEL<>"" and not sqlInjection(TEL) Then
	findKey=findKey&" and TEL ='"&TEL&"'"
End If

MOBILE	 	= Request("MOBILE")
If MOBILE<>"" and not sqlInjection(MOBILE) Then
	findKey=findKey&" and MOBILE ='"&MOBILE&"'"
End If

EMAIL	 	= Request("EMAIL")
If EMAIL<>"" and not sqlInjection(EMAIL) Then
	findKey=findKey&" and EMAIL ='"&EMAIL&"'"
End If

CHECKED	 	= Request("CHECKED")
If CHECKED="1" or CHECKED="0"  Then
	findKey=findKey&" and CHECKED="&CHECKED
End If

LOCKED	 	= Request("LOCKED")
If LOCKED="1" or LOCKED="0"  Then
	findKey=findKey&" and LOCKED="&LOCKED
End If

IS_MAKEUP	 	= Request("IS_MAKEUP")
If IS_MAKEUP="1" or IS_MAKEUP="0"  Then
	findKey=findKey&" and USER_ID in (select USER_ID From T_TEST_RESULT where IS_MAKEUP="&IS_MAKEUP&")"
End If

SCHOOL_ID	 	= sys_school_id
If SCHOOL_ID<>"0" and is_Id(SCHOOL_ID)  Then
	findKey=findKey&" and SCHOOL_ID="&SCHOOL_ID
End If

If sys_exam_id<>"0" and is_Id(sys_exam_id)  Then
	findKey=findKey&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&")"
End If

TEST_NO	 	= trim(Request("TEST_NO"))
If TEST_NO<>"" and is_Id(TEST_NO)  Then
	findKey=findKey&" and TEST_NO="&TEST_NO&" "
End If

IS_PASS	 	= Request("IS_PASS")
If IS_PASS="1" or IS_PASS="0"  Then
	findKey=findKey&" and IS_PASS ="&IS_PASS
End If

TOPIC_ID	 	= Request("TOPIC_ID")
If TOPIC_ID<>"-1" and is_Id(TOPIC_ID)  Then
	findKey=findKey&" and USER_ID in(select USER_ID from T_TOPIC_RESULT where TOPIC_ID="&TOPIC_ID&" and test_id="&test_id&")"
	topicFind=topicFind&" and TOPIC_ID="&TOPIC_ID
End If

sql="select TEST_CATE from T_TEST where ID="&test_id&""
set rs=conn.execute(sql)
if  not(rs.eof or err) then
	TOPIC_CATE=rs("TEST_CATE")
end if

rs.close
set rs=nothing

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
    <td colspan="16" id="title" style="height:60px;"><%=test_title&"考生报考信息"%></td>
  </tr>
  <tr>
    <th width="63"   bgcolor="#006633" style="color:#CCC; height:40px; width:10px">序号</th>
     <th width="109"  bgcolor="#006633" style="color:#CCC; height:40px;">准考证号</th>
      <th width="109"  bgcolor="#006633" style="color:#CCC; height:40px;">证书编号</th>
      <th width="116"  bgcolor="#006633" style="color:#CCC; height:40px;">所在学校</th>
     <th width="79"  bgcolor="#006633" style="color:#CCC; height:40px;">姓名</th>
    <th width="36"  bgcolor="#006633" style="color:#CCC; height:40px;">性别</th>
    <th width="171"  bgcolor="#006633" style="color:#CCC; height:40px;">身份证号</th>
    <th width="69"  bgcolor="#006633" style="color:#CCC; height:40px;">最后学历</th>
    <th width="210"  bgcolor="#006633" style="color:#CCC; height:40px;">毕业学校</th>
    <th width="107"  bgcolor="#006633" style="color:#CCC; height:40px;">毕业时间</th>
    <th width="83"  bgcolor="#006633" style="color:#CCC; height:40px;">所学专业</th>
    <th width="115"  bgcolor="#006633" style="color:#CCC; height:40px;">
    <table>
        <tr>
        	<th colspan="5"  bgcolor="#006633" style="color:#CCC; height:40px;">各科考试成绩</th>
        </tr>
    <tr>
    		<%set rs=conn.execute("SELECT [TOPIC_ID],TNAME FROM T_TEST_TOPIC_RELATION as TR Left Join T_TEST_TOPIC AS TT on TR.TOPIC_ID=TT.ID where TT.TEST_CATE="&TOPIC_CATE&" and TR.TEST_ID="&test_id&"  order by [TOPIC_ID] asc")
			do while not(rs.eof or err)
				%>
				<th  bgcolor="#006633" style="color:#CCC; height:40px; "><%=rs("TNAME")%></th>
				<%
			rs.moveNext
			loop
			rs.close
			set rs=nothing%>              
            </tr></table></th>
    <th width="65"  bgcolor="#006633" style="color:#CCC; height:40px;">电话</th>
    <th width="68"  bgcolor="#006633" style="color:#CCC; height:40px;">手机</th>
    <th width="63"  bgcolor="#006633" style="color:#CCC; height:40px;">Email</th>
    <th width="61"  bgcolor="#006633" style="color:#CCC; height:40px;">补考</th>
   
    
  </tr>
  <%set rs=server.createobject("adodb.recordset")
	sql="select USER_ID,SCHOOL_ID,SCHOOL,GENDER,REAL_NAME,AGE,EDUCATION,TEL,MOBILE,EMAIL,LAST_LOGIN_TIME,CHECKED,LOCKED,IS_MAKEUP,TRCHECKED,EDU_NAME, GRADUATE_SCHOOL, GRADUATE_DATE,PRO,IS_MAKEUP,CRS_NO from V_EXAMINEE_TEST where 1=1 and TEST_ID="&test_id&"  "&findKey&" order by TEST_NO asc"
	rs.open sql,conn,1,1
	If(rs.eof or rs.bof) Then%>
  <tr>
    <td colspan="16">没有查到数据</td>
  </tr>
  <%
  ELse
  	i=0
  	do while not (rs.eof or err)
	i=i+1
	%>
	
	
	  <tr>
    <td width="63"><%=i%></td>
    <td width="109"><%=getTestNo(rs("USER_ID"))%></td>
     <td width="116"><%=rs("CRS_NO")%></td>
     <td width="116"><%=rs("SCHOOL")%></td>
    <td width="79"><%=rs("REAL_NAME")%></td>
    <td width="36"><%=getGender(rs("GENDER"))%></td>
    <td width="171">'<%=rs("USER_ID")%></td>
    <td width="69"><%=rs("EDU_NAME")%></td>
    <td width="210"><%=rs("GRADUATE_SCHOOL")%></td>
    <td width="107"><%=Format_Time(FromUnixTime(rs("GRADUATE_DATE"),+8),7)%></td>
    <td width="83"><%=rs("PRO")%></td>
    <td width="115">
    <table width="100%">
    <tr>
           
                
               <%rs_tt_sql="SELECT [TOPIC_ID],TNAME,TT.IS_PRATICE,(select RESULT from V_SCORE where  V_SCORE.TOPIC_ID=TR.TOPIC_ID and IS_SCORE=1 and TEST_CATE="&TOPIC_CATE&"  and TEST_ID="&test_id&" and USER_ID='"&rs("USER_ID")&"') as RESULT  FROM T_TEST_TOPIC_RELATION as TR Left Join T_TEST_TOPIC AS TT on TR.TOPIC_ID=TT.ID where TT.TEST_CATE="&TOPIC_CATE&" and TR.TEST_ID="&test_id&" order by [TOPIC_ID] asc"

			set rs_tt=conn.execute(rs_tt_sql)
			do while not(rs_tt.eof or err)
				exam_score="&nbsp;"


					If rs_tt("IS_PRATICE")=1 Then
						If not isNull(rs_tt("RESULT")) then
							If Cstr(rs_tt("RESULT"))="1" Then
								exam_score="合格"
							Else
								exam_score="不合格"
							End If
						Else
							exam_score="0"
						End If
					Else 
						If isNull(rs_tt("RESULT")) then
							exam_score=0
						Else
							exam_score=rs_tt("RESULT")
						End if
					End If

				%>
				<td ><%=exam_score%></td>
				<%
			rs_tt.moveNext
			loop
			rs_tt.close
			set rs_tt=nothing%>
               </tr>
          </table>  
    </td>
    <td width="65"><%=rs("TEL")%></td>
    <td width="68"><%=rs("MOBILE")%></td>
    <td width="63"><%=rs("EMAIL")%></td>
    <td width="61"><%=getTrue(rs("IS_MAKEUP"))%></td>
    
  </tr>
	
	<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
End If
  %>
</table>

</body>
</html>
