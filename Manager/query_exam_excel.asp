<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
BrowserString = Request.ServerVariables("HTTP_USER_AGENT")
BrowserString = Lcase(BrowserString)

fileName=test_title&"统计报表"
response.ContentType ="application/vnd.ms-excel"
If inStr(BrowserString,"firefox")<>0 Then
	Response.AddHeader "content-disposition","attachment;filename="&fileName&".xls" 
Else
	Response.AddHeader "content-disposition","attachment;filename="&Server.UrlEncode(fileName)&".xls" 
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
<table  width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
          	<th width="2%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">序号</th>
            <th width="7%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">学校名称</th>
            <th width="6%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">培训人数</th>
            <th width="4%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">补考人数</th>
            <th width="6%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">缺考人数</th>
            <th width="9%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">考试人数</th>
            <th width="7%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">不及格人数</th>
            <th width="6%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">及格人数</th>
            <th width="37%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
           	  <th colspan="2" bgcolor="#006633" style="color:#CCC; height:40px;"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=1")(0)%></th>
              <th colspan="2" bgcolor="#006633" style="color:#CCC; height:40px;"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=2")(0)%></th>
              <th colspan="2" bgcolor="#006633" style="color:#CCC; height:40px;"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=3")(0)%></th>
              <th colspan="2" bgcolor="#006633" style="color:#CCC; height:40px;"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=4")(0)%></th>
            </tr>
           <tr>
           	<th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">人数</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">及格率%</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">人数</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">及格率%</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">人数</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">及格率%</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">人数</th>
            <th width="12%" bgcolor="#006633" style="color:#CCC; height:40px;">及格率%</th>
            
           </tr>
            
            </table></th>
            <th width="7%" height="26" bgcolor="#006633" style="color:#CCC; height:40px;">及格率</th>
  </tr>
         

  <%
If sys_exam_id<>"0" and is_Id(sys_exam_id) and Request("lSort")<>"1"  Then
	findKey=findKey&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&")"
End If
sql="select [SCHOOL],[SCHOOL_ID],[TEST_CATE] from [V_TEST_SCHOOL] where 1=1 and TEST_ID="&test_id&" "&findKey&" order by SCHOOL desc" 

	set rs=conn.execute(sql)     

If rs.eof or err Then%>  
          <tr>
            <td height="18"  colspan="11">[没有查到记录！]</div></td>
          </tr>

<%
Else
col1=0
col2=0
col3=0
col4=0
col5=0
col6=0
col7_1=0
col7_2=0
col8_1=0
col8_2=0 
col9_1=0
col9_2=0 
col10_1=0
col10_2=0 
i=0     
   do while not(rs.eof or err)
   i=i+1
	%>    
          <tr>
          <%
		  totalNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
		  ismakupNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where IS_MAKEUP=1 and SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
		  qkNum=conn.execute("select count(distinct(USER_ID)) from V_SCORE where IS_OLD_SCORE=0 and RESULT=0 and IS_PRATICE=0 and SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
		  
		  totalIsPassNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
		
		col1=col1+(totalNum-ismakupNum)
		col2=col2+(ismakupNum)
		col3=col3+(qkNum)
		col4=col4+(totalNum-qkNum)
		col5=col5+(totalNum-totalIsPassNum)
		col6=col6+(totalIsPassNum)
		  %>
          <td height="18" ><%=i+1%></td>
            <td height="18" ><%=rs("SCHOOL")%></td>
           <td height="18" ><%=totalNum-ismakupNum%></td>
            <td height="18" ><%=ismakupNum%></td>
            <td height="18" ><%=qkNum%></td>
            <td height="18" ><%=totalNum-qkNum%></td>
            <td height="18" ><%=totalNum-totalIsPassNum%></td>
            <td height="18" ><%=totalIsPassNum%></td>
            
            <td height="18" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
           	<td width="12%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)%></td>
            <td width="12%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								col7_1=col7_1+isPassNum
								col7_2=col7_2+isAllNum
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and IS_PASS=1  and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)%></td>
            <td width="12%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								col8_1=col8_1+isPassNum
								col8_2=col8_2+isAllNum
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and IS_PASS=1  and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)%></td>
            <td width="12%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								col9_1=col9_1+isPassNum
								col9_2=col9_2+isAllNum
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)%></td>
            <td width="12%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and IS_PASS=1 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and  SCHOOL_ID="&rs("SCHOOL_ID")&" and TEST_ID="&test_id)(0)
								col10_1=col10_1+isPassNum
								col10_2=col10_2+isAllNum
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>

           </tr>
            
            </table></td>
            <td height="18" ><%	
								If totalNum>0 Then
									perNum=ROUND(totalIsPassNum/totalNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
  </tr>
       
    <%
   rs.moveNext	
loop

%>

<td height="18" colspan="2">合计</td>
            
            <td height="18" ><%=col1%></td>
            <td height="18" ><%=col2%></td>
            <td height="18" ><%=col3%></td>
            <td height="18" ><%=col4%></td>
            <td height="18" ><%=col5%></td>
            <td height="18" ><%=col6%></td>
            <td height="18" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
           	<td width="12.5%"><%=col7_1%></thd>
            <td width="12.5%"><%If col7_2>0 Then
									perNum=ROUND(col7_1/col7_2,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=col8_1%></td>
            <td width="12.5%"><%If col8_2>0 Then
									perNum=ROUND(col8_1/col8_2,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=col9_1%></td>
            <td width="12.5%"><%If col9_2>0 Then
									perNum=ROUND(col9_1/col9_2,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=col10_1%></td>
            <td width="12.5%"><%If col10_2>0 Then
									perNum=ROUND(col10_1/col10_2,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            
           </tr></table>
           </td>
           <td><%	
								If (col1+col2)>0 Then
									perNum=ROUND(col6/(col1+col2),2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
</tr></table>
<%End If%>
</body>
</html>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>