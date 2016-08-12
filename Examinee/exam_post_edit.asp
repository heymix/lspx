<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Examinee/CheckExamineeOL.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()

'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理员基础信息</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
 <script type="text/javascript">

 $(document).ready(function(e) {


		
				
});


	function resetForm(){
		$('#editForm')[0].reset();;
	}
	
	function submitForm(){
		if ($(":checkbox#IS_MAKEUP").attr("checked") == "checked") {
			var n = $("#editForm input:[name=subBox][checked]").length;
			if (n == 0) {
				alert("如果是补考 请选择科目！");
				return false;
			}
		} 
			$.post(	"../Public/exam_postEdit.asp", 
					$("#editForm").serialize(), 
					function(data,st){
							
							var resultArr=data.split("|");
							if(resultArr[0]=="1"){
								alert(resultArr[1]);
							}
							else{
								alert(resultArr[1]);
								
							}	
					});
		
	
	
	
	
	}





</script>
</head>

<body  >
<%id=Request("id")
If not is_Id(id) Then
	Response.write errStr("参数错误！","-1")
	appEnd()
	
End If

	set rs=server.createobject("adodb.recordset")
	sql="SELECT R.USER_ID,E.REAL_NAME,R.TEST_ID,T.TITLE,R.TEST_YEAR,R.IS_MAKEUP,E.CHECKED FROM T_TEST_RESULT AS R Left Join T_EXAMINEE AS E On R.USER_ID=E.USER_ID left join T_TEST AS T ON R.TEST_ID=T.ID where R.TEST_ID="&ID&" and E.USER_ID='"&EXAMINEE_USER_ID&"'"
	'Response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else

		TEST_ID	=rs("TEST_ID")
		TITLE	=rs("TITLE")
		IS_MAKEUP=rs("IS_MAKEUP")
			
	End If
	rs.close
	set rs=nothing

%>

<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm" >
        <input id="operateType" name="operateType" value="1" type="hidden">
       <label for="myTest_title">报考的考试</label><input type="text" id="myTest" name="myTest" value="<%=TITLE%>" disabled="disabled" ><input type="hidden" id="myTest_id" name="myTest_id" value="<%=TEST_ID%>" ><br>
		<input type="hidden" id="PUR_SYS_ID" name="PUR_SYS_ID" value="<%=EXAMINEE_USER_ID%>" /><br>
            <input name="IS_MAKEUP" id="IS_MAKEUP" type="checkbox" value="1" class="radioType" <%If IS_MAKEUP=1 then response.write "checked"%>/><label for="IS_MAKEUP">是否补考</label><br>
              <ul class="sList" id="isTopic" >
               <label for="TOPIC_ID">报考的科目:</label>
               <%
					set rs=server.createobject("adodb.recordset")
					sql="select T.ID,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID WHERE R.TOPIC_ID in(select TOPIC_ID from T_TOPIC_RESULT where USER_ID='"&EXAMINEE_USER_ID&"' and TEST_ID="&ID&") And R.TEST_ID="&ID
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
					
						%>
                      
            <li style="height:30px"><input name="subBox" type="checkbox" value="<%=rs("ID")%>" class="radioType" checked="checked" /><label for="subBox"><%=rs("TNAME")%></label> </li>

           
              
               <%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                 <%
					set rs=server.createobject("adodb.recordset")
					sql="select T.ID,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID WHERE R.TOPIC_ID not in(select TOPIC_ID from T_TOPIC_RESULT where USER_ID='"&EXAMINEE_USER_ID&"' and TEST_ID="&ID&") And R.TEST_ID="&ID
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
					
						%>
                      
            <li style="height:30px"><input name="subBox" type="checkbox" value="<%=rs("ID")%>" class="radioType" /><label for="subBox"><%=rs("TNAME")%></label> </li>

           
              
               <%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%> </ul>
      

   <br><br>
            <div  class="suggestionsMain" style="text-align:right">
 					<img src="../Images/005.gif" width="14" height="14" />[<a id="submitLink" href="javascript:void(0)" onClick="submitForm();">提交数据</a>] <img src="../Images/002.gif" width="14" height="14"/>[<a href="javascript:void(0)" onClick="resetForm();">重置</a>]
              </div>
   
        </form></td>
        <td width="9" background="../Images/tab_16.gif">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="29"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="29"><img src="../Images/tab_20.gif" width="15" height="29" /></td>
        <td background="../Images/tab_21.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td  height="29" >&nbsp;</td>
          </tr>
        </table></td>
        <td width="14"><img src="../Images/tab_22.gif" width="14" height="29" /></td>
      </tr>
    </table></td>
  </tr>
</table>
<!--本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒-->

</body>
</html>

<%
appEnd()

%>