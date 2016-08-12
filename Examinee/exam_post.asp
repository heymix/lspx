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

 $("#IS_MAKEUP_bm").click(function() {
		
		if ($(":checkbox#IS_MAKEUP_bm").attr("checked") != "checked") {
			$("#isTopic").hide()
		}else{
			$("#isTopic").show()
		}
  });
		
				
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

<style>
#editForm label{ width:100px;	
}

</style>
</head>

<body  >
<script>
function submitChoiceTest(){
	if ($("#TEST_ID").val()=="0"){
		alert("未选择数据,无法操作！");
		return false;
	}
		$.post(	"../Public/test_id_check_post.asp", 
				{test_id:$("#TEST_ID").val(),test_title:$("#choiceTestForm option:selected").text()},
				function(data,st){		
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
						}	
				});
}
</script>
<% test_id_post= request.Cookies("test")("test_id_post")
   test_title_post= request.Cookies("test")("test_title_post")
%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
         <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
         <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/nav_01.gif" width="16" height="16" /> <span class="tbTitle">选择报考的考试>>当前考试：<%=test_title%></span></td>
         <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
         <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
       </tr>
     </table></td>
   </tr>
   <tr>
     <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
         <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
         <td bgcolor="#f3ffe3"  class="tdStyle"><div  class="choiceTest">
           <form method="post" id="choiceTestForm" name="choiceTestForm">
             <label for="TEST_ID">报考的考试:</label>
             <select id="TEST_ID" name="TEST_ID">
               <option value="0">选择考试</option>
               <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,TITLE,TEST_CATE from T_TEST as main where CHECKED=1 and IS_OVER=0 and (select SCHOOL_ID from T_EXAMINEE where USER_ID='"&SYS_USER_ID&"') in (select SCHOOL_ID  from T_EXA_SCHOOL_RELATION where EXA_ID in (select EXA_ID from T_TEST_EXA_RELATION as TE where main.ID=TE.TEST_ID)) and END_TIME>"&ToUnixTime(now(),+8)&" and START_TIME<"&ToUnixTime(now(),+8)
					response.write sql
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
               <option value="<%=rs("ID")%>|<%=rs("TEST_CATE")%>" <%If Cstr(rs("ID"))=test_id_post Then Response.write "selected='selected'" %>><%=rs("TITLE")%></option>
               <%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
             </select>
             <img src="../Images/005.gif" width="14" height="14" /><a href="javascript:void(0)" onclick="submitChoiceTest();">[确定]</a>
           </form>
         </div></td>
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
 <%
if test_id_post="" then
	appEnd()
End If
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
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm"  action="../Public/exam_postEdit.asp" >
        <input id="operateType" name="operateType" value="0" type="hidden">
       <label for="myTest_title">报考的考试</label><input type="text" id="myTest" name="myTest" value="<%=test_title_post%>" disabled="disabled" style="width:300px" ><input type="hidden" id="myTest_id" name="myTest_id" value="<%=test_id_post%>" ><br>
		<input type="hidden" id="PUR_SYS_ID" name="PUR_SYS_ID" value="<%=EXAMINEE_USER_ID%>" /><br>
            <input name="IS_MAKEUP_bm" id="IS_MAKEUP_bm" type="checkbox" value="1" class="radioType" /><label for="IS_MAKEUP" style="width:500px;">是否补考&nbsp;&nbsp;&nbsp;&nbsp;<font style="color:#F00;">如果补考请勾上选择要补考的科目。如果第一次报考请直接点《确定报考》按钮</font></label> <br>
              <ul class="sList" id="isTopic" style="display:none" >
               <label for="TOPIC_ID">补考的科目:</label>
               <%
					set rs=server.createobject("adodb.recordset")
					sql="select T.ID,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID WHERE R.TEST_ID="&test_id_post
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
 					<img src="../Images/005.gif" width="14" height="14" />[<a id="submitLink" href="javascript:void(0)" onClick="submitForm();">确认报考</a>] <img src="../Images/002.gif" width="14" height="14"/>[<a href="javascript:void(0)" onClick="resetForm();">重置</a>]
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