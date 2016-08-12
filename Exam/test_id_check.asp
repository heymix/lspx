<script>
function submitChoiceTest(){
	if ($("#TEST_ID").val()=="0"){
		alert("未选择数据,无法操作！");
		return false;
	}
		$.post(	"../Public/test_id_check_exam.asp", 
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
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr>
         <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
         <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/nav_01.gif" width="16" height="16" /> <span class="tbTitle">选择管理的考试>>当前考试：<%=test_title%></span></td>
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
           <form method="post" id="choiceTestForm" name="choiceTestForm"  title="学校管理" style=" width:500px;" >
             <label for="TEST_ID">管理的考试:</label>
             <select id="TEST_ID" name="TEST_ID" style="width:300px;">
               <option value="0">选择考试</option>
               <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,TITLE,TEST_CATE from T_TEST where CHECKED=1  and id in(select TEST_ID from T_TEST_EXA_RELATION where  EXA_ID="&sys_exam_id&")"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
               <option value="<%=rs("ID")%>|<%=rs("TEST_CATE")%>" <%If Cstr(rs("ID"))=test_id Then Response.write "selected='selected'" %>><%=rs("TITLE")%></option>
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
if test_id="" then
	appEnd()
End If
%>
