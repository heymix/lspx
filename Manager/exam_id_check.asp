<script>
function submitCExam(){
	if ($("#sys_exam_id").val()=="0"){
		alert("未选择数据,无法操作！");
		return false;
	}
		$.post(	"../Public/exam_id_check.asp", 
				{sys_exam_id:$("#sys_exam_id").val(),sys_exam_title:$("#sys_exam_id option:selected").text()},
				function(data,st){		
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							//alert(resultArr[1]);
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
         <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/product.gif" width="16" height="16" /> <span class="tbTitle">选择管理的考点>>当前考点：<%=sys_exam_title%></span></td>
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
           <form method="post" id="cExamForm" name="cExamForm"  title="学校管理" >
             <label for="sys_exam_id">管理的考点:</label>
             <select id="sys_exam_id" name="sys_exam_id"  style=" width:300px;">
               <option value="0">选择考点</option>
               <%
					set rs=server.createobject("adodb.recordset")
					sql="select E.ID,E.EXA_NAME from T_TEST_EXA_RELATION AS ER Left Join T_EXA AS E ON ER.EXA_ID=E.ID where  1=1 and ER.TEST_ID="&TEST_ID&" and E.CHECKED=1 order by E.EXA_NAME asc"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
               <option value="<%=rs("ID")%>" <%If Cstr(rs("ID"))=sys_exam_id Then Response.write "selected='selected'" %>><%=rs("EXA_NAME")%></option>
               <%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
             </select>
             <img src="../Images/005.gif" width="14" height="14" /><a href="javascript:void(0)" onclick="submitCExam();">[确定]</a>
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
if sys_exam_id="" then
	appEnd()
End If
%>
