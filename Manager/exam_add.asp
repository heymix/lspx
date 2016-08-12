<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
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
<link href="../Css/news.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
 <script type="text/javascript">

 $(document).ready(function(e) {

	
  
  //单选框操作
$("#isSame").click(function() {
	 // $('input[name="subBox"]').attr("checked",this.checked); 
	  if ($("#isSame").attr('checked')=="checked") {
		$("#EXA_NAME").attr("disabled","disabled");
		  	$("#EXA_NAME").css("disabled","true")
		 }else{
			$("#EXA_NAME").removeAttr("disabled");
		}
  });
		
		
				
});

	function resetForm(){
		$('#editForm')[0].reset();
	}
	function submitForm(){
		var operateType =$("#operateType").val()
			if($("#SCHOOL_ID").val()=="0"|| $("#SCHOOL_ID").val()==""){
				alert("学校不能为空！");
				return false;
			}
			if(!($("#isSame").attr('checked')=="checked")){
				if($("#EXA_NAME").val()==""){
					alert("考点名称不能为空！");
					return false;
				}
			}
			if(!($("#isSame").attr('checked')=="checked")){
				if($("#EXA_NAME").val()==""){
					alert("考点名称不能为空！");
					return false;
				}
			}
			$.post(	"../Public/examEdit.asp", 
					$("#editForm").serialize(), 
					function(data,st){
							var resultArr=data.split("|");
							if(resultArr[0]=="1"){
								alert(resultArr[1]);
								window.location.reload();
							}
							else{
								alert(resultArr[1]);
								
							}	
					});
		
	}
</script>
</head>

<body><%

id=request("id")
operateType=0
CHECKED=1
IS_TOP=0
SCHOOL_ID=0
If is_Id(id) Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT E.ID,S.SCHOOL,E.SCHOOL_ID,E.EXA_NAME,E.IS_RUN,E.CHECKED,E.TEST_CATE,E.CITY_ID FROM T_EXA AS E left join T_SCHOOL AS S on E.SCHOOL_ID=S.ID where E.ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		operateType=1
		SCHOOL=rs("SCHOOL")
		SCHOOL_ID=rs("SCHOOL_ID")
		TEST_CATE=rs("TEST_CATE")
		CITY_ID=rs("CITY_ID")
		EXA_NAME=rs("EXA_NAME")
		CHECKED=rs("CHECKED")	
	End If
	rs.close
	set rs=nothing
End If

%>

<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">操作面板</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)" onClick="window.location='exam.asp'">查询</a> </td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm"  title="学校管理" >
        <input id="operateType" name="operateType" value="<%=operateType%>" type="hidden">
       

<%If operateType=1 Then%>
<input id="ID" name="ID" value="<%=id%>" type="hidden">
<label for="inputString">学校:</label> <input disabled="disabled" value="<%=SCHOOL%>"> 
<input  type="hidden" name="SCHOOL_ID" id="SCHOOL_ID" value="<%=SCHOOL_ID%>">
 <input type="hidden" id="inputString" name="inputString" value="<%=SCHOOL%>" />
<%Else%>
	<label for="SCHOOL_ID">学校:</label>
            <select id="SCHOOL_ID" name="SCHOOL_ID" onchange="$('#inputString').val($(this).find('option:selected').text())">
                <option value="0">*选择学校</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,SCHOOL from T_SCHOOL where CHECKED=1  order by SCHOOL asc"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=SCHOOL_ID Then Response.write "selected=""selected"""  %>><%=rs("SCHOOL")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select>
                 <input type="hidden" id="inputString" name="inputString" value="<%=SCHOOL%>" />
 <%End If%>  
 <br/> 
 <label for="TEST_CATE">类别:</label>
 <select id="TEST_CATE" name="TEST_CATE">
                <option value="0">选择类别</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,KEY_ID,KEY_VALUE from T_DATA_KEY where KEY_ID=4"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=TEST_CATE Then Response.write "selected=""selected"""  %>><%=rs("KEY_VALUE")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select><br>
  <label for="TEST_CATE">城市:</label>
 <select id="CITY_ID" name="CITY_ID">
                <option value="0">选择考点所在地</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,KEY_ID,KEY_VALUE from T_DATA_KEY where KEY_ID=5"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=CITY_ID Then Response.write "selected=""selected"""  %>><%=rs("KEY_VALUE")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select><br>
    <label for="EXA_NAME">考点名称:</label>
<input type="text" id="EXA_NAME" name="EXA_NAME" value="<%=EXA_NAME%>" /><label for="CHECKED">与学校同名:</label>  <input type="checkbox" value="1" id="isSame" name="isSame" class="radioType"> <br />
<label for="CHECKED">审核:</label>  <input type="checkbox" value="1" id="CHECKED" name="CHECKED" class="radioType" <%If CHECKED=1 Then Response.Write "checked" End If%>> 
<br/>
   
    
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