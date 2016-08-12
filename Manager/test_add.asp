<!--#include file="../inc/conn.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()
id=request("id")
operateType=0
CHECKED=1
IS_TOP=0
If  is_Id(id) Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT ID,TITLE,START_TIME,END_TIME,NOTICE_ID,CHECKED,TEST_CATE,CERT_TIME,CHECK_TIME,CERT_NO from T_TEST where ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		operateType=1
		NOTICE_ID=rs("NOTICE_ID")
		TITLE=rs("TITLE")
		START_TIME=format_time(FromUnixTime(rs("START_TIME"),+8),2)
		END_TIME=format_time(FromUnixTime(rs("END_TIME"),+8),2)
		CHECKED=rs("CHECKED")
	    TEST_CATE=rs("TEST_CATE")
		CERT_TIME = rs("CERT_TIME")
		CHECK_TIME = format_time(FromUnixTime(rs("CHECK_TIME"),+8),2)
		CERT_NO	 = rs("CERT_NO")
	End If
	rs.close
	set rs=nothing
End If

'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>考试信息管理</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<link href="../Css/news.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
    <link rel="stylesheet" href="../jquery/base/jquery.ui.all.css">
	<script src="../jquery/ui/jquery.timePicker.js"></script>
	<script src="../jquery/ui/jquery.ui.core.min.js"></script>
	<script src="../jquery/ui/jquery.ui.widget.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker-zh-CN.min.js"></script>
 <script type="text/javascript">

 $(document).ready(function(e) {
	$(function() {
		$( "#START_TIME" ).datepicker();
		$( "#END_TIME" ).datepicker();
		$( "#CERT_TIME" ).datepicker({dateFormat: "yy"});
		$( "#CHECK_TIME" ).datepicker({changeYear: true});
		
	});
		//table 颜色   
		$(".list_table tr").mouseover(function(){   
	   //如果鼠标移到class为stripe的表格的tr上时，执行函数   
	  		$(this).addClass("over");}).mouseout(function(){   
			//给这行添加class值为over，并且当鼠标一出该行时执行函数   
			$(this).removeClass("over");}) //移除该行的class  

  			$(".list_table tr:even").addClass("alt");
			
			
			
		 $(".list_table tr").click(function(){
			  		
				  if(this.className.indexOf('checkBgColor')<0){
					 	$(this).addClass('checkBgColor');
						$(this).find(":checkbox").attr("checked",true);
				  }
				  else{
					 $(this).removeClass('checkBgColor');
					 $(this).find(":checkbox").attr("checked",false);
				  }
			  });			
				
});



	function resetForm(){
		$('#editForm')[0].reset();
	}
	function submitForm(){
		
		
		var rMark =$("#rMark").val()
		var operateType =$("#operateType").val()

			
			if($("#TITLE").val()==""){
				alert("考试标题不能为空！");
				return false;
			}
			if($("#START_TIME").val()==""){
				alert("报考开始日期不能为空！");
				return false;
			}
			if($("#END_TIME").val()==""){
				alert("报考结束日期不能为空！");
				return false;
			}
			if($("#NOTICE_ID").val()=="0"){
				alert("请选择报考注意事项！");
				return false;
			}
					
	
			$.post(	"../Public/testEdit.asp", 
					$("#editForm").serialize(), 
					function(data,st){
							
							var resultArr=data.split("|");
							if(resultArr[0]=="1"){
								alert(resultArr[1]);
								
								window.location.reload();
								resetForm();
							}
							else{
								alert(resultArr[1]);
								
							}	
					});
		
	
	
	
	
	}

</script>
</head>

<body  >
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>考试添加</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)" onClick="window.location='manager.asp'">查询</a> <img src="../Images/001.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="resetForm();">添加记录</a></td>
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
      	<input id="ID" name="ID" value="<%=id%>" type="hidden">
<label for="TITLE">考试标题:</label>
<input type="text" id="TITLE" name="TITLE" value="<%=TITLE%>" />
<br>
<label for="START_TIME">报名开始时间:</label>
<input id="START_TIME" name="START_TIME" type="text" class="input" value="<%=START_TIME%>" readonly><br />
<label for="USER_NAME">报名结束时间:</label>
<input id="END_TIME" name="END_TIME" type="text" class="input" value="<%=END_TIME%>" readonly><br>
		<label for="NOTICE_ID">考试事项:</label> 		
        		<select id="NOTICE_ID" name="NOTICE_ID">
                <option value="0">选择考试注意事项<%=id%></option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,TITLE from T_NOTICE where CATEGORY=1"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=NOTICE_ID Then Response.write "selected=""selected"""  %>><%=rs("TITLE")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select>     <br/>
                
                <label for="CERT_TIME">合格证年份:</label>
<input id="CERT_TIME" name="CERT_TIME" type="text" class="input" value="<%=CERT_TIME%>" readonly><br>
				<label for="CHECK_TIME">签发日期:</label>
<input id="CHECK_TIME" name="CHECK_TIME" type="text" class="input" value="<%=CHECK_TIME%>" readonly><br>
				<label for="CERT_NO">起始编号:</label>
<input id="CERT_NO" name="CERT_NO" type="text" class="input" value="<%=CERT_NO%>" ><br>
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
   
<label for="CHECKED">审核:</label>  <input type="checkbox" value="1" id="CHECKED" name="CHECKED" class="radioType" <%If CHECKED=1 Then Response.Write "checked" End If%>> 
<br/>
<br/>
<br/>
<br/><br/>

   
    
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