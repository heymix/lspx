<!--#include file="../inc/conn.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()
TOPIC_ID=request("id")
If TEST_ID<>"" and isNumeric(TEST_ID) and TOPIC_ID<>"" and isNumeric(TOPIC_ID) Then
	set rs=server.createobject("adodb.recordset")
	sql="select R.TEST_ID,R.TOPIC_ID,R.EXA_DATE,R.EXA_TIME,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID where R.TEST_ID="&TEST_ID&" And R.TOPIC_ID="&TOPIC_ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		operateType=1
		TNAME=rs("TNAME")
		EXA_DATE=format_time(FromUnixTime(rs("EXA_DATE"),+8),2)
		EXA_TIME=rs("EXA_TIME")	
	End If
	rs.close
	set rs=nothing
Else
	Response.write errStr ("错误：没有查到需要修改的记录","-1")
	appEnd()
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
		$( "#EXA_DATE" ).datepicker();
		//$("#startTime").timePicker();
	});

});



	function resetForm(){
		$('#editForm')[0].reset();
	}
	function submitForm(){
		
			if($("#EXA_DATE").val()==""){
				alert("考试开始日期不能为空！");
				return false;
			}
			if($("#EXA_TIME").val()==""){
				alert("考试时间不能为空！");
				return false;
			}
					
	
			$.post(	"../Public/current_topic_reEdit.asp", 
					$("#editForm").serialize(), 
					function(data,st){
							
							var resultArr=data.split("|");
							if(resultArr[0]=="1"){
								alert(resultArr[1]);
								
								window.location.href="current_topic.asp";
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
      	<input id="ID" name="ID" value="<%=TOPIC_ID%>" type="hidden">
<label for="TNAME">考试标题:</label>
<input type="text" id="TNAME" name="TNAME" value="<%=TNAME%>"  disabled="disabled" />
<br>
<label for="EXA_DATE">科目考试日期:</label>
<input id="EXA_DATE" name="EXA_DATE" type="text" class="input" value="<%=EXA_DATE%>" readonly><br />
<label for="EXA_TIME">科目考试时间:</label>
<input id="EXA_TIME" name="EXA_TIME" type="text" class="input" value="<%=EXA_TIME%>" ><br>

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