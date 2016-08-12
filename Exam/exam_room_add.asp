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


    function lookup(inputString) {
        if(inputString.length == 0) {
            // Hid''e the suggestion box.
            $('#suggestions').hide();
        } else {
            $.post("../Public/examListKey.asp", {queryString: inputString}, function(data){
                if(data.length >0) {
                    $('#suggestions').show();
                    $('#autoSuggestionsList').html(data);
                }
            });
        }
    } // lookup
    
    function fill(thisValue,id) {
        $('#inputString').val(thisValue);
		$('#SCHOOL_ID').val(id);
        setTimeout("$('#suggestions').hide();", 200);
    }
	function resetForm(){
		$('#editForm')[0].reset();
	}
	function submitForm(){
		var operateType =$("#operateType").val()
			if($("#EXA_ID").val()=="0"){
				alert("考点不能为空！");
				return false;
			}
			if($("#T_ROOM").val()==""){
					alert("教室称不能为空！");
					return false;
			}
			if($("#T_ROOM_ORDER").val()==""){
					alert("考场号不能为空！");
					return false;
			}
			if($("#T_MAXNUM").val()==""){
					alert("考场最大座位数！");
					return false;
			}
			$.post(	"../Public/exam_roomEdit.asp", 
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

<body><%

id=request("id")
operateType=0
CHECKED=1
IS_TOP=0
SCHOOL_ID=0
T_MAX_NUM=30
If is_Id(id) Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT *  FROM [T_EXA_ROOM] where ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		operateType=1
		EXA_ID=rs("EXA_ID")
		T_ROOM=rs("T_ROOM")
		T_ROOM_ORDER=rs("T_ROOM_ORDER")
		T_MAX_NUM=rs("T_MAX_NUM")
		CHECKED=rs("CHECKED")	
	End If
	rs.close
	set rs=nothing
End If

%>
<!--#include file="test_id_check.asp"-->
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>编辑考场号</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)" onClick="window.location.href='exam_room.asp'">查询</a> </td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm"  action="../Public/exam_roomEdit.asp" >
        <input id="operateType" name="operateType" value="<%=operateType%>" type="hidden">
       <input id="ID" name="ID" value="<%=ID%>" type="hidden">
       <input id="pre_test_id" name="pre_test_id" value="<%=test_id%>" type="hidden">
        <input id="pre_sys_exam_id" name="pre_sys_exam_id" value="<%=sys_exam_id%>" type="hidden">
       

<br/> 
    <label for="T_ROOM">教室:</label>
<input type="text" id="T_ROOM" name="T_ROOM" value="<%=T_ROOM%>" /> <br />
<label for="T_ROOM_ORDER">考场号:</label>
<input type="text" id="T_ROOM_ORDER" name="T_ROOM_ORDER" value="<%=T_ROOM_ORDER%>" /> <br />
<label for="T_MAX_NUM">最大座位数:</label>
<input type="text" id="T_MAX_NUM" name="T_MAX_NUM" value="<%=T_MAX_NUM%>" /> <br><br>
<% If operateType=1 Then%>
<p>如果座位数已经分配完毕 请不要在此处修改座位最大值 如果修改将按最大的座位号码，进行增、减。</p><br />
<%End If%>
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