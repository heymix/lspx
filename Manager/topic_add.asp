<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
			if($("#TNAME").val()==""){
				alert("科目名称不能为空！");
				return false;
			}
			if($("#TEST_CATE").val()=="0"|| $("#TEST_CATE").val()==""){
				alert("考试类别不能为空！");
				return false;
			}
			$.post(	"../Public/topicEdit.asp", 
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
	sql="SELECT T.ID,T.TNAME,T.TEST_CATE,T.IS_PRATICE,D.KEY_VALUE AS TEST_CATE_NAME FROM T_TEST_TOPIC AS T Left Join T_DATA_KEY AS D ON T.TEST_CATE=D.ID where T.ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		operateType=1
		TNAME=rs("TNAME")
		TEST_CATE_NAME=rs("TEST_CATE_NAME")
		TEST_CATE=rs("TEST_CATE")
		IS_PRATICE=rs("IS_PRATICE")
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
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>考试科目</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)" onClick="window.location.href='topic.asp'">查询</a> </td>
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
       

    <label for="TNAME">科目名称:</label> <input type="text" size="30"id="TNAME" name="TNAME"  value="<%=TNAME%>" />
                    <input  type="hidden" name="ID" id="ID" value="<%=ID%>"><br/>
 
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

<label for="IS_PRATICE"  >实践科目</label><input type="checkbox" name="IS_PRATICE" value="1"  class="radioType"<%If IS_PRATICE=1 Then response.write "checked"%>/><br><br>
    
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