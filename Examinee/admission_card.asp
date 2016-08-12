<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Examinee/CheckExamineeOL.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()

TITLE	 	= Request("TITLE")

If TITLE<>"" and not sqlInjection(TITLE) Then
	findKey=findKey&" and N.TITLE like'%"&TITLE&"%'"
End If
'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理员基础信息</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
    <link rel="stylesheet" href="../jquery/base/jquery.ui.all.css">
	<script src="../jquery/ui/jquery.timePicker.js"></script>
	<script src="../jquery/ui/jquery.ui.core.min.js"></script>
	<script src="../jquery/ui/jquery.ui.widget.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker-zh-CN.min.js"></script>
 <script type="text/javascript">

 $(document).ready(function(e) {

	//多选框操作
	 $("#checkAll").click(function() {
	  $('input[name="subBox"]').attr("checked",this.checked); 
  });
  var $subBox = $("input[name='subBox']");
  $subBox.click(function(){
	  $("#checkAll").attr("checked",$subBox.length == $("input[name='subBox']:checked").length ? true : false);
  });
  
  		 $("#inverse").click(function () {//反选
                $("#listForm :checkbox").each(function () {
                    $(this).attr("checked", !$(this).attr("checked"));
                });
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

		$("#editForm").submit();
	
	
	}


//锁定  status 0 审校 1完结状态 2刷新 4删除
function statusPost(status,type,checkType){
	
	if (status=="4"){
		if(!confirm("确认删除数据吗? 不可恢复...")) return false;
	}
		$.post(	"../Public/myexamStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
				$("#listForm").serialize(), 
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
</head>

<body  >
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">考试列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
            <th width="15%" height="26">身份证号</th>
            <th width="8%" height="26">姓名</th>
            <th width="10%" height="26">考试</th>
            <th width="10%" height="26">科目</th>
            <th width="7%" height="26">分数</th>
            <th width="7%" height="26">及格</th>
            <th width="14%" height="26">是否补考</th>
            <th width="13%" height="26">是否出分</th>
            </tr>
          

<%
sql="SELECT [USER_ID],[TEST_ID],[TOPIC_ID],[RESULT],[IS_PASS],[IS_MAKEUP],[IS_SCORE],[REAL_NAME],[TNAME],[TITLE] FROM [V_SCORE] where USER_ID='"&EXAMINEE_USER_ID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	If(rs.eof or rs.bof) Then
%>   
          <tr>
            <td height="18"  colspan="10">[没有查到记录！]</div></td>
          </tr>

<%
Else     
   
	do while not (rs.eof or err)
	%>    
          <tr>
            <td height="18" ><%=rs("USER_ID")%></td>
            <td height="18" ><%=rs("REAL_NAME")%></td>
            <td height="18" ><%=rs("TITLE")%></td>
            <td height="18" ><%=rs("TNAME")%></td>
            <td height="18" ><%=rs("RESULT")%></td>
            <td height="18" ><%=getTrue(rs("IS_PASS"))%></td>
            <td height="18" ><%=getTrue(rs("IS_MAKEUP"))%></td>
            <td height="18" ><%=getTrue(rs("IS_SCORE"))%></tr>
       
    <%
    rs.movenext
  loop	
End If
%>
        </table></form></td>
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
            <td  height="29" id="bNavStr"></td>
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