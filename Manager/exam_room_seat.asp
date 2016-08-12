<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()
TITLE	 	= Request("TITLE")
If TITLE<>"" and not sqlInjection(TITLE) Then
	findKey=findKey&" and E.EXA_NAME like'%"&TITLE&"%'"
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

	var add_id_arr="";
	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		if (status=="4"){
			if(!confirm("确认删除数据吗? 不可恢复...")) return false;
		}
		
		if(status=="8")	{
			 $("#listForm input:checked").each(function(index, element) {
                add_id_arr+=$(this).val()+",";
            });
			add_id_arr=add_id_arr.substring(0,add_id_arr.length-1);
			window.location.href='exam_room_seat_add.asp?id='+add_id_arr+'&ROOM_ID=<%=request("ID")%>';
		return false	
		}
		$.post(	"../Public/exam_room_seat_addStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
				$("#listForm").serialize(), 
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
		
}





</script>
</head>

<body>
<%
id=Request("id")
If not is_Id(id) Then
	Response.write errStr ("错误：没有查到需要调整的记录","-1")
	appEnd()
End If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">考试列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/run.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="statusPost('4','multi','0');">移除考生</a> <img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('8','multi','1');">添加考生</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"><input id="ROOM_ID" name="ROOM_ID" type="hidden" value="<%=ID%>"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
            <th width="10%" height="26">选择</th>
            <th width="12%" height="26">座位号</th>
            <th width="21%" height="26">考生</th>
            <th width="18%" height="26">身份证号</th>
            <th width="25%" height="26">操作</th>
            </tr>
          

<%
sql="SELECT R.ROOM_ID,R.TEST_ID,E.USER_ID,R.SEAT_ID,E.REAL_NAME FROM T_EXA_ROOM_SEAT_ID AS R Left join T_EXAMINEE AS E ON R.USER_ID=E.USER_ID where R.ROOM_ID="&ID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	If(rs.eof or rs.bof) Then
	 
%>   
          <tr>
            <td height="18"  colspan=5">[没有查到记录！]</div></td>
          </tr>

<%
Else
	do while not (rs.eof or err)

	%>    
          <tr>
            <td height="18" ><div align="center" >
              <input name="subBox" type="checkbox" value="<%=rs(3)%>" />
            </div></td>
            <td height="18" ><%=rs(3)%></td>
            <td height="18" ><%=rs(4)%></td>
            <td height="18" ><%=rs(2)%></td>
            <td height="18" >
            
            <%If isNULL(rs(2)) Then %>
            <img src="../Images/001.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location.href='exam_room_seat_add.asp?id=<%=rs(3)%>&room_id=<%=ID%>&change_seat=0&&pre_test_id=<%=test_id%>&pre_sys_exam_id=<%=sys_exam_id%>'">[添加考生]</a>
            <%Else%>
            <img src="../Images/change.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location.href='exam_room_seat_add.asp?id=<%=rs(3)%>&room_id=<%=ID%>&change_seat=1&pre_test_id=<%=test_id%>&pre_sys_exam_id=<%=sys_exam_id%>'">[座位调换]</a>
            <%End If%>
             </td>
            </tr>
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
<%appEnd()%>