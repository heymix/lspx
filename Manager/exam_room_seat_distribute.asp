<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
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


//锁定  
function statusPost(status,type,checkType){

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}	
		$.post(	"../Public/exam_room_seat_distributeStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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


function insertSeat(type,id,t_max_num){
	var typeName="";
	if (type=="0")  typeName="普通考生"
	else typeName="补考考生"
		if(!confirm(" 将未分配座位的"+typeName+" 随机插入空座位...")) return false;
		$.post(	"../Public/autoInsertSeat.asp", 
				{subBox:id,status:type,t_max_num:t_max_num}, 
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
<style>
#examNum{
	position:relative;
	float:left;
	vertical-align:middle;
	margin-top:3px;
}

</style>
</head>

<body>
<!--#include file="test_id_check.asp"-->
<!--#include file="exam_id_check.asp"-->
<%
examinee_all=conn.execute("select COUNT(*) from V_EXAMINEE_TEST where TRCHECKED=1 and TEST_ID="&test_id&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&") and USER_ID not  in (select USER_ID From T_EXA_ROOM_SEAT_ID where TEST_ID="&test_id&")")(0)
examinee_normal=conn.execute("select COUNT(*) from V_EXAMINEE_TEST where TRCHECKED=1 and TEST_ID="&test_id&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&") and USER_ID not  in (select USER_ID From T_EXA_ROOM_SEAT_ID where TEST_ID="&test_id&") and IS_MAKEUP=0")(0)
examinee_makeup=conn.execute("select COUNT(*) from V_EXAMINEE_TEST where TRCHECKED=1 and TEST_ID="&test_id&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&") and USER_ID not  in (select USER_ID From T_EXA_ROOM_SEAT_ID where TEST_ID="&test_id&") and IS_MAKEUP=1")(0)
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">考场号列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><div id="examNum">未分配座位的考生共<%=examinee_all%>人 其中普通考生<%=examinee_normal%>人 补考生<%=examinee_makeup%>人</div><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/run.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="statusPost('11','multi','0');">生成考号</a> <img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('11','multi','1');">撤销生成</a> </td>
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
            <th width="5%" height="26">选择</th>
            <th width="28%" height="26">教室</th>
            <th width="5%" height="26">考场号</th>
            <th width="6%" height="26">座位数</th>
            <th width="14%" height="26">已分配座位</th>
            <th width="14%" height="26">考号生成</th>
            <th width="21%" height="26">操作</th>
            </tr>
          

<%
Dim ors
Set ors=new Cls_ShowoPage	'创建对象
With ors
	.Conn=conn			'得到数据库连接对象
	.DbType="AC"
	'数据库类型,AC为access,MSSQL为sqlserver2000,MSSQL_SP为存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	.RecType=0
	'取记录总数方法(0执行count,1自写sql语句取,2固定值)
	.RecSql=0
	'如果RecType＝1则=取记录sql语句，如果是2=数值，等于0=""
	.RecTerm=2
	'取从记录条件是否有变化(0无变化,1有变化,2不设置cookies也就是及时统计，适用于搜索时候)
	.CookieName=""	'如果RecTerm＝2,cookiesname="",否则写cookiesname
	.Order=0			'排序(0顺序1降序),注意要和下面sql里面主键ID的排序对应
	.PageSize=20	'每页记录条数
	.JsUrl="../Inc/"			'showo_page.js的路径
	.Sql="ER.ID,ER.TEST_ID,ER.EXA_ID,ER.T_ROOM,ER.T_ROOM_ORDER,ER.T_MAX_NUM,ER.CHECKED,ER.IS_RUN$[T_EXA_ROOM] AS ER Left join T_EXA AS E ON ER.EXA_ID=E.ID $1=1 And ER.TEST_ID="&TEST_ID&" and EXA_ID="&sys_exam_id&" "&findKey&"$ER.ID asc$ER.ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>   
          <tr>
            <td height="18"  colspan="7">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	%>    
          <tr>
            <td height="18" ><div align="center" >
              <input name="subBox" type="checkbox" value="<%=iRs(0,i)%>" />
            </div></td>
            <td height="18" ><%=iRs(3,i)%></td>
            <td height="18" ><%=iRs(4,i)%></td>
            <td height="18" ><%=iRs(5,i)%></td>
            <td height="18" >(<%=conn.execute("SELECT COUNT(*)FROM [T_EXA_ROOM_SEAT_ID] where USER_ID<>'' and ROOM_ID="&iRs(0,i))(0)%>)
           </td>
            <td height="18" ><%=getTrue(iRs(7,i))%></td>
            <td height="18" ><%If iRs(7,i)=1 Then %>
			考号已生成 <img src="../Images/seek.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location.href='exam_room_seat_order.asp?id=<%=iRs(0,i)%>'">[考场对应单]</a>
			<%Else%><img src="../Images/114.gif" width="12" height="12" /><a href="javascript:void(0)" onClick="insertSeat('0','<%=iRs(0,i)%>','<%=iRs(5,i)%>');">[插入普通]</a> <img src="../Images/114.gif" width="12" height="12" /><a href="javascript:void(0)" onClick="insertSeat('1','<%=iRs(0,i)%>','<%=iRs(5,i)%>');">[插入补考]</a><img src="../Images/a1.gif" width="12" height="12" /><a href="javascript:void(0)" onClick="window.location.href='exam_room_seat.asp?id=<%=iRs(0,i)%>'">[查看分配]</a>
			<%End if%></td>
            </tr>
       
    <%
    Next	
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
<%ors.ShowPage()%>
<%
iRs=NULL
ors=NULL
Set ors=NoThing

%>