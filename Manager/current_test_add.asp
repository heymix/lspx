<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()

oBy=request("oBy")
sKey=request("sKey")
If oBy="" or oBy="asc" Then
	orderKey=" E.EXA_NAME ASC"
Else
	orderKey=" E.EXA_NAME DESC"
End If


If sKey<>"" and not sqlInjection(sKey)Then
	findKey=" And E.EXA_NAME like'%"&sKey&"%'"
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


//锁定  status 0 审校 1置顶 2刷新 4删除
function statusPost(status,type,checkType){

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
	if (status=="4"){
		if(!confirm("确认删除数据吗? 不可恢复...")) return false;
	}
		
		$.post(	"../Public/current_testStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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

<body  >
<!--#include file="test_id_check.asp"-->
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/back_01.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location='current_test.asp'"> 返回</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="get" id="editForm" name="editForm"  title="考点管理" action="?1=1" >
        <input id="operateType" name="operateType" value="0" type="hidden">
       
        <label for="sKey">标题:</label>
        <input type="text" id="sKey" name="sKey" value="<%=sKey%>" />
		<br/>
   
    
    
        
    
            <div  class="suggestionsMain" style="text-align:right">
 					<img src="../Images/005.gif" width="14" height="14" /><a id="submitLink" href="javascript:void(0)" onClick="submitForm();">[查询]</a> <img src="../Images/002.gif" width="14" height="14"/><a href="javascript:void(0)" onClick="resetForm();">[重置]</a>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">未在"<%=test_title%>" 中的考点</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/001.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="statusPost('6','multi','<%=TEST_ID%>');">添加</a></font> <img src="../Images/excel.gif" width="14" height="14" /><font><a href="javascript:void(0)">导出Excel</a></font></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post">
          <table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
            <tr>
              <th width="12%" height="26">选择</th>
              <th width="45%" height="26">考点<%If oBy="" or oBy="asc" THen%>
                <a href="javascript:void(0)" onClick="javasscript:window.location='?oBy=desc&amp;sKey=<%=sKey%>'"><img src="../Images/hmenu-asc.gif" width="16" height="16" /></a>
                <%Else%>
                <a href="javascript:void(0)" onClick="javasscript:window.location='?oBy=asc&amp;sKey=<%=sKey%>'"><img src="../Images/hmenu-desc.gif" width="16" height="16" /></a>
                <%End If%></th>
              <th width="13%" height="26">管理员</th>
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
	.Sql="E.ID,E.EXA_NAME,E.TEST_CATE $T_EXA AS E$1=1 and  E.ID not in(SELECT E.ID FROM T_TEST_EXA_RELATION AS ER Left Join T_EXA AS E ON ER.EXA_ID=E.ID where ER.TEST_ID="&TEST_ID&") and E.TEST_CATE="&test_cate&" "&findKey&"$"&orderKey&"$E.ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With
iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>
            <tr>
              <td height="18"  colspan="3">[没有查到记录！]
                </div></td>
            </tr>
            <%
Else     
    For i=0 To Ubound(iRs,2)
	%>
            <tr>
              <td height="18" ><div align="center" >
                <input name="subBox" type="checkbox" value="<%=iRs(0,i)%>" />
              </div></td>
              <td height="18" ><span class="tdStyle "><%=iRs(1,i)%></span></td>
              <td height="18" ><img src="../Images/seek.gif" width="14" height="14" />[<a href="javascript:void(0)" onClick="window.location='manager.asp?schoolID=<%=iRs(0,i)%>&inputString=<%=iRs(1,i)%>'">查看</a>]</td>
              </tr>
            <%
    Next	
End If
%>
          </table>
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