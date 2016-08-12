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

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
	if (status=="4"){
		if(!confirm("确认删除数据吗? 不可恢复...")) return false;
	}
		
		$.post(	"../Public/testStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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
		
}

</script>
</head>

<body  >
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>考试管理</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)">查询</a> <img src="../Images/001.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location='test_add.asp'">添加记录</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="get" id="editForm" name="editForm"  title="学校管理" action="?1=1" >
        <input id="operateType" name="operateType" value="0" type="hidden">
       
        <label for="TITLE">标题:</label>
        <input type="text" id="TITLE" name="TITLE" value="<%=TITLE%>" />
		<br/>
   
    
    
        
    
            <div  class="suggestionsMain" style="text-align:right">
 					<img src="../Images/005.gif" width="14" height="14" /><a id="submitLink" href="javascript:void(0)" onClick="submitForm();">[查询]</a> 
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
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">考试列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 
		  <img src="../Images/run.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('2','multi','0');">打印准考证</a> 
		  <img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('2','multi','1');">取消打印准考证</a>
		  <%if SYS_USER_ID=1 or SYS_USER_ID=181 then %>
			<img src="../Images/run.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('3','multi','0');">录入加锁</a> 
			<img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('3','multi','1');">录入解锁</a>
		  <%end if%>
		  
		  <img src="../Images/run.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="statusPost('0','multi','0');">审核</a> 
		  <img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('0','multi','1');">反审核</a> 
		  <img src="../Images/is_top_2.gif" width="14" height="14" /><a href="javascript:vodi(0)" onClick="statusPost('1','multi','0')">结束</a> 
		  <img src="../Images/is_uptop.gif" width="14" height="14" /><a href="javascript:vodi(0)" onClick="statusPost('1','multi','1')">激活</a> 
		  <img src="../Images/refresh_2.gif" width="14" height="14" /><font><!--a href="javascript:void(0)" onClick="statusPost('2','multi','-1');">刷新排序</a></font--> 
		  <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="statusPost('4','multi','-1');">删除选中</a></font></td>
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
            <th width="3%" height="26">选择</th>
            <th width="3%" height="26">序号</th>
            <th width="11%" height="26">标题</th>
            <th width="7%" height="26">报名开始时间</th>
            <th width="9%" height="26">报名结束时间</th>
            <th width="9%" height="26">考试注意事项</th>
            <th width="11%" height="26">考试科目</th>
            <th width="12%" height="26">考试类别</th>
            <th width="6%" height="26">审核</th>
			<th width="6%" height="26">打印准考证</th>
			<th width="6%" height="26">录入是否加锁</th>
            <th width="5%" height="26">状态-完结</th>
            <th width="24%" height="26">操作</th>
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
	.Sql="T.ID,T.TITLE,T.START_TIME,T.END_TIME,T.NOTICE_ID,N.TITLE,T.INFO_TIME,T.CHECKED,T.IS_OVER,T.TEST_CATE,D.KEY_VALUE,T.IS_PRINT,T.IS_EDIT$T_TEST AS T Left Join T_NOTICE AS N  ON T.NOTICE_ID=N.ID Left Join T_DATA_KEY As D On T.TEST_CATE=D.ID $1=1 "&findKey&"$T.ID desc$T.ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>   
          <tr>
            <td height="18"  colspan="12">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	%>    
          <tr>
            <td height="18" ><div align="center" >
              <input name="subBox" type="checkbox" value="<%=iRs(0,i)%>" />
            </div></td>
            <td height="18" ><%=i+1%></td>
            <td height="18" ><%=iRs(1,i)%></td>
            <td height="18" ><%=format_time(FromUnixTime(iRs(2,i),+8),2)%></td>
            <td height="18" ><%=format_time(FromUnixTime(iRs(3,i),+8),2)%></td>
            <td height="18" ><a href="javascript:void(0)"><%=iRs(5,i)%></a></td>
            <td height="18" ><ul class="sList">
            <%set rs=server.createobject("ADODB.Recordset")
			 sql="select T.ID,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID Where R.TEST_ID="&iRs(0,i)
			 rs.open sql,conn,1,3
			 	delStr="&nbsp;"
			  	If rs.eof or rs.bof Then
			 		Response.write "<li>没有记录</li>"
				Else
					do while not (rs.eof or err)
						Response.write "<li><span  class=""slMiddle"">"&rs("TNAME")&"</span></li>"
				   	rs.movenext
				   	loop
			  	End If
			   rs.close
			   set rs=nothing
			   %>
            </ul></td>
            <td height="18" ><%=iRs(10,i)%></td>
            <td height="18" ><%=getChecked(iRs(7,i))%></td>
            <td height="18" ><%=getTrue(iRs(11,i))%></td>
			<td height="18" ><%=getTrue(iRs(12,i))%></td>
			<td height="18" ><%=getTrue(iRs(8,i))%></td>
            <td height="18" ><img src="../Images/edit.gif" width="14" height="14" />[<a href="javascript:void(0)" onClick="window.location.href='test_add.asp?id=<%=iRs(0,i)%>'">编辑</a>]</td>
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