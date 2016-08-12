<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()


sKey=request("sKey")
If sKey<>"" and not sqlInjection(sKey)Then
	findKey=" And KEY_VALUE like'%"&sKey&"%'"
End If

sort_Id=request("sort_Id")
If sort_Id="1" Then
	cate_name="公告通知类别"
End If

If sort_Id="2" Then
	cate_name="下载类别"
End If

If sort_Id="3" Then
	cate_name="学历"
End If

If sort_Id="4" Then
	cate_name="考试类别"
End If

If sort_Id="5" Then
	cate_name="考点所在地"
End If
'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理员基础信息</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
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
   		$('#opValue').bind('keypress',function(event){
            if(event.keyCode == "13")    
            {
                subEdit();
            }
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


function subEdit(){
	var opType=$("#operateType").val()
	var sKey=$("#opValue").val()
	var sID =$("#sID").val()
	if(opType=="0"){
		
		window.location.href='?sKey='+sKey+'&sort_Id=<%=sort_Id%>';
	}
	if(opType=="2"){

		if($("#opValue").val()==""){
			alert("名称不能为空！");
			return false;
			}
	
		$.post(	"../Public/datakeyEdit.asp", 
				$("#editForm").serialize(), 
				function(data,st){
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							resetForm();
							alert(resultArr[1]);
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
						}	
				});
	}
	if(opType=="1"){

		if($("#opValue").val()==""){
			alert("名称不能为空！");
			return false;
			}
		if(isNaN(sID)){
			alert("参数错误,请联系管理员！");
			return false;
		}

		$.post(	"../Public/datakeyEdit.asp", 
				$("#editForm").serialize(), 
				function(data,st){
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							resetForm();
							alert(resultArr[1]);
							//window.location.reload();
						}
						else{
							alert(resultArr[1]);
						}	
				});
	}
}

function delPost(id,type){

	if(!confirm("确认删除数据吗? 不可恢复...")) return false;

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		$.post(	"../Public/datakeyDel.asp", 
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

function opType(type,name,id){

		$("#operateType").val(type)
	if(type=="0"){
		$("#opName").html("查询")
		$("#opValue").val('');
		$("#editForm>form").attr("title","ddd");
	}
	if(type=="1"){
		$("#opName").html("修改")
		$("#opValue").val(name);
		$("#sID").val(id);
	}
	if(type=="2"){
		$("#opName").html("添加");
		$("#opValue").val('');
		
	}
	
}
	function resetForm(){
		$('#editForm')[0].reset();
	}
</script>
</head>

<body  >
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>><%=cate_name%></span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)" onClick="opType('0','','')">查询</a> <img src="../Images/001.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="opType('2','','')">添加<%=cate_name%></a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm" style="line-height:28px" >
        <label style="width:120px"><span id="opName">查询</span> <%=cate_name%>：</label><input type="hidden" id="KEY_ID" name="KEY_ID" value="<%=sort_Id%>"><input type="hidden" id="KEY_CONTENT" name="KEY_CONTENT" value="<%=cate_name%>">
        <input type="text" id="opValue" name="opValue" style="width:200px" value="<%=sKey%>" ><input type="hidden" value="0" id="operateType" name="operateType"> <input type="hidden" value="0" id="sID" name="sID"> <img src="../Images/g_page.gif" width="14" height="14" /> <a href="javascript:void(0)" onClick="subEdit();">提交</a></form></td>
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
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle"><%=cate_name%>列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选  <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="delPost('-1','multi');">删除选中</a></font></td>
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
            <th width="21%" height="26">选择</th>
            <th width="15%" height="26">序号</th>
            <th width="49%" height="26"><%=cate_name%>名称</th>
            <th width="15%" height="26">操作</th>
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
	.Sql="ID,KEY_VALUE,KEY_ID,KEY_CONTENT,IS_SYSTEM$T_DATA_KEY$1=1 And KEY_ID="&sort_Id&" And IS_SYSTEM=0 "&findKey&"$ID desc$ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>   
          <tr>
            <td height="18"  colspan="4">[没有查到记录！]</div></td>
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
            <td height="18" ><img src="../Images/edit.gif" width="14" height="14" />[<a href="javascript:void(0)" onClick="opType('1','<%=iRs(1,i)%>','<%=iRs(0,i)%>')">编辑</a>]</td>
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