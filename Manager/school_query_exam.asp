<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()
oBy=request("oBy")
sKey=request("sKey")
If oBy="" or oBy="asc" Then
	orderKey=" S.SCHOOL ASC"
Else
	orderKey=" S.SCHOOL DESC"
End If


If sKey<>"" and not sqlInjection(sKey)Then
	findKey=" And S.SCHOOL like'%"&sKey&"%'"
End If
'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>学校基础信息</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
<script src="../Jquery/jquery.bgiframe.min.js" type="text/javascript" charset='utf-8'></script>
<script src="../Jquery/loading-min.js" type="text/javascript"  charset="UTF-8"></script>

 <script type="text/javascript">
 $(document).ready(function(e) {
    loading=new ol.loading({id:"tablePanel"});

	
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

function subEdit(){
	var opType=$("#operateType").val()
	var sKey=$("#opValue").val()
	var sID =$("#sID").val()
	if(opType=="0"){
		window.location='?sKey='+sKey+'&oBy=<%=oBy%>';
	}
	if(opType=="2"){

		if($("#opValue").val()==""){
			alert("学校名称不能为空！");
			return false;
			}
		loading.show();
		$.post(	"../Public/schoolEdit.asp", 
				$("#editForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							loading.hide();
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
						}	
				});
	}
	if(opType=="1"){

		if($("#opValue").val()==""){
			alert("学校名称不能为空！");
			return false;
			}
		if(isNaN(sID)){
			alert("参数错误,请联系管理员！");
			return false;
		}
		loading.show();
		$.post(	"../Public/schoolEdit.asp", 
				$("#editForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							loading.hide();
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
						}	
				});
	}
}

//审核 id type single multi多选单选 checkType 0 复核 1 反复核
function checkPost(id,type,checkType){

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		loading.show();
		$.post(	"../Public/schoolChecked.asp?checkType="+checkType+"&operateType="+type, 
				$("#listForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							loading.hide();
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
						}	
				});
		
	}
	if(type=="single"){
		if(isNaN(id)||Number(id)<0){
			alert("ID错误,无法操作！");
			return false;
		}
		loading.show();
		$.post(	"../Public/schoolChecked.asp?checkType="+checkType+"&operateType="+type, 
				{subBox:+id}, 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							loading.hide();
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
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
		loading.show();
		$.post(	"../Public/schoolDel.asp", 
				$("#listForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
						}	
				});
		
	}
	if(type=="single"){
		if(isNaN(id)||Number(id)<0){
			alert("ID错误,无法操作！");
			return false;
		}
		loading.show();
		$.post(	"../Public/schoolDel.asp", 
				{subBox:+id}, 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							loading.hide();
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
							loading.hide();
						}	
				});
		
		
	}
		
}
</script>


</head>

<body >
<!--#include file="test_id_check.asp"-->
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>学校查询所在考点</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm" style="line-height:28px" >
        <label style="width:120px"><span id="opName">查询</span> 学校名称：</label><input type="text" id="opValue" name="opValue" style="width:200px" value="<%=sKey%>"><input type="hidden" value="0" id="operateType" name="operateType"> <input type="hidden" value="0" id="sID" name="sID"> <img src="../Images/g_page.gif" width="14" height="14" /> <a href="javascript:void(0)" onClick="subEdit();">提交</a></form></td>
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
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">学校列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <img src="../Images/excel.gif" width="14" height="14" /><font><a href="javascript:void(0)">导出Excel</a></font></td>
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
              <th width="11%" height="26">编号</th>
              <th width="32%" height="26">学校<%If oBy="" or oBy="asc" THen%>
                <a href="javascript:void(0)" onClick="javasscript:window.location='?oBy=desc&amp;sKey=<%=sKey%>'"><img src="../Images/hmenu-asc.gif" width="16" height="16" /></a>
                <%Else%>
                <a href="javascript:void(0)" onClick="javasscript:window.location='?oBy=asc&amp;sKey=<%=sKey%>'"><img src="../Images/hmenu-desc.gif" width="16" height="16" /></a>
                <%End If%></th>
              <th width="10%" height="26">考点编号</th>
              <th width="47%" height="26">考点</th>

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
	.Sql="S.ID,S.SCHOOL,E.ID AS EXA_ID,E.EXA_NAME$T_EXA_SCHOOL_RELATION AS SR Left Join T_SCHOOL AS S ON SR.SCHOOL_ID=S.ID Left Join T_EXA AS E on SR.EXA_ID=E.ID$1=1 and EXA_ID in(select EXA_ID from T_TEST_EXA_RELATION WHERE TEST_ID="&test_id&") "&findKey&"$"&orderKey&"$S.ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>
            <tr>
              <td height="18"  colspan="4">[没有查到记录！]
                </div></td>
            </tr>
            <%
Else     
    For i=0 To Ubound(iRs,2)
	%>
            <tr>
              <td height="18" ><div align="center" >
               <%=iRs(0,i)%>
              </div></td>
              <td height="18" ><span class="tdStyle "><%=iRs(1,i)%></span></td>
              <td height="18" ><%=iRs(2,i)%></td>
              <td height="18" ><%=iRs(3,i)%></td>

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