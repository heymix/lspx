<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()

schoolID		=request("schoolID")
inputString		=request("inputString")
USER_NAME_find	=request("USER_NAME")
examID			=request("examID")

If schoolID<>"" and isNumeric(schoolID) Then
	findKey=" and A.SCHOOL_ID="&schoolID
End If

If examID<>"" and isNumeric(examID) Then
	findKey=findKey&" and A.EXA_ID="&examID
End If

If USER_NAME_find<>"" and not sqlInjection(USER_NAME_find) Then
	findKey=findKey&" and A.USER_NAME like'%"&USER_NAME_find&"%'"
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
  
  
  //单选框操作
  	$('#editForm input:radio').each(function(){;
			 $(this).click(function(){
				if(this.checked){
					$("#shoolPanel").hide();
					$("#examPanel").hide();
					$('#rMark').val($(this).val());
					if($(this).val()=="2"){
						$("#shoolPanel").show();
					}
					if($(this).val()=="3"){
						$("#examPanel").show();
					}
					
				} 
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


    function lookup(inputString) {
        if(inputString.length == 0) {
            // Hid''e the suggestion box.
            $('#suggestions').hide();
        } else {
            $.post("../Public/FindListKey.asp", {queryString: inputString}, function(data){
                if(data.length >0) {
                    $('#suggestions').show();
                    $('#autoSuggestionsList').html(data);
                }
            });
        }
    } // lookup
    
    function fill(thisValue,id) {
        $('#inputString').val(thisValue);
		$('#schoolID').val(id);
        setTimeout("$('#suggestions').hide();", 200);
    }
	function resetForm(){
		$('#editForm')[0].reset();
	}
	function submitForm(){

		$("#editForm").submit();
	
	
	}
function resetPsd(id,opType){
	if(opType=="multi"){
	   var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
	}
	if(!confirm("确认重置密码吗? 不可恢复...")) return false;
	
	$.post(	"../Public/resetPsd.asp", 
				$("#listForm").serialize(), 
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

//锁定  checkType 0 正常 1 锁定
function checkPost(id,type,checkType){

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		
		$.post(	"../Public/managerChecked.asp?checkType="+checkType+"&operateType="+type, 
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

function delPost(id,type){

	if(!confirm("确认删除数据吗? 不可恢复...")) return false;

	if(type=="multi"){
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		$.post(	"../Public/managerDel.asp", 
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
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>管理员管理</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)">查询</a> <img src="../Images/001.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="window.location='manager_add.asp'">添加记录</a></td>
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
       
<label for="USER_NAME">用户名:</label>
<input type="text" id="USER_NAME" name="USER_NAME" value="<%=USER_NAME_find%>" />
<div id="examPanel" class="suggestionsMain" >
		<label for="examID">输入考点:</label> 		
        		<select id="examID" name="examID">
                <option value="0">选择考点</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,EXA_NAME from T_EXA"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>"<%If Cstr(rs("ID"))=examID then Response.write "selected"%>><%=rs("EXA_NAME")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select>         
    </div><br/>
   
    <div id="shoolPanel" class="suggestionsMain" >
    <label for="inputString">输入学校:</label> <input type="text" size="30" value="<%=inputString%>" id="inputString" name="inputString" onKeyUp="lookup(this.value);" onBlur="fill();"  autocomplete="off"  />
                    <input  type="hidden" name="schoolID" id="schoolID" value="<%=schoolID%>">
            
        <div class="suggestionsBox" id="suggestions">
            <div class="suggestionList" id="autoSuggestionsList">
                &nbsp;
            </div>
        </div>          
    </div></br>
    
    
        
    
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">管理员列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/locked.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="checkPost('-1','multi','0');">锁定</a> <img src="../Images/unlocked.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="checkPost('-1','multi','1');">解锁</a> <img src="../Images/key.gif" width="16" height="14" /><a href="javascript:vodi(0)" onClick="resetPsd('-1','multi')">重置密码</a> <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="delPost('-1','multi');">删除选中</a></font> <img src="../Images/excel.gif" width="14" height="14" /><font><a href="javascript:void(0)">导出Excel</a></font></td>
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
            <th width="4%" height="26">序号</th>
            <th width="21%" height="26">用户名</th>
            <th width="9%" height="26">级别类型</th>
            <th width="28%" height="26">校区/考点/中心</th>
            <th width="11%" height="26">最后登录时间</th>
            <th width="11%" height="26">最后登录IP</th>
            <th width="11%" height="26">用户状态</th>
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
	.Sql="A.USER_NAME,A.PURVIEW,A.SCHOOL_ID,S.SCHOOL,A.EXA_ID,E.EXA_NAME,A.LAST_LOGIN_IP,A.LAST_LOGIN_TIME,A.LOGIN_TIMES,A.CHECKED,A.ID$T_ADMIN AS A left join T_SCHOOL AS S ON A.SCHOOL_ID=S.ID Left Join T_EXA AS E on A.EXA_ID=E.ID$1=1 "&findKey&"$A.USER_NAME desc$A.USER_NAME" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>   
          <tr>
            <td height="18"  colspan="9">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	If(iRs(1,i)="0") then
		w_sort="考试中心"
		w_sotr_name="考试中心"
	End If
	If(iRs(1,i)="2") then
		w_sort="学校"
		w_sotr_name=iRs(3,i)
	End If
	If(iRs(1,i)="3") then
		w_sort="考点"
		w_sotr_name=iRs(5,i)
	End If
	%>    
          <tr>
            <td height="18" ><div align="center" >
              <input name="subBox" type="checkbox" value="<%=iRs(10,i)%>" />
            </div></td>
            <td height="18" ><%=i+1%></td>
            <td height="18" ><%=iRs(0,i)%></td>
            <td height="18" ><%=w_sort%></td>
            <td height="18" ><%=w_sotr_name%></td>
            <td height="18" ><%=Format_Time(FromUnixTime(iRs(7,i),+8),1)%></td>
            <td height="18" ><%=iRs(6,i)%></td>
            <td height="18" ><%=getLocked(iRs(9,i))%></td>
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