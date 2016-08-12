<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime		=timer()

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
  
  $("#checkAll_ispass").click(function(e) {
	var check = $("input:[name='subBox']"); //得到所有被选中的checkbox   
	var actor_config;             			//定义变量 
	//$('#checkAll').attr("checked","checked"); 
	 check.each(function(i){
		  actor_config=$(this).val().split("|")
	 	if($("#checkAll_ispass").attr("checked")=="checked"){
			$("#"+actor_config[0]+actor_config[1]+"_check").attr("checked",true); 
		}
		else{ 
			$("#"+actor_config[0]+actor_config[1]+"_check").attr("checked",false);
		}          
	 });   

});
  		 $("#inverse").click(function () {//反选
                $("input[name='subBox']").each(function () {
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
					 	//$(this).addClass('checkBgColor');
				  }
				  else{
					// $(this).removeClass('checkBgColor');
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


function savePost(){
	var n = $("input[name='subBox']:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		$.post(	"../Public/scoreEdit.asp", 
				$("#listForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							//window.location.reload();
						}
						else{
							alert(resultArr[1]);
						}	
				});
		
	
		
}
function changePageSize(val){
	if(!confirm("确定要改变页码吗? 将重新排序...")) return false;
	window.location.href='score.asp?pSize='+val;
}

function cc(){
	$("#listForm select option[value='0']").attr("selected", true);
}




</script>
<style>
.score {width:100px;}
.score2{width:105px;}

</style>

</head>

<body  >
<!--#include file="test_id_check.asp"-->

<%

pSize=request("pSize")
If pSize<>"" or is_Id(pSize) Then
	response.Cookies("page")("pSize")=pSize
End If
resetPage= request.Cookies("page")("pSize")
If resetPage="" or not is_Id(resetPage) Then
	resetPage=20
End If



TOPIC_ID	 	= Request("TOPIC_ID")
If TOPIC_ID<>"-1" and is_Id(TOPIC_ID)  Then
	findKey=findKey&" and TOPIC_ID ="&TOPIC_ID
End If

SCORE	 	= Request("SCORE")
If SCORE<>"-1" and is_Id(SCORE)  Then
	If SCORE="2" Then
		findKey=findKey&" and RESULT <60 and RESULT >=0"
	ElseIf SCORE="3" Then
		findKey=findKey&" and RESULT >=60"
	ElseIf SCORE="1" Then
		findKey=findKey&" and RESULT =1"
	ElseIf SCORE="0" Then
		findKey=findKey&" and RESULT =0"
	End If
	
	
End If

TEST_NO	 	= trim(Request("TEST_NO"))
If is_Id(TEST_NO)  Then
	findKey=findKey&" and USER_ID in (select USER_ID From T_TEST_RESULT where TEST_NO="&TEST_NO&" and test_id="&test_id&")"
End If

IS_PASS	 	= Request("IS_PASS")
If IS_PASS="1" or IS_PASS="0"  Then
	findKey=findKey&" and IS_PASS ="&IS_PASS
End If
findKey=findKey&" and SCHOOL_ID in(select SCHOOL_ID From T_EXA_SCHOOL_RELATION AS ESR Left Join T_TEST_EXA_RELATION AS TER on ESR.EXA_ID=TER.EXA_ID where TER.TEST_ID="&test_id&")"
%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>学生管理</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)">[20]</a> <a href="javascript:void(0)">[40]</a> <a href="javascript:void(0)">[60]</a>&nbsp;</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="get" id="editForm" name="editForm"   action="?1=1" >
        <input id="operateType" name="operateType" value="0" type="hidden">
       
      <br/>

         <label for="TOPIC_ID">按科目:</label>
        <select id="TOPIC_ID" name="TOPIC_ID">
        	<option value="-1">未选择</option>
             <%
					set rs=server.createobject("adodb.recordset")
					sql="select T.ID,T.TNAME from T_TEST_TOPIC_RELATION AS R Left Join T_TEST_TOPIC AS T ON R.TOPIC_ID=T.ID where 1=1 and R.TEST_ID="&TEST_ID
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If Cstr(rs("ID"))=TOPIC_ID Then Response.write "selected=""selected"""  %>><%=rs("TNAME")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
        </select>
		<br/>
        <label for="SCORE">按成绩:</label>
        <select id="SCORE" name="SCORE">
        	<option value="-1">未选择</option>
            <option value="2" <%If SCORE="2" Then Response.write "selected='selected'"%>>0-60</option>
            <option value="3" <%If SCORE="3" Then Response.write "selected='selected'"%>>60-100</option>  
            <option value="1" <%If SCORE="1" Then Response.write "selected='selected'"%>>合格</option>
            <option value="0" <%If SCORE="0" Then Response.write "selected='selected'"%>>不合格</option> 
        </select>
		<br/>

         <label for="IS_PASS">及格:</label>
        <select id="IS_PASS" name="IS_PASS">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_PASS="0" Then Response.write "selected='selected'"%>>未及格</option>
            <option value="1"<%If IS_PASS="1" Then Response.write "selected='selected'"%>>及格</option>
        </select>
        <br/><br/>
        
   
    
    
        
    
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
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">成绩列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"> &nbsp;&nbsp;&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post" action="../Public/scoreEdit.asp"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
          <th width="7%" height="26">选择<input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" /></th>
            <th width="26%" height="26">身份证号</th>
            <th width="15%" height="26">准考证号</th>
            <th width="12%" height="26">姓名</th>
            <th width="9%" height="26">科目</th>
            <th width="19%" height="26">分数<input id="checkAll_isOK" type="checkbox" name="checkbox62" value="checkbox" onClick="cc()" /> 选中不合格</th>
            <th width="6%" height="26">及格<input id="checkAll_ispass" type="checkbox" name="checkAll_ispass" /></th>
            <th width="6%" height="26">是否补考</th>
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
	.PageSize=resetPage	'每页记录条数
	.JsUrl="../Inc/"			'showo_page.js的路径
	.Sql="[USER_ID],[TEST_ID],[TOPIC_ID],[RESULT],[IS_PASS],[IS_MAKEUP],[IS_SCORE],[REAL_NAME],[TNAME],[TITLE],[ID],[IS_PRATICE]$V_SCORE$1=1 and TEST_ID="&TEST_ID&" "&findKey&"$TEST_NO asc$USER_ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>  
          <tr>
            <td height="18"  colspan="10">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	%>    
          <tr>
          <td><input name="subBox" type="checkbox" value="<%=iRs(0,i)%>|<%=iRs(10,i)%>" /></td>
            <td height="18" ><%=iRs(0,i)%></td>
            <td height="18" ><%=getTestNo(iRs(0,i))%></td>
            <td height="18" ><%=iRs(7,i)%></td>
            <td height="18" ><%=iRs(8,i)%></td>
            <td height="18" ><%If iRs(11,i)=0 Then%><input class="score" id="<%=iRs(0,i)%><%=iRs(10,i)%>" name="<%=iRs(0,i)%><%=iRs(10,i)%>" value="<%=iRs(3,i)%>" onChange="if(isNaN($(this).val())|| $(this).val().replace(' ','').length==0||$(this).val()<0||  $(this).val()>100){alert('<%=iRs(7,i)%>的<%=iRs(8,i)%>分数录入不对！');$(this).val(0)}" tabindex="<%=i+1%>" />
            <%Else%>
            	<select class="score2"  id="<%=iRs(0,i)%><%=iRs(10,i)%>" name="<%=iRs(0,i)%><%=iRs(10,i)%>">
                	<option value="1" <%If Cint(iRs(3,i))=1 Then Response.write "selected='selected'"%>>合格</option>
                	<option value="0" <%If Cint(iRs(3,i))=0 Then Response.write "selected='selected'"%>>不合格</option>
                    
                </select>
            <%End If%>
            
            </td>
            <td height="18" ><input type="checkbox" id="<%=iRs(0,i)%><%=iRs(10,i)%>_check" name="<%=iRs(0,i)%><%=iRs(10,i)%>_check" value="1" <%if iRs(4,i)=1 Then Response.write "checked" %>></td>
            <td height="18" ><%=getTrue(iRs(5,i))%></td>
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