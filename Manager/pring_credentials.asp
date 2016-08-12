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
	if (status=="8"){
		if(!confirm("确认取消所以生成的编号吗? 不可恢复...")) return false;
	}


		$.post(	"../Public/pring_Status.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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


//取消所有编号
function cancelNO(status,type,checkType){



		if(!confirm("确认取消所以生成的编号吗? 不可恢复...")) return false;
	


		$.post(	"../Public/pring_cancelNO.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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


//取消所有编号
function cancelNO_muti(status,type,checkType){



		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		if(!confirm("确定要取消编号吗.多选取消..")) return false;
	


		$.post(	"../Public/pring_cancelNO_Muti.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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


//按当前考点生成
function creatNO(status,type,checkType){

		if(!confirm("确定要生成编号吗.生成当前考点的编号..")) return false;


		$.post(	"../Public/pring_CreatNO.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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

//多选生成
function creatNO_muti(status,type,checkType){
	
		
		var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
		if(!confirm("确定要生成编号吗.多选生成..")) return false;


		$.post(	"../Public/pring_CreatNO_Muti.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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
function openWin(){
	var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
window.open ('about:blank', 'newwindow', 'height=768, width=830, top=0, left=0, toolbar=yes, menubar=yes, scrollbars=yes, resizable=no,location=no, status=no')
$("#listForm").submit();
}


</script>
</head>

<body  >
<!--#include file="test_id_check.asp"-->
<!--#include file="exam_id_check.asp"-->
<%
REAL_NAME	 	= Request("REAL_NAME")
If REAL_NAME<>"" and not sqlInjection(REAL_NAME) Then
	findKey=findKey&" and REAL_NAME like'%"&REAL_NAME&"%'"
End If

USER_ID	 	= Request("USER_ID")
If USER_ID<>"" and not sqlInjection(USER_ID) Then
	findKey=findKey&" and USER_ID like'%"&USER_ID&"%'"
End If
TEL	 	= Request("TEL")
If TEL<>"" and not sqlInjection(TEL) Then
	findKey=findKey&" and TEL like'%"&TEL&"%'"
End If

MOBILE	 	= Request("MOBILE")
If MOBILE<>"" and not sqlInjection(MOBILE) Then
	findKey=findKey&" and MOBILE like'%"&MOBILE&"%'"
End If

EMAIL	 	= Request("EMAIL")
If EMAIL<>"" and not sqlInjection(EMAIL) Then
	findKey=findKey&" and MOBILE like'%"&EMAIL&"%'"
End If

CHECKED	 	= Request("CHECKED")
If CHECKED="1" or CHECKED="0"  Then
	findKey=findKey&" and CHECKED="&CHECKED
End If

LOCKED	 	= Request("LOCKED")
If LOCKED="1" or LOCKED="0"  Then
	findKey=findKey&" and LOCKED="&LOCKED
End If

IS_PRINT	 	= Request("IS_PRINT")
If IS_PRINT="1" or IS_PRINT="0"  Then
	findKey=findKey&" and  IS_PRINT="&IS_PRINT
End If

IS_MAKEUP	 	= Request("IS_MAKEUP")
If IS_MAKEUP="1" or IS_MAKEUP="0"  Then
	findKey=findKey&" and  IS_MAKEUP="&IS_MAKEUP
End If

SCHOOL_ID	 	= Request("SCHOOL_ID")
If SCHOOL_ID<>"0" and is_Id(SCHOOL_ID)  Then
	findKey=findKey&" and SCHOOL_ID="&SCHOOL_ID
End If

If sys_exam_id<>"0" and is_Id(sys_exam_id)  Then
	findKey=findKey&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&")"
End If

TEST_NO	 	= trim(Request("TEST_NO"))
If is_Id(TEST_NO)  Then
	findKey=findKey&" and USER_ID in (select USER_ID From T_TEST_RESULT where TEST_NO="&TEST_NO&" and test_id="&test_id&")"
End If

CRS_NO	 	= trim(Request("CRS_NO"))
If is_Id(CRS_NO)  Then
	findKey=findKey&" and CRS_NO="&CRS_NO
End If

%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>学生管理</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</a></td>
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
       
        <label for="REAL_NAME">姓名:</label>
        <input type="text" id="REAL_NAME" name="REAL_NAME" value="<%=REAL_NAME%>" />
         <label for="USER_ID">身份证号:</label>
        <input type="text" id="USER_ID" name="USER_ID" value="<%=USER_ID%>" />
        <label for="CHECKED">审核:</label>
        <select id="CHECKED" name="CHECKED">
        	<option value="-1">未选择</option>
            <option value="0"<%If CHECKED="0" Then Response.write "selected='selected'"%>>未审核</option>
            <option value="1"<%If CHECKED="1" Then Response.write "selected='selected'"%>>已审核</option>
        </select>
		<br/>
        <label for="SCHOOL_ID">报考学校:</label>
            <select id="SCHOOL_ID" name="SCHOOL_ID" >
                <option value="0">选择学校</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,SCHOOL from T_SCHOOL where CHECKED=1 and ID in(select SCHOOL_ID from T_EXA_SCHOOL_RELATION  where EXA_ID="&sys_exam_id&")  order by SCHOOL asc"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If Cstr(rs("ID"))=SCHOOL_ID Then Response.write "selected=""selected"""  %>><%=rs("SCHOOL")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                </select>
         <label for="TEL">电话:</label>
        <input type="text" id="TEL" name="TEL" value="<%=TEL%>" />
        <label for="LOCKED">锁定:</label>
        <select id="LOCKED" name="LOCKED">
        	<option value="-1">未选择</option>
            <option value="0"<%If LOCKED="0" Then Response.write "selected='selected'"%> >正常</option>
            <option value="1"<%If LOCKED="1" Then Response.write "selected='selected'"%>>锁定</option>
        </select>
		<br/>
         <label for="MOBILE">手机:</label>
        <input type="text" id="MOBILE" name="MOBILE" value="<%=MOBILE%>" />
        <label for="IS_PRINT">打印标记:</label>
        <select id="IS_PRINT" name="IS_PRINT">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_PRINT="0" Then Response.write "selected='selected'"%>>否</option>
            <option value="1"<%If IS_PRINT="1" Then Response.write "selected='selected'"%>>是</option>
        </select>
         <label for="IS_MAKEUP">补考:</label>
        <select id="IS_MAKEUP" name="IS_MAKEUP">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_MAKEUP="0" Then Response.write "selected='selected'"%>>否</option>
            <option value="1"<%If IS_MAKEUP="1" Then Response.write "selected='selected'"%>>是</option>
        </select>
		<br/>
             <label for="EMAIL">e-mail:</label>
        <input type="text" id="EMAIL" name="EMAIL" value="<%=EMAIL%>" />
         <label for="TEST_NO">准考证号:</label>
        <input type="text" id="TEST_NO" name="TEST_NO" value="<%=TEST_NO%>" />
        <label for="CRS_NO">证书编号:</label>
        <input type="text" id="CRS_NO" name="CRS_NO" value="<%=CRS_NO%>" />
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
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">考生列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/is_top_2.gif" width="14" height="14" /><font><a href="javascript:void(0)"   onClick="statusPost('5','multi','-1');">打印标记</a> <img src="../Images/is_uptop.gif" width="14" height="14" /><font><a href="javascript:void(0)"   onClick="statusPost('6','multi','-1');">取消标记</a><img src="../Images/exam.gif" width="14" height="14" /><a href="javascript:void(0)"  onClick="creatNO('7','multi','-1');">按考点生成编号</a>  <img src="../Images/exam.gif" width="14" height="14" /><a href="javascript:void(0)"  onClick="creatNO_muti('7','multi','-1');">多选生成编号</a>  <img src="../Images/is_uptop.gif" width="14" height="14" /><font><a href="javascript:void(0)"   onClick="cancelNO('8','multi','-1');">取消所有编号</a> <img src="../Images/is_uptop.gif" width="14" height="14" /><font><a href="javascript:void(0)"   onClick="cancelNO_muti('8','multi','-1');">多选取消编号</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post" action="print_CRS.asp" target="newwindow"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
            <th width="5%" height="26">选择</th>
            <th width="6%" height="26">姓名</th>
            <th width="9%" height="26">身份证号</th>
            <th width="7%" height="26">准考证号</th>
            <th width="8%" height="26">证书编号</th>
            <th width="3%" height="26">性别</th>
            <th width="8%" height="26">报考学校</th>
            <th width="7%" height="26">电话</th>
            <th width="7%" height="26">手机</th>
            <th width="6%" height="26">Email</th>
            <th width="28%" height="26">考试科目</th>
            <th width="6%" height="26">是否打印</th>
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
	.PageSize=100	'每页记录条数
	.JsUrl="../Inc/"			'showo_page.js的路径
	.Sql="USER_ID,SCHOOL_ID,SCHOOL,GENDER,REAL_NAME,AGE,EDUCATION,TEL,MOBILE,EMAIL,LAST_LOGIN_TIME,CHECKED,LOCKED,IS_MAKEUP,TRCHECKED,CRS_NO,IS_PRINT$V_EXAMINEE_TEST$1=1 and TEST_ID="&test_id&"  "&findKey&"$TEST_NO asc$USER_ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
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
            <td height="18" ><%=iRs(4,i)%></td>
            <td height="18" ><%=iRs(0,i)%></td>
            <td height="18" ><%=getTestNo(iRs(0,i))%></td>
            <td height="18" ><%=iRs(15,i)%></td>
            <td height="18" ><%=getGender(iRs(3,i))%></td>
            <td height="18" ><%=iRs(2,i)%></td>
            <td height="18" ><%=iRs(7,i)%></td>
            <td height="18" ><%=iRs(8,i)%></td>
            <td height="18" ><%=iRs(9,i)%></td>
            <td height="18" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
            	<th width="47%">科目</th>
                <th width="11%">分数</th>
                <th width="10%">及格</th>
                <th width="16%">是否补考</th>
                <th width="16%">是否公布</th>
            </tr>
            <%set rs_topic=server.createobject("ADODB.Recordset")
			 sql="SELECT [USER_ID],[TEST_ID],[TOPIC_ID],[RESULT],[IS_PASS],[IS_MAKEUP],[IS_SCORE],[TITLE],[TNAME],[IS_PRATICE] FROM [V_SCORE] where USER_ID='"&iRs(0,i)&"' and TEST_ID="&test_id
			 rs_topic.open sql,conn,1,3
			 	delStr="&nbsp;"
			  	If rs_topic.eof or rs_topic.bof Then
			 		Response.write "<li>没有记录</li>"
				Else
					
					do while not (rs_topic.eof or err)
					If rs_topic("IS_PRATICE")=1 Then
					 	If Cstr(rs_topic("RESULT"))="1" Then
							exam_score="合格"
						Else
							exam_score="不合格"
						End If
					Else 
						exam_score=rs_topic("RESULT")
					End If
					If rs_topic("IS_SCORE")=0 Then exam_score=""
						Response.write "<tr><td>"&rs_topic("TNAME")&"</td><td>"&exam_score&"</td><td>"&getTrue(rs_topic("IS_PASS"))&"</td><td>"&getTrue(rs_topic("IS_MAKEUP"))&"</td><td>"&getTrue(rs_topic("IS_SCORE"))&"</td></tr>"
				   	rs_topic.movenext
				   	loop
			  	End If
			   rs_topic.close
			   set rs_topic=nothing
			   %>
            
            </table></td>
            <td height="18" ><%=getTrue(iRs(16,i))%></td>
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