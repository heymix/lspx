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
		
		$.post(	"../Public/examineeStatus.asp?status="+status+"&checkType="+checkType+"&operateType="+type, 
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
function find_examinee(){
	var n = $("#listForm input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}
	listForm.submit();

}
</script>
</head>

<body  >
<%
REAL_NAME	 	= Request("REAL_NAME")
If REAL_NAME<>"" and not sqlInjection(REAL_NAME) Then
	findKey=findKey&" and E.REAL_NAME ='"&REAL_NAME&"'"
End If
USER_ID	 	= Request("USER_ID")
If USER_ID<>"" and not sqlInjection(USER_ID) Then
	findKey=findKey&" and E.USER_ID ='"&USER_ID&"'"
End If

TEL	 	= Request("TEL")
If TEL<>"" and not sqlInjection(TEL) Then
	findKey=findKey&" and E.TEL ='"&TEL&"'"
End If

MOBILE	 	= Request("MOBILE")
If MOBILE<>"" and not sqlInjection(MOBILE) Then
	findKey=findKey&" and E.MOBILE ='"&MOBILE&"'"
End If

EMAIL	 	= Request("EMAIL")
If EMAIL<>"" and not sqlInjection(EMAIL) Then
	findKey=findKey&" and E.EMAIL ='"&EMAIL&"'"
End If

CHECKED	 	= Request("CHECKED")
If CHECKED="1" or CHECKED="0"  Then
	findKey=findKey&" and E.CHECKED="&CHECKED
End If

LOCKED	 	= Request("LOCKED")
If LOCKED="1" or LOCKED="0"  Then
	findKey=findKey&" and E.LOCKED="&LOCKED
End If



IS_EXAM	 	= Request("IS_EXAM")
If IS_EXAM="1" or IS_EXAM="0"  Then
	If IS_EXAM="0" Then
		findKey=findKey&" and USER_ID not in(select USER_ID from T_TEST_RESULT)"
	End If
	If IS_EXAM="1" Then
		findKey=findKey&" and USER_ID in(select USER_ID from T_TEST_RESULT)"
	End If
End If



SCHOOL_ID	 	= Request("SCHOOL_ID")
If SCHOOL_ID<>"0" and is_Id(SCHOOL_ID)  Then
	findKey=findKey&" and E.SCHOOL_ID="&SCHOOL_ID
End If

EXAMER_INFO_TIME = Request("EXAMER_INFO_TIME")
If is_Id(EXAMER_INFO_TIME) and EXAMER_INFO_TIME >0 Then
	findKey=findKey&" and INFO_TIME>="&ToUnixTime(EXAMER_INFO_TIME&" 01-01",+8)&" and INFO_TIME<="&ToUnixTime(cStr(cInt(EXAMER_INFO_TIME)+1)&" 01-01",+8)
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
        <label for="SCHOOL_ID">所在学校:</label>
            <select id="SCHOOL_ID" name="SCHOOL_ID" onChange="$('#inputString').val($(this).find('option:selected').text())">
                <option value="0">选择学校</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,SCHOOL from T_SCHOOL where CHECKED=1  order by SCHOOL asc"
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
      <label for="EMAIL">e-mail:</label>
        <input type="text" id="EMAIL" name="EMAIL" value="<%=EMAIL%>" /> 
        <label for="IS_EXAM">报考:</label>
        <select id="IS_EXAM" name="IS_EXAM">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_EXAM="0" Then Response.write "selected='selected'"%> >否</option>
            <option value="1"<%If IS_EXAM="1" Then Response.write "selected='selected'"%>>是</option>
        </select><br>
        <label for="EXAMER_INFO_TIME">注册年份:</label>
        <select id="EXAMER_INFO_TIME" name="EXAMER_INFO_TIME">
        	<option value="-1">未选择</option>
            <%for i=2013 to year(date)
			 %>
				 <option value="<%=i%>"<%If i=Cint(EXAMER_INFO_TIME) Then Response.write "selected='selected'"%> ><%=i%></option>
			<%
			 next
			%>
           
        </select>
        
            <br><br>
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
          反选 <img src="../Images/run.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="statusPost('0','multi','0');">审核</a> <img src="../Images/002.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('0','multi','1');">反审核</a>  <img src="../Images/key.gif" width="16" height="14" /><a href="javascript:vodi(0)" onClick="statusPost('5','multi','1');">重置密码</a> <img src="../Images/locked.gif" width="14" height="14" /><a href="javascript:void(0)" onClick="statusPost('1','multi','0');">锁定</a> <img src="../Images/unlocked.gif" width="16" height="14" /><a href="javascript:void(0)" onClick="statusPost('1','multi','1');">解锁</a>  <img src="../Images/seek.gif" width="16" height="16" /><a href="javascript:void(0)"   onclick="find_examinee();">查看资料</a> <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="statusPost('4','multi','-1');">彻底删除考生</a></font> </td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"  action="../Html/examinee_detail.asp"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
            <th width="3%" height="26">选择</th>
            <th width="4%" height="26">序号</th>
            <th width="9%" height="26">姓名</th>
            <th width="14%" height="26">身份证号</th>
            <th width="5%" height="26">性别</th>
            <th width="13%" height="26">所在学校</th>
            <th width="13%" height="26">电话</th>
            <th width="12%" height="26">手机</th>
            <th width="13%" height="26">Email</th>
            <th width="7%" height="26">审核</th>
            <th width="7%" height="26">锁定</th>
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
	.Sql="E.USER_ID,E.SCHOOL_ID,S.SCHOOL,E.GENDER,E.REAL_NAME,E.AGE,E.EDUCATION,E.TEL,E.MOBILE,E.EMAIL,E.LAST_LOGIN_TIME,E.CHECKED,E.LOCKED,E.IS_MAKEUP$T_EXAMINEE AS E Left Join T_SCHOOL AS S ON E.SCHOOL_ID=S.ID$1=1 "&findKey&"$E.USER_ID desc$E.USER_ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
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
            <td height="18" ><%=iRs(4,i)%></td>
            <td height="18" ><%=iRs(0,i)%></td>
            <td height="18" ><%=getGender(iRs(3,i))%></td>
            <td height="18" ><%=iRs(2,i)%></td>
            <td height="18" ><%=iRs(7,i)%></td>
            <td height="18" ><%=iRs(8,i)%></td>
            <td height="18" ><%=iRs(9,i)%></td>
            <td height="18" ><%=getChecked(iRs(11,i))%></td>
            <td height="18" ><%=getTrue(iRs(12,i))%></td>
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