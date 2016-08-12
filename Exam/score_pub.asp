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
				  }
				  else{
					 $(this).removeClass('checkBgColor');
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
function statusPost(status){
	
		if(status==0){
			if(!confirm("确认发布数据吗? ")) return false;	
		}
		if(status==1){
			if(!confirm("确认发取消布数据吗? ")) return false;	
		}
		

		$.post(	"../Public/score_pubStatus.asp?status="+status, 
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



</script>
<style>
.score{
	width:50px;}

</style>
</head>

<body  >
<!--#include file="test_id_check.asp"-->
<%

REAL_NAME	 	= Request("REAL_NAME")
If REAL_NAME<>"" and not sqlInjection(REAL_NAME) Then
	findKey=findKey&" and REAL_NAME ='"&REAL_NAME&"'"
End If
USER_ID	 	= Request("USER_ID")
If USER_ID<>"" and not sqlInjection(USER_ID) Then
	findKey=findKey&" and USER_ID ='"&USER_ID&"'"
End If

TEL	 	= Request("TEL")
If TEL<>"" and not sqlInjection(TEL) Then
	findKey=findKey&" and TEL='"&TEL&"'"
End If

MOBILE	 	= Request("MOBILE")
If MOBILE<>"" and not sqlInjection(MOBILE) Then
	findKey=findKey&" and MOBILE='"&MOBILE&"'"
End If

EMAIL	 	= Request("EMAIL")
If EMAIL<>"" and not sqlInjection(EMAIL) Then
	findKey=findKey&" and EMAIL ='"&EMAIL&"'"
End If


IS_MAKEUP	 	= Request("IS_MAKEUP")
If IS_MAKEUP="1" or IS_MAKEUP="0"  Then
	findKey=findKey&" and IS_MAKEUP ="&IS_MAKEUP
End If


SCHOOL_ID	 	= Request("SCHOOL_ID")
If SCHOOL_ID<>"0" and is_Id(SCHOOL_ID)  Then
	findKey=findKey&" and SCHOOL_ID="&SCHOOL_ID
End If
If sys_exam_id<>"0" and is_Id(sys_exam_id)  Then
	findKey=findKey&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&")"
End If
TEST_NO	 	= trim(Request("TEST_NO"))
If TEST_NO<>"" and is_Id(TEST_NO)  Then
	findKey=findKey&" and TEST_NO="&TEST_NO&" "
End If

IS_PASS	 	= Request("IS_PASS")
If IS_PASS="1" or IS_PASS="0"  Then
	findKey=findKey&" and IS_PASS ="&IS_PASS
End If

TOPIC_ID	 	= Request("TOPIC_ID")
If TOPIC_ID<>"-1" and is_Id(TOPIC_ID)  Then
	findKey=findKey&" and USER_ID in(select USER_ID from T_TOPIC_RESULT where TOPIC_ID="&TOPIC_ID&" and test_id="&test_id&")"
	topicFind=topicFind&" and TOPIC_ID="&TOPIC_ID
End If
%>
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>成绩列表</span></td>
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
        <label for="IS_PASS">及格:</label>
        <select id="IS_PASS" name="IS_PASS">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_PASS="0" Then Response.write "selected='selected'"%>>未及格</option>
            <option value="1"<%If IS_PASS="1" Then Response.write "selected='selected'"%>>及格</option>
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
         <label for="TEST_NO">准考证号:</label>
        <input type="text" id="TEST_NO" name="TEST_NO" value="<%=TEST_NO%>" />
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
        
       
         <label for="IS_MAKEUP">补考:</label>
        <select id="IS_MAKEUP" name="IS_MAKEUP">
        	<option value="-1">未选择</option>
            <option value="0"<%If IS_MAKEUP="0" Then Response.write "selected='selected'"%>>否</option>
            <option value="1"<%If IS_MAKEUP="1" Then Response.write "selected='selected'"%>>是</option>
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
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><img src="../Images/excel.gif" width="14" height="14" /><font><a href="javascript:void(0)"  onclick="window.location.href='examinee_score_export_excel.asp'+window.location.search;">导出Excel</a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
          	<th width="4%" height="26">序号</th>
            <th width="13%" height="26">身份证号</th>
            <th width="11%" height="26">准考证号</th>
            <th width="11%" height="26">证书编号</th>
            <th width="7%" height="26">姓名</th>
            <th width="54%" height="26">科目</th>
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
	.Sql="[USER_ID],[TEST_YEAR],[TEST_NO],[CRS_NO],[IS_PASS],[IS_MAKEUP],[INFO_TIME],[REAL_NAME],[TITLE],[TEST_ID]$V_TEST_RESULT$1=1 and TEST_ID="&TEST_ID&" "&findKey&"$TEST_NO asc$USER_ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>  
          <tr>
            <td height="18"  colspan="6">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	%>    
          <tr>
          	<td height="18" ><%=i+1%></td>
            <td height="18" ><%=iRs(0,i)%></td>
            <td height="18" ><%=iRs(2,i)%></td>
            <td height="18" ><%=iRs(3,i)%></td>
            <td height="18" ><%=iRs(7,i)%></td>
            <td height="18" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
            	<th width="47%">科目</th>
                <th width="18%">分数</th>
                <th width="10%">及格</th>
                <th width="12%">是否补考</th>
                <th width="13%">是否公布</th>
            </tr>
            <%set rs_topic=server.createobject("ADODB.Recordset")
			 sql="SELECT [USER_ID],[TEST_ID],[TOPIC_ID],[RESULT],[IS_PASS],[IS_MAKEUP],[IS_SCORE],[TITLE],[TNAME],[IS_PRATICE] FROM [V_SCORE] where USER_ID='"&iRs(0,i)&"' and TEST_ID="&test_id&" "&topicFind
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
						if rs_topic("IS_SCORE")=1 then 
						isPassStr=getTrue(rs_topic("IS_PASS"))
					else
						isPassStr=""
					end if
					
						Response.write "<tr><td>"&rs_topic("TNAME")&"</td><td>"&exam_score&"</td><td>"&isPassStr&"</td><td>"&getTrue(rs_topic("IS_MAKEUP"))&"</td><td>"&getTrue(rs_topic("IS_SCORE"))&"</td></tr>"
				   	rs_topic.movenext
				   	loop
			  	End If
			   rs_topic.close
			   set rs_topic=nothing
			   %>
            
            </table></td>
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