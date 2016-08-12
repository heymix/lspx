<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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

function createNO(status){

		if(!confirm("确定生成及格人数吗? 时间比较长，请安确定键，耐心等待返回.... ")) return false;	

		$.post(	"../Public/creatPass.asp.asp", 
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
<BODY>
<!--#include file="test_id_check.asp"-->
<!--include file="exam_id_check.asp"-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="275" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">成绩统计</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"><table class="list_table" width="99%" border="0" align="center" cellpadding="0" cellspacing="1"  >
          <tr>
          	<th width="2%" height="26">序号</th>
            <th width="7%" height="26">学校名称</th>
            <th width="6%" height="26">培训人数</th>
            <th width="4%" height="26">补考人数</th>
            <th width="6%" height="26">缺考人数</th>
            <th width="9%" height="26">考试人数</th>
            <th width="7%" height="26">不及格人数</th>
            <th width="6%" height="26">及格人数</th>
            <th width="37%" height="26">及格情况</th>
            <th width="7%" height="26">及格率</th>
            </tr>
         

  <%
SCHOOL_ID	 	= sys_school_id
If SCHOOL_ID<>"0" and is_Id(SCHOOL_ID)  Then
	findKey=findKey&" and SCHOOL_ID="&SCHOOL_ID
End If
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
	.Sql="[SCHOOL],[SCHOOL_ID],[TEST_CATE]$[V_TEST_SCHOOL]$1=1 and TEST_ID="&test_id&" "&findKey&"$SCHOOL desc$SCHOOL_ID" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>  
          <tr>
            <td height="18"  colspan="11">[没有查到记录！]</div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	%>    
          <tr>
          <%
		  totalNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
		  ismakupNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where IS_MAKEUP=1 and SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
		  qkNum=conn.execute("select count(distinct(USER_ID)) from V_SCORE where IS_OLD_SCORE=0 and RESULT=0 and IS_PRATICE=0 and SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
		  
		  totalIsPassNum=conn.execute("select count(USER_ID) from V_TEST_RESULT where IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)

		  %>
          <td height="18" ><%=i+1%></td>
            <td height="18" ><%=iRs(0,i)%></td>
           <td height="18" ><%=totalNum-ismakupNum%></td>
            <td height="18" ><%=ismakupNum%></td>
            <td height="18" ><%=qkNum%></td>
            <td height="18" ><%=totalNum-qkNum%></td>
            <td height="18" ><%=totalNum-totalIsPassNum%></td>
            <td height="18" ><%=totalIsPassNum%></td>
            
            <td height="18" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
            	<th colspan="2"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=1")(0)%></th>
                <th colspan="2"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=2")(0)%></th>
                <th colspan="2"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=3")(0)%></th>
                <th colspan="2"><%=conn.execute("SELECT [TNAME] FROM [T_TEST_TOPIC] where ID=4")(0)%></th>
            </tr>
           <tr>
           	<th width="12.5%">人数</th>
            <th width="12.5%">及格率%</th>
            <th width="12.5%">人数</th>
            <th width="12.5%">及格率%</th>
            <th width="12.5%">人数</th>
            <th width="12.5%">及格率%</th>
            <th width="12.5%">人数</th>
            <th width="12.5%">及格率%</th>
            
           </tr>
            <tr>
           	<td width="12.5%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)%></td>
            <td width="12.5%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)%></td>
            <td width="12.5%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=2 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)%></td>
            <td width="12.5%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=3 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
            <td width="12.5%"><%=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)%></td>
            <td width="12.5%"><%isPassNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and IS_PASS=1 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								isAllNum=conn.execute("select count(USER_ID) from V_SCORE where TOPIC_ID=4 and  SCHOOL_ID="&iRs(1,i)&" and TEST_ID="&test_id)(0)
								If isAllNum>0 Then
									perNum=ROUND(isPassNum/isAllNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>

           </tr>
            
            </table></td>
            <td height="18" ><%	
								If totalNum>0 Then
									perNum=ROUND(totalIsPassNum/totalNum,2)
								Else
									perNum=0
								End If
								Response.write perNum*100&"%"%></td>
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
</BODY>
</HTML>
<%ors.ShowPage()%>
<%
iRs=NULL
ors=NULL
Set ors=NoThing

%>