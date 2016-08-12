<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()
id=request("id")
operateType=0
CHECKED=1
IS_TOP=0
If id<>"" and isNumeric(id) Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT N.ID,CATEGORY,D.KEY_VALUE,TITLE,CONTENT,INFO_TIME,CHECKED,IS_TOP,IS_TOP_TIME,DOWN_LINK FROM T_DOWNLOAD AS N Left Join T_DATA_KEY AS D ON N.CATEGORY=D.ID where N.ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		errStr "错误：没有查到需要修改的记录","-1"
		appEnd()
	Else
		operateType=1
		CATEGORY=rs("CATEGORY")
		TITLE=rs("TITLE")
		CONTENT=rs("CONTENT")
		DOWN_LINK=rs("DOWN_LINK")
		CHECKED=rs("CHECKED")
		IS_TOP=rs("IS_TOP")
	End If
	
	
	rs.close
	set rs=nothing
End If
'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理员基础信息</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<link href="../Css/news.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>


<script type="text/javascript">

 $(document).ready(function(e) {

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

function backfn(fname){
	$("#score_files").val(fname);
	$.post(	"../Public/score_upEdit.asp", 
				{subBox:fname}, 
				function(data,st){
					if (st=="success") $("#score_data").html(data)
					
				});
		
}


function submitForm(){


		$.post(	"../Public/score_up_import.asp", 
				$("#editForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1])
						}
						else{
							alert(resultArr[1]);
						
						}	
				});
	
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
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>导入分数</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm" >
    <input type="hidden" id="score_files" name="score_files" value="" /><br />
<iframe style="top:2px; margin-top:15px; margin-left:80px;" ID="UploadFiles" name="upload" src="upload_score.asp" frameborder=0 scrolling=no width="430" height="23"></iframe><br><br>
	<div  class="suggestionsMain" style="text-align:right">	选择文件 -上传 - 检查数据，之后 提交并完成导入。</div>

<br/>
<br/>
 <div  class="suggestionsMain" style="text-align:right">
 			 <img src="../Images/005.gif" width="14" height="14" />[<a id="submitLink" href="javascript:void(0)" onClick="submitForm();">提交并导入分数</a>] <img src="../Images/002.gif" width="14" height="14"/>[<a id="resetLink" href="javascript:void(0)">重置</a>]
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
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">操作面板&gt;&gt;导入分数检测</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn">&nbsp;</td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle" id="score_data">无数据！</td>
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


<!--本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒-->

</body>
</html>

<%
appEnd()

%>