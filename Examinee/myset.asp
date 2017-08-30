<!--#include file="../inc/conn.asp"-->
<!--#include file="../Examinee/CheckExamineeOL.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()

'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>考试信息管理</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<link href="../Css/news.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
    <link rel="stylesheet" href="../jquery/base/jquery.ui.all.css">
	<script src="../jquery/ui/jquery.timePicker.js"></script>
	<script src="../jquery/ui/jquery.ui.core.min.js"></script>
	<script src="../jquery/ui/jquery.ui.widget.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker.min.js"></script>
	<script src="../jquery/ui/jquery.ui.datepicker-zh-CN.min.js"></script>
    
    
    <script src="../Jquery/jquery.bgiframe.min.js" type="text/javascript" charset='utf-8'></script>
<script src="../Jquery/loading-min.js" type="text/javascript"  charset="UTF-8"></script>
 <script type="text/javascript">

 $(document).ready(function(e) {
	  loading=new ol.loading({id:"tablePanel"});
	$(function() {
		$( "#GRADUATE_DATE" ).datepicker({
            changeYear: true,
			dateFormat: 'yy年mm月'
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

			if($("#REAL_NAME").val()==""){
				alert("姓名不能为空！");
				return false;
			}
			if($("#GRADUATE_SCHOOL").val()==""){
				alert("毕业学校不能为空！");
				return false;
			}
			if($("#GRADUATE_DATE").val()==""){
				alert("毕业日期不能为空！");
				return false;
			}
			if($("#PRO").val()=="0"){
				alert("专业不能为空！");
				return false;
			}
			
			if($("#EDUCATION").val()=="0"){
				alert("请选择学历！");
				return false;
			}
			if($("#SCHOOL_ID").val()=="-1"){
				alert("选择所在学校！");
				return false;
			}
			
			if($("#PHOTO").val()==""){
				alert("没有上传照片！");
				return false;
			}
			if($("#TEL").val()=="" || $("#MOBILE").val()==""){
				alert("电话和手机必填一项！");
				return false;
			}
					
			//loading.show();
			$.post(	"../Public/examineeEdit.asp", 
					$("#editForm").serialize(), 
					function(data,st){
							
							var resultArr=data.split("|");
							if(resultArr[0]=="1"){
								
								//alert(resultArr[1]);
								//window.location.reload();
								window.location.href="exam_post.asp"
								resetForm();
							}
							else{
								alert(resultArr[1]);
								
							}	
					});
		
	}

</script>
</head>
<script>
function backfn(fname){
	document.getElementById("PHOTO").value=fname;
	document.getElementById('PHOTO_IMG').src="../UploadFiles/userPhoto/"+fname;
}
</script>
<body  >
<%
If EXAMINEE_USER_ID<>"" Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT * from T_EXAMINEE where USER_ID='"&EXAMINEE_USER_ID&"'"
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		Response.write errStr ("错误：没有查到需要修改的记录","-1")
		appEnd()
	Else
		GENDER=rs("GENDER")
		REAL_NAME=rs("REAL_NAME")
		AGE=rs("AGE")
		EDUCATION=rs("EDUCATION")
		SCHOOL_ID=rs("SCHOOL_ID")
		TEL=rs("TEL")
		MOBILE=rs("MOBILE")
		MOBILE_BAK=rs("MOBILE_BAK")
		EMAIL=rs("EMAIL")
		GRADUATE_SCHOOL=rs("GRADUATE_SCHOOL")
		GRADUATE_DATE=Format_time(FromUnixTime(rs("GRADUATE_DATE"),+8),7)
		PRO=rs("PRO")
		PHOTO=rs("PHOTO")
		INFO_TIME=rs("INFO_TIME")
		CHECKED=rs("CHECKED")
		LOCKED=rs("LOCKED")
		PHOTO=rs("PHOTO")
		If PHOTO="" or isNULL(PHOTO) Then
			PHOTO_show="../Images/nonepic.jpg"
		Else
			PHOTO_show="../UpLoadFiles/userPhoto/"&PHOTO
		End If
	End If
	rs.close
	set rs=nothing
End If
%>
<%If CHECKED=1 Then%>
<script>
 $(document).ready(function(e) {
	  	$("input").attr("disabled", "disabled"); 
		$("select").attr("disabled", "disabled"); 
	});
</script>
<%End If%>
<table id="tablePanel"  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>资类信息修改</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm">
        <input id="operateType"  name="operateType" value="1" type="hidden">
            <div class="editFormLeft"> <label for="REAL_NAME">*姓名:</label>
            <input type="text" id="REAL_NAME" name="REAL_NAME" value="<%=REAL_NAME%>" />
            <br>
            <label for="GENDER">性别:</label>
           <input type="text" id="GENDER" name="GENDER" value="<%=getGender(GENDER)%>" disabled="disabled" />
            <br>
            <label for="AGE">年龄:</label>
            <input type="text" id="AGE" name="AGE" value="<%=AGE%>" disabled="disabled" />
            <br>
            <label for="GRADUATE_SCHOOL">*毕业学校:</label>
            <input type="text" id="GRADUATE_SCHOOL" name="GRADUATE_SCHOOL" value="<%=GRADUATE_SCHOOL%>" />
            <br>
            <label for="GRADUATE_DATE">*毕业时间:</label>
            <input type="text" id="GRADUATE_DATE" name="GRADUATE_DATE" value="<%=GRADUATE_DATE%>" readonly="readonly" />
            </div><div class="editFormRight">
            <div class="editFormPhoto"><img id="PHOTO_IMG" src="<%=PHOTO_show%>" height="151" width="108"></div><br>
            <label for="UploadFiles">*上传照片:<br>照片尺寸1寸<br> 413*295像素 <br>小于2MB<br>照片背景蓝色或者浅色</label>
            <input type="hidden" id="PHOTO" name="PHOTO" value="<%=PHOTO%>">
           <%If CHECKED<>1 Then%>
            <iframe style="margin-top:10px; margin-left:0px;" ID="UploadFiles" name="upload" src="upload.asp" frameborder=0 scrolling=no width="430" height="23"></iframe>
            <%End If%>
            </div>
            <br>
            <label for="PRO">*专业:</label>
            <input type="text" id="PRO" name="PRO" value="<%=PRO%>" />
            <br>
            <label for="EDUCATION">*最后学历:</label>
            <select id="EDUCATION" name="EDUCATION">
                <option value="0">*选择学历</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,KEY_ID,KEY_VALUE from T_DATA_KEY where KEY_ID=3"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=EDUCATION Then Response.write "selected=""selected"""  %>><%=rs("KEY_VALUE")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select>
            <br>
            <label for="SCHOOL_ID">*工作单位:</label>
            <select id="SCHOOL_ID" name="SCHOOL_ID" style=" width:263px;">
                <option value="-1">*选择工作单位</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,SCHOOL from T_SCHOOL where CHECKED=1 order by SCHOOL asc"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=SCHOOL_ID Then Response.write "selected=""selected"""  %>><%=rs("SCHOOL")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
              </select>
            <label id="SCHOOL_ID_DESC" for="SCHOOL_ID"> 若列表内没有您的所在单位，请与考试中心联系（电话：0411-82158139/0411-82158550）</label>
            <br>
            <label for="USER_ID">身份证号:</label>
            <input type="text" id="USER_ID" name="USER_ID" value="<%=EXAMINEE_USER_ID%>" disabled="disabled" />
            <br>
            <label for="TEL">电话:</label>
            <input type="text" id="TEL" name="TEL" value="<%=TEL%>" /><label id="TEL_DESC" for="SCHOOL_ID">请在您的电话号码前加区号，如：010-65906903 </label>
            <br>
            <label for="MOBILE">手机:</label>
            <input type="text" id="MOBILE" name="MOBILE" value="<%=MOBILE%>" /><label id="MOBILE_DESC" for="MOBILE"></label><br />
             <label for="MOBILE_BAK">备用手机:</label>
            <input type="text" id="MOBILE_BAK" name="MOBILE_BAK" value="<%=MOBILE_BAK%>" /><label id="MOBILE_DESC" for="MOBILE_BAK">备用手机号 仅做联系备用</label>

            <br>
            <label for="EMAIL">Email:</label>
            <input id="EMAIL" name="EMAIL" type="text" class="input" value="<%=EMAIL%>"><br />   
            <label for="CHECKED">审核:</label>  <input type="text" id="CHECKED" value="<%=getChecked(CHECKED)%>" disabled="disabled"> 
<br/>
<br/><br/>

   
     <%If CHECKED<>1 Then%>
            <div  class="suggestionsMain" style="text-align:right">
 					<img src="../Images/005.gif" width="14" height="14" />[<a id="submitLink" href="javascript:void(0)" onClick="submitForm();">确认下一步</a>] <img src="../Images/002.gif" width="14" height="14"/>[<a href="javascript:void(0)" onClick="resetForm();">重置</a>]
              </div>
   <%End If%>
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
<!--本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒-->

</body>
</html>

<%
appEnd()

%>