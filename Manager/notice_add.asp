<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()
id=request("id")
operateType=0
CHECKED=1
IS_TOP=0
IS_READ=0
If id<>"" and isNumeric(id) Then
	set rs=server.createobject("adodb.recordset")
	sql="SELECT N.ID,CATEGORY,D.KEY_VALUE,TITLE,NOTICE,INFO_TIME,CHECKED,IS_TOP,IS_TOP_TIME,IS_READ FROM T_NOTICE AS N Left Join T_DATA_KEY AS D ON N.CATEGORY=D.ID where N.ID="&ID
	'response.write sql
	rs.open sql,conn,1,1
	If rs.eof or rs.bof or err Then
		errStr "错误：没有查到需要修改的记录","-1"
		appEnd()
	Else
		operateType=1
		CATEGORY=rs("CATEGORY")
		TITLE=rs("TITLE")
		content=rs("NOTICE")
		CHECKED=rs("CHECKED")
		IS_TOP=rs("IS_TOP")
		IS_READ=rs("IS_READ")
	End If
	
	
	rs.close
	set rs=nothing
End If
'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>公告信息管理</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<link href="../Css/news.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>

	<link rel="stylesheet" href="../editor/themes/default/default.css" />
	<link rel="stylesheet" href="../editor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../editor/kindeditor.js"></script>
	<script charset="utf-8" src="../editor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../editor/plugins/code/prettify.js"></script>
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





		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="content1"]', {
				cssPath : '../editor/plugins/code/prettify.css',
				uploadJson : '../editor/asp/upload_json.asp',
				fileManagerJson : '../editor/asp/file_manager_json.asp',
				allowFileManager : true,
				afterCreate : function() {
					
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=addNEWS]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=addNEWS]')[0].submit();
					});
				}
			});
			prettyPrint();

				K('#resetLink').click(function(e) {
					editor.html('');
					$('#editForm')[0].reset();
				});
			
				K('#submitLink').click(function(e) {
					
					$("#content").val(editor.html());
						if($("#TITLE").val()==""){
							alert("标题不能为空！");
							return false;
						}	
						if($("#CATEGORY").val()=="0"){
							alert("没有选择类别！");
							return false;
						}
						if($("#content").val()==""){
							alert("内容不能为空！");
							return false;
						}						
						$.post(	"../Public/noticeEdit.asp", 
								$("#editForm").serialize(), 
								function(data,st){
										
										var resultArr=data.split("|");
										if(resultArr[0]=="1"){
											
											if(!confirm("操作成功!\n\n[确定] 继续操作，[取消] 查看列表...")){
												window.location.href='notice.asp';
												return false;
												} 
											editor.html('');
											$('#editForm')[0].reset();;
										}
										else{
											alert(resultArr[1]);
											
										}	
								});
				
				});
		});


</script>
</head>

<body  >
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/operatePanle.gif" width="16" height="16" /> <span class="tbTitle">操作面板>>添加通知公告</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location='notice.asp'"><img src="../Images/seek.gif" width="14" height="14" />查询</a></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"  class="tdStyle"><form method="post" id="editForm" name="editForm"  title="学校管理" >
        <input id="operateType" name="operateType" value="<%=operateType%>" type="hidden">
       <input id="ID" name="ID" value="<%=ID%>" type="hidden">
		<label for="TITLE">标题:</label>
<input type="text" id="TITLE" name="TITLE" value="<%=TITLE%>" /><br />
		<label for="CATEGORY">类别:</label> 		
        		<select id="CATEGORY" name="CATEGORY">
                <option value="0">选择类别</option>
                <%
					set rs=server.createobject("adodb.recordset")
					sql="select ID,KEY_ID,KEY_VALUE from T_DATA_KEY where KEY_ID=1"
					rs.open sql,conn,1,3
					do while not (rs.eof or err)
						%>
						<option value="<%=rs("ID")%>" <%If rs("ID")=CATEGORY Then Response.write "selected=""selected"""  %>><%=rs("KEY_VALUE")%></option>
						<%
					rs.movenext
					loop
					rs.close
					set rs=nothing
				%>
                	
                </select>         
<br/>
<label for="IS_READ">权限:</label> 		
        		<select id="IS_READ" name="IS_READ">

						<option value="0" <%If IS_READ=0 Then Response.write "selected=""selected"""  %>>所有人可读</option>
                        <option value="1" <%If IS_READ=1 Then Response.write "selected=""selected"""  %>>仅考生可读</option>
                        <option value="2" <%If IS_READ=2 Then Response.write "selected=""selected"""  %>>仅学校可读</option>
                        <option value="3" <%If IS_READ=3 Then Response.write "selected=""selected"""  %>>仅考点可读</option>
                        <option value="5" <%If IS_READ=5 Then Response.write "selected=""selected"""  %>>仅学校和考点</option>   	
                </select>         
<br/>
<br/>
<textarea name="content1" class="textarea_editor"><%=htmlspecialchars(content)%></textarea>
<input id="content" name="content" value="<%=htmlspecialchars(content)%>" type="hidden">

<br>
<label for="CHECKED">审核:</label>  <input type="checkbox" value="1" id="CHECKED" name="CHECKED" class="radioType" <%If CHECKED=1 Then Response.Write "checked" End If%>> 
<label for="IS_TOP">置顶:</label>  <input type="checkbox" value="1" id="IS_TOP" name="IS_TOP" class="radioType" <%If IS_TOP=1 Then Response.write "checked"  End If%>> 
<br/>
 <div  class="suggestionsMain" style="text-align:right">
 			 <img src="../Images/005.gif" width="14" height="14" />[<a id="submitLink" href="javascript:void(0)" onClick="submitForm();">提交数据</a>] <img src="../Images/002.gif" width="14" height="14"/>[<a id="resetLink" href="javascript:void(0)">重置</a>]
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
<!--本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒-->

</body>
</html>

<%
appEnd()

%>