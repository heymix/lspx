<!--#include file="../Inc/UpLoad_Class.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script src="../Inc/jquery-1.7.1.min.js" type="text/javascript" language="javascript"></script>
    <script src="../Inc/jquery.filestyle.js" type="text/javascript" language="javascript"></script>

    <script language="javascript">
		$(document).ready(function() {
		$("input.file1").filestyle({image: "../../Images/Choose_file_zh.gif",imageheight : 22,imagewidth : 82,width : 240});  
		});
	</script>
<SCRIPT language=javascript>
function check() 
{
	var strFileName=document.upfrm.file1.value;
	if (strFileName=="")
	{
    	alert("请选择要上传的文件");
		document.upfrm.file1.focus();
    	return false;
  	}
	$('#upfrm').submit();
}
</SCRIPT>
<style type="text/css">
#upfrm input {
	border: 1px solid #999;
	vertical-align:middle;
}
body {
	background-color: #F3FFE3;
}
#upfrm img{
	vertical-align:middle;
}
#upfrm a{
	vertical-align:middle;
}
#upfrm span{
	line-height:20px;
	height:20px;
	vertical-align:middle;
	cursor:pointer;
}
</style>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" topmargin="0">
<%
if lcase(request.ServerVariables("REQUEST_METHOD"))="post" then
dim upload
set upload = new AnUpLoad
upload.Exe = "jpg|bmp|jpeg|gif|png|rar|zip|doc|xls"
upload.MaxSize = 2 * 1024 * 1024 '2M
upload.GetData()
if upload.ErrorID>0 then 
	response.Write upload.Description
else
	dim file,savpath
	savepath = "../UploadFiles/files"
	set file = upload.files("file1")
	if file.isfile then
		result = file.saveToFile(savepath,0,true)
		if result then
			msg = file.filename
		else
			msg = file.Exception
		end if
	end if
end if
set upload = nothing
%>

<a href="upload.asp">重新上传</a>
<script>window.parent.backfn.apply(this,["<%=msg%>"]);</script>
<%else%>

<form action="upload.asp" method="post" enctype="multipart/form-data" target="upload" id="upfrm" name="upfrm"> 
<input type="file" name="file1"  class="file1" style="border: 1px solid #000;"/>  
<a href="javascript:void(0)" style="position:absolute; top:1px; left:340px;" onClick="check();"><img src="../Images/is_top.gif" width="14" height="14" /> <span>上 传</span></a><br />
</form>
<%end if%>
</body>
</html>