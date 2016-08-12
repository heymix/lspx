<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/CheckUserOL.asp"-->
<%
sys_exam_id=request("sys_exam_id")
sys_exam_title=request("sys_exam_title")
'Response.write "0|"&test_id&test_title
'Response.End()
If PURVIEW="" or  sys_exam_id="" or sys_exam_title="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	response.End()
End If
test_Arr=split(test_id,"|")

response.Cookies("test")("sys_exam_id")	=sys_exam_id
response.Cookies("test")("sys_exam_title")=sys_exam_title
Response.write "1|选择成功"

%>