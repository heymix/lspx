<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
PURVIEW		= request.Cookies("login")("PURVIEW")
EXAMINEE_USER_ID=request.Cookies("login")("EXAMINEE_USER_ID")
If not(PURVIEW<>"" or EXAMINEE_USER_ID<>"") Then
	Response.End()
End If
test_id=request("test_id")
test_title=request("test_title")
'Response.write "0|"&test_id&test_title
'Response.End()
If  test_id="" or test_title="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	response.End()
End If
test_Arr=split(test_id,"|")

response.Cookies("test")("test_id")	=test_Arr(0)
response.Cookies("test")("test_cate")=test_Arr(1)
response.Cookies("test")("test_title")=test_title
Response.write "1|选择成功"

%>