<%
dim ComeUrl,cUrl,AdminName

PURVIEW		= request.Cookies("login")("PURVIEW")
USER_NAME	= request.Cookies("login")("USER_NAME")
SYS_USER_ID	= request.Cookies("login")("SYS_USER_ID")

test_title	= request.Cookies("test")("test_title")
test_id		= request.Cookies("test")("test_id")
test_cate	= request.Cookies("test")("test_cate")

sys_exam_title	= request.Cookies("test")("sys_exam_title")
sys_exam_id		= request.Cookies("test")("sys_exam_id")
'response.write "ddd"&sys_exam_title&sys_exam_id
if PURVIEW<>"3" or PURVIEW="" or not is_Id(sys_exam_id) then
	call CloseConn()
	response.redirect "/login.html"
	response.End()
end if


%>
