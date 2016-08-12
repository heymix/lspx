<%
dim ComeUrl,cUrl,AdminName

PURVIEW		= request.Cookies("login")("PURVIEW")
USER_NAME	= request.Cookies("login")("USER_NAME")
SYS_USER_ID	= request.Cookies("login")("SYS_USER_ID")

test_title	= request.Cookies("test")("test_title")
test_id		= request.Cookies("test")("test_id")
test_cate	= request.Cookies("test")("test_cate")

sys_school_title	= request.Cookies("test")("sys_school_title")
sys_school_id		= request.Cookies("test")("sys_school_id")

if PURVIEW<>"2" or PURVIEW="" then
	call CloseConn()
	response.redirect "/login.html"
	response.End()
end if


%>
