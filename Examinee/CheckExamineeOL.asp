<%
dim ComeUrl,cUrl,AdminName

EXAMINEE_USER_ID=request.Cookies("login")("EXAMINEE_USER_ID")
SYS_USER_ID		=request.Cookies("login")("SYS_USER_ID")
USER_NAME		=request.Cookies("login")("USER_NAME")

test_title=request.Cookies("test")("test_title")
test_id=request.Cookies("test")("test_id")

if EXAMINEE_USER_ID="" then
	call CloseConn()
	response.redirect "/examinee_login.html"
	response.End()
end if



%>
