<%

'session.Abandon()
response.Cookies("login")=""
response.write "<script>window.location.href='../examinee_login.html'</script>"

%>