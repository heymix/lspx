<%

'session.Abandon()
response.Cookies("login")=""
response.Cookies("test")=""
response.write "<script>window.location.href='../login.html'</script>"

%>