<%

'session.Abandon()
response.Cookies("login")=""
response.write "<scriprt>window.location.href='../login.html'</scriprt>"

%>