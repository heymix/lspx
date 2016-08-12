<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

opValue = Request("opValue")
operateType = Request("operateType")
opID = Request("sID")

If PURVIEW="" or  test_id="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

set rs_main=conn.execute("select USER_ID from T_TEST_RESULT where IS_MAKEUP=1 and TEST_ID="&test_id)
do while not(rs_main.eof or err)
	set rs=conn.execute("select TOPIC_ID from T_TEST_TOPIC_RELATION where TOPIC_ID not in(select TOPIC_ID from T_TOPIC_RESULT where USER_ID='"&rs_main("USER_ID")&"') and TEST_ID="&test_id)
	do while not(rs.eof or err)
		conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_PASS,IS_MAKEUP,IS_SCORE,IS_OLD_SCORE)values('"&rs_main("USER_ID")&"',"&test_id&","&rs("TOPIC_ID")&",1,1,0,1)")
		'response.write rs_main("USER_ID")&"|||" &rs("TOPIC_ID")&"<br>"
		rs.movenext
	loop
	rs.close
	set rs=nothing
rs_main.movenext
loop
rs_main.close
set rs_main=nothing



call CloseConn()
%>
