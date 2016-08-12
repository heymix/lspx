<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

If PURVIEW=""or test_id="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If
 '查询当前考试类别
test_cate=0
set rs=conn.execute("select TEST_CATE from T_TEST where ID="&test_id)
	test_cate=rs("TEST_CATE")
rs.close
set rs=nothing
If test_cate=0 Then
	Response.write "0|错误考试类别出错"
End If
'查询当前考试的总科目
set rs=conn.execute("select count(ID) from T_TEST_TOPIC where TEST_CATE="&test_cate)
	topic_num=rs(0)
rs.close
set rs=nothing
'查询考试过的科目 并更新
set rs=conn.execute("select USER_ID from V_TEST_RESULT where TEST_ID="&test_id)
do while not (rs.eof or err)
	i=0
	set rs_score=conn.execute("select count(USER_ID) from V_SCORE where IS_PASS=1 and USER_ID='"&rs("USER_ID")&"' and TEST_CATE="&TEST_CATE)
		If Cint(rs_score(0))=Cint(topic_num) Then
			conn.execute("update T_TEST_RESULT set IS_PASS=1 where USER_ID='"&rs("USER_ID")&"' and  TEST_ID="&test_id)
		End If
	rs_score.close
	set rs_score=nothing
rs.movenext
loop
rs.close
set rs=nothing
Response.write "1|生成结束"
call CloseConn()

%>
