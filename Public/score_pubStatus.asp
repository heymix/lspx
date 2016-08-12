<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除

oStatus = Request("status")
operateCon=""
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or oStatus="" or test_id=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" Then
		conn.begintrans
	
		If oStatus="0" Then
			conn.execute("update T_TOPIC_RESULT set IS_SCORE=1 where TEST_ID="&test_id)
			conn.execute("update T_TEST_RESULT set IS_PASS=1 where USER_ID in (select USER_ID from T_TOPIC_RESULT Left Join T_TEST_TOPIC ON T_TOPIC_RESULT.TOPIC_ID=T_TEST_TOPIC.ID  where IS_PASS=1 and T_TEST_TOPIC.TEST_CATE=12  and TEST_ID="&test_id&"  GROUP BY USER_ID HAVING COUNT(*)=5) and  TEST_ID="&test_id)
			operateCon="发布成绩"
		ElseIf oStatus="1" Then 
			conn.execute("update T_TOPIC_RESULT set IS_SCORE=0 where TEST_ID="&test_id)
			operateCon="取消发布成绩"
		End If

	
		if Err.Number = 0 Then
			operateLog "score_pub.asp",operateCon&"："&subBox
			Response.write "1|操作成功！"
			conn.CommitTrans 
		Else
			Response.write "0|操作失败！"
			conn.RollbackTrans
		End If
	Else
		Response.write "0|操作失败！没有权限"
	End If
End If
call CloseConn()
%>
