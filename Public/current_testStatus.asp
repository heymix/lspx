<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../Manager/CheckManager.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
oStatus = Request("status")
operateCon=""
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" or TEST_ID="" or oStatus=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="5" Then 
					conn.execute("DELETE T_TEST_EXA_RELATION where TEST_ID="&TEST_ID&" AND EXA_ID="&udID(i))
					operateCon="从考试("&test_id&":"&test_title&")中移除考点"
				End If
				If oStatus="6" Then 
					theSame=conn.execute("select count(*) from T_TEST_EXA_RELATION where TEST_ID="&TEST_ID&" AND EXA_ID="&udID(i))(0)			
					If theSame=0 Then
						conn.execute("insert into T_TEST_EXA_RELATION (TEST_ID,EXA_ID)values("&TEST_ID&","&udID(i)&")")
					End If
					operateCon="添加考点到考试("&test_id&":"&test_title&")"
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "current_test.asp",operateCon&"："&subBox
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
