<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../Examinee/CheckExamineeOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("checkType")
oStatus = Request("status")
operateCon=""
'Response.write "0|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If subBox="" or oStatus=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else


		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="4" Then
					conn.execute("delete  from T_TEST_RESULT  where USER_ID='"&EXAMINEE_USER_ID&"' and TEST_ID="&udID(i))
					conn.execute("delete  from T_TOPIC_RESULT where USER_ID='"&EXAMINEE_USER_ID&"' and TEST_ID="&udID(i))
					operateCon="删除了考试和科目"
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "exam.asp",operateCon&"："&subBox
			Response.write "1|操作成功！"
			conn.CommitTrans 
		Else
			Response.write "0|操作失败！"
			conn.RollbackTrans
		End If
	
End If
call CloseConn()
%>
