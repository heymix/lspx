<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
operateCon=""
'Response.write "0|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" or checkType="" or oStatus=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="0" Then
					If checkType="0" Then 
						conn.execute("update T_EXA set CHECKED=1 where  ID="&udID(i))
						operateCon="审核"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXA set CHECKED=0 where ID="&udID(i))
						operateCon="反审核"
					End If
				ElseIf oStatus="1" Then
					If checkType="0" Then 
						conn.execute("update T_EXA set IS_Run=1 where ID="&udID(i))
						operateCon="提交"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXA set IS_Run=0 where ID="&udID(i))
						operateCon="撤回"
					End If
				ElseIf oStatus="4" Then
					'conn.execute("delete  from T_EXA  where ID="&udID(i))
					'conn.execute("delete  from T_EXA_SCHOOL_RELATION where EXA_ID="&udID(i))
					'conn.execute("delete  from T_ADMIN where EXA_ID="&udID(i))
					operateCon="删除了考点下的学校和考点管理员 考点ID"
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
	Else
		Response.write "0|操作失败！没有权限"
	End If
End If
call CloseConn()
%>
