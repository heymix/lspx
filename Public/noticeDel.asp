<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%


subBox	= Request("subBox")
operateCon=""
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then 
				conn.execute("delete  from T_NOTICE  where ID="&udID(i))
				operateCon="删除"
			End If
		next
		if Err.Number = 0 Then
			operateLog "topic.asp",operateCon&"："&subBox
			Response.write "1|操作成功！"
			conn.CommitTrans 
		Else
			Response.write "0|操作失败！"
			conn.RollbackTrans
		End If
	Else
		Response.write "0|操作失败！没有权限"
	End If

call CloseConn()
%>
