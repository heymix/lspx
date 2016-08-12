<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%


subBox	= Request("subBox")
checkType = Request("checkType")
operateCon=""
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" or checkType=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If checkType="0" Then 
					conn.execute("update T_NOTICE set IS_TOP=1 where ID="&udID(i))
					operateCon="置顶"
				End If
				If checkType="1" Then 
					conn.execute("update T_NOTICE set IS_TOP=0 where ID="&udID(i))
					operateCon="取消置顶"
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "notice.asp",operateCon&"："&subBox
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
