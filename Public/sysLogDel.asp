<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%


subBox	= Request("subBox")
operateType= Request("operateType")
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If
If PURVIEW<>"0" Then 
	Response.write "0|操作失败！没有权限"
	Call appEnd()
End If
	
	If operateType=0 Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then 
				conn.execute("delete  from T_LOG  where OPERATE_TIME="&udID(i))
			End If
		next
		if Err.Number = 0 Then
			Response.write "1|操作成功！"
			conn.CommitTrans 
		Else
			Response.write "0|操作失败！"
			conn.RollbackTrans
		End If
	ElseIf operateType=1 Then
		conn.execute("delete  from T_LOG ")
		if Err.Number = 0 Then
			Response.write "1|操作成功！"
		Else
			Response.write "0|操作失败！"
		End If
	End If

call CloseConn()
%>
