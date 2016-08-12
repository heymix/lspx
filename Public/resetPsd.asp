<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%
subBox	= Request("subBox")

'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox=""    Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				conn.execute("update T_ADMIN set PASSWORD='"&md5("h123456")&"' where ID="&udID(i))
			End If
		next
		if Err.Number = 0 Then
			operateLog "manager.asp  ","重置用户密码："&subBox
			Response.write "1|操作成功！密码被重置为:h123456"
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
