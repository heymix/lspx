<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
test_id	= Request.Cookies("test")("test_id")
operateCon=""
'Response.write "0|"&subBox&oStatus&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" or checkType="" or oStatus=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	notIsNull=0
	If PURVIEW<>"" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If checkIDCard(trim(udID(i))) Then
		
				If oStatus="5" Then 
					conn.execute("update T_TEST_RESULT set IS_PRINT=1 where TEST_ID="&test_id&" and USER_ID='"&trim(udID(i))&"'")
					operateCon="标记"
				End If
				If oStatus="6" Then 
					conn.execute("update T_TEST_RESULT set IS_PRINT=0 where TEST_ID="&test_id&" and USER_ID='"&trim(udID(i))&"'")
					operateCon="取消标记"
				End If
				
				If oStatus="8" Then 
					conn.execute("update T_TEST_RESULT set CRS_NO='' where TEST_ID="&test_id)
					conn.execute("update T_TEST set CERT_NO_N=CERT_NO where ID="&test_id)
				End If
	
				
			End If
		next
		if Err.Number = 0 Then
			operateLog "pring_credentials.asp",operateCon&"："&subBox&resetPsd&noCheckUser&notDelUser
			Response.write "1|操作成功！"&noCheckUser&resetPsd&notDelUser
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
