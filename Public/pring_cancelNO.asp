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
If PURVIEW=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
operateCon="取消证书编号生成"
	notIsNull=0
	If PURVIEW<>"" Then
		conn.begintrans

					conn.execute("update T_TEST_RESULT set CRS_NO='' where TEST_ID="&test_id)
					conn.execute("update T_TEST set CERT_NO_N=CERT_NO where ID="&test_id)

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
