<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
operateCon=""
'Response.write "0|"&subBox&oStatus&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" or checkType="" or oStatus=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW<>"" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If checkIDCard(trim(udID(i))) Then
				If oStatus="0" Then
					If checkType="0" Then 
						conn.execute("update T_EXAMINEE set CHECKED=1 where USER_ID='"&trim(udID(i))&"'")
						operateCon="审核"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXAMINEE set CHECKED=0 where USER_ID='"&trim(udID(i))&"'")
						'response.write ("update T_EXAMINEE set CHECKED=0 where USER_ID='"&trim(udID(i))&"'")
						operateCon="反审核"
					End If
				ElseIf oStatus="1" Then
					If checkType="0" Then 
						conn.execute("update T_EXAMINEE set LOCKED=1 where USER_ID='"&trim(udID(i))&"'")
						operateCon="锁定"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXAMINEE set LOCKED=0 where USER_ID='"&trim(udID(i))&"'")
						operateCon="解锁"
					End If
				ElseIf oStatus="4" Then
						set rs=conn.execute("select count(*) from T_TEST_RESULT where USER_ID='"&trim(udID(i))&"'")
							notIsNull=rs(0)
						rs.close
						set rs=nothing
						If notIsNull=0 Then
							conn.execute("delete  from T_EXAMINEE where USER_ID='"&trim(udID(i))&"'")
						Else
							notDelUser=notDelUser&trim(udID(i))&","
						End if
					operateCon="删除"
				ElseIf oStatus="5" Then
					conn.execute("update T_EXAMINEE set PASSWORD='"&md5("ks123456")&"' where USER_ID='"&trim(udID(i))&"'")
					operateCon="删除"
					resetPsd="密码被重置为：ks123456"

				End If
			End If
		next
		if Err.Number = 0 Then
			If notDelUser<>"" Then 
				notDelUser="其中"&notDelUser&"删除未成功，已经参与考试不允许删除，如有必要可以锁定用户！"
			End If
			operateLog "examinee.asp",operateCon&"："&subBox&noCheckUser&resetPsd&notDelUser
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
