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
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
	noCheckUser=""
	notDelUser=""
	resetPsd=""
	notIsNull=0
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
				ElseIf oStatus="3" Then
					If checkType="0" Then 
						conn.execute("update T_TEST_RESULT set CHECKED=1 where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
						operateCon="审核"
					End If
					If checkType="1" Then
						set rs=conn.execute("select count(*) from T_EXA_ROOM_SEAT_ID where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
						noSeat=rs(0)
						rs.close
						set rs=nothing
						operateCon="反审核"
						If noSeat=0 Then
							conn.execute("update T_TEST_RESULT set CHECKED=0 where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
						Else
							noCheckUser=noCheckUser&trim(udID(i))&","
						End if
					End If
				ElseIf oStatus="4" Then
						'删除考试信息  报考信息 科目信息 分座信息
						conn.execute("delete  from T_TEST_RESULT where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
						conn.execute("delete  from T_TOPIC_RESULT where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
						conn.execute("delete  from T_EXA_ROOM_SEAT_ID where USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
					operateCon="删除报考信息 考试ID："&test_id
				ElseIf oStatus="5" Then
					conn.execute("update T_EXAMINEE set PASSWORD='"&md5("ks123456")&"' where USER_ID='"&trim(udID(i))&"'")
					operateCon="重置密码"
					resetPsd="密码被重置为：ks123456"
				End If
				
			End If
		next
		if Err.Number = 0 Then
			If noCheckUser<>"" Then 
				noCheckUser="其中"&noCheckUser&"反审核未成功，已经分配座会。请移除座位再反审核"
			End If
			If notDelUser<>"" Then 
				notDelUser="其中"&notDelUser&"删除未成功，已经参与考试不允许删除，如有必要可以锁定用户！"
			End If
			operateLog "examinee.asp",operateCon&"："&subBox&resetPsd&noCheckUser&notDelUser
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
