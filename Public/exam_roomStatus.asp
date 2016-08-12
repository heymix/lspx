<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
operateCon=""
'Response.write "0|"&subBox&"++"&operateType&"++"&checkType&"++"&oStatus
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
'教室座位上有人 不可以删除
If PURVIEW="" or  subBox="" or checkType="" or oStatus=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
	noOperateID=""
	rs_delCheck_num=0
	If PURVIEW="0" or  PURVIEW="3"  Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="0" Then
					If checkType="0" Then 
						conn.execute("update T_EXA_ROOM set CHECKED=1 where  ID="&udID(i))
						operateCon="审核"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXA_ROOM set CHECKED=0 where ID="&udID(i))
						operateCon="反审核"
					End If
				ElseIf oStatus="1" Then
					If checkType="0" Then 
						conn.execute("update T_EXA_ROOM set IS_Run=1 where ID="&udID(i))
						operateCon="提交"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXA_ROOM set IS_Run=0 where ID="&udID(i))
						operateCon="撤回"
					End If
				ElseIf oStatus="4" Then
					 set rs_delCheck=conn.execute("SELECT COUNT(*) FROM [T_EXA_ROOM_SEAT_ID] where USER_ID<>'' and ROOM_ID="&udID(i)&" and TEST_ID="&test_id&"")
					 rs_delCheck_num=rs_delCheck(0)
					 rs_delCheck.close
					 set rs_delCheck=nothing
					 If rs_delCheck_num=0 Then
						 conn.execute("delete  from T_EXA_ROOM  where ID="&udID(i))
						 conn.execute("delete  from T_EXA_ROOM_SEAT_ID  where ROOM_ID="&udID(i))
						operateCon="删除了教室下所有座位 考点ID"
					Else
						noOperateID=noOperateID&udID(i)&","
					End If
				End If
			End If
		next
		if Err.Number = 0 Then
			If noOperateID<>"" Then
				noOperateID="其中"&noOperateID&"未操作成功，原因：座位上有人"
			End If
			operateLog "exam_room.asp",operateCon&"："&subBox&noOperateID
			Response.write "1|操作成功！"&noOperateID
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
