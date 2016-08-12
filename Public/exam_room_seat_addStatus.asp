<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
ROOM_ID = Request("ROOM_ID")
ID=Request("ID")'座位增加学生的 座位ID组

operateCon=""
'Response.write "1|"&test_id&ROOM_ID&subBox&"---"&ID
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If test_id="" or  ROOM_ID="" or oStatus=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
 If ID<>"" Then
 	uSearID=Split(ID,",")
 End If
	If PURVIEW="0" or PURVIEW="3" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="0" Then
					'坐满之后 就不再增加了
					If i<=Ubound(uSearID)Then
						If checkType=0 Then '0  是新增
							'清除之前的座位号 添加到新的座位
							conn.execute("update T_EXA_ROOM_SEAT_ID  set USER_ID=''  where ROOM_ID="&ROOM_ID&" and TEST_ID="&test_id&" and USER_ID='"&trim(udID(i))&"'")
							conn.execute("update T_EXA_ROOM_SEAT_ID  set USER_ID='"&trim(udID(i))&"' where ROOM_ID="&ROOM_ID&" and TEST_ID="&test_id&" and SEAT_ID="&uSearID(i))
							
						ElseIf checkType=1 Then '调换
						
							set rs=server.createobject("adodb.recordset")'传过来的用户名
							sql="select USER_ID from T_EXA_ROOM_SEAT_ID WHERE TEST_ID="&test_id&" and ROOM_ID="&ROOM_ID&" and SEAT_ID="&uSearID(i)
							rs.open sql,conn,1,3
							If rs.eof or rs.bof Then
								post_USER_ID="{||}"
							Else	
								post_USER_ID=rs("USER_ID")
							End If
							rs.close
							set rs=nothing		
							conn.execute("update T_EXA_ROOM_SEAT_ID  set USER_ID='"&post_USER_ID&"' where USER_ID='"&udID(i)&"' and test_id="&test_id)
							conn.execute("update T_EXA_ROOM_SEAT_ID  set USER_ID='"&trim(udID(i))&"' where ROOM_ID="&ROOM_ID&" and TEST_ID="&test_id&" and SEAT_ID="&uSearID(i))
						End If
					End If
				ElseIf oStatus="4" Then
					conn.execute("update T_EXA_ROOM_SEAT_ID  set USER_ID=''  where ROOM_ID="&ROOM_ID&" and TEST_ID="&test_id&" and SEAT_ID="&udID(i))
					operateCon="删除"
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "exam_room_seat.asp",operateCon&"教室编号："&ROOM_ID&" 考试ID："&test_id&"座位号："&subBox
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
