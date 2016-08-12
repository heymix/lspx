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
If PURVIEW="" or  subBox="" or checkType="" or oStatus=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
	If PURVIEW="0" or PURVIEW="3" Then
		conn.begintrans
		udID=split(subBox,",")
		updateStr=""
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="11" Then
					If checkType="0" Then 
						conn.execute("update T_EXA_ROOM set IS_RUN=1 where ID="&udID(i))
						set rs=server.createobject("adodb.recordset")
						sql="SELECT RSI.USER_ID,RSI.SEAT_ID,ER.T_ROOM_ORDER,ER.EXA_ID FROM T_EXA_ROOM_SEAT_ID AS RSI left join T_EXA_ROOM AS ER ON RSI.ROOM_ID=ER.ID  where ROOM_ID="&udID(i)&" And RSI.TEST_ID="&test_id
						'response.write sql
						rs.open sql,conn,1,1
						do while not (rs.eof or err)
							If rs("USER_ID")<>"" Then
								sql="update T_TEST_RESULT set TEST_NO='"&mid(year(now()),3,4)&numberCover(rs("EXA_ID"))&numberCover(rs("T_ROOM_ORDER"))&numberCover(rs("SEAT_ID"))&"' where TEST_ID="&test_id&" and USER_ID='"&rs("USER_ID")&"'"
								'response.Write sql
								conn.execute sql,updateNo
								If updateNo=0 Then
									updateStr=updateStr&rs("USER_ID")&","
								End If
							End If
						rs.moveNext
						loop
						rs.close
						set rs=nothing
						operateCon="生成考号"
					End If
					If checkType="1" Then 
						conn.execute("update T_EXA_ROOM set IS_RUN=0 where ID="&udID(i))
						set rs=server.createobject("adodb.recordset")
						sql="SELECT RSI.USER_ID,RSI.SEAT_ID,ER.T_ROOM_ORDER,ER.EXA_ID FROM T_EXA_ROOM_SEAT_ID AS RSI left join T_EXA_ROOM AS ER ON RSI.ROOM_ID=ER.ID  where ROOM_ID="&udID(i)&" And RSI.TEST_ID="&test_id
						'response.write sql
						rs.open sql,conn,1,1
						do while not (rs.eof or err)
							If rs("USER_ID")<>"" Then
								sql="update T_TEST_RESULT set TEST_NO='' where TEST_ID="&test_id&" and USER_ID='"&rs("USER_ID")&"'"
								'response.Write sql
								conn.execute sql,updateNo
								If updateNo=0 Then
									updateStr=updateStr&rs("USER_ID")&","
								End If
							End If
						rs.moveNext
						loop
						rs.close
						set rs=nothing
						operateCon="撤消生成"
					End If
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "exam_room.asp",operateCon&"："&subBox
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
