<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
ID	 		 = Request("ID")
EXA_ID	     = sys_exam_id
operateType	 = Request("operateType")
T_ROOM	 	 = Request("T_ROOM")
T_ROOM_ORDER = Request("T_ROOM_ORDER")
T_MAX_NUM	 = Request("T_MAX_NUM")
CHECKED	 	 = Request("CHECKED")

'response.write "0|"&PURVIEW&"()"&T_ROOM&"()"&T_ROOM_ORDER&"()"&T_MAXNUM&"()"&EXA_ID
'Call appEnd() 

If PURVIEW="" or  T_ROOM="" or not is_Id(T_MAX_NUM) or not is_Id(T_ROOM_ORDER) Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

'If sqlInjection(opValue) Then
'	Response.write "0|操作失败,名称包含敏感信息！"
'	Call appEnd()
'End If

If not(PURVIEW="0" or PURVIEW="3") Then
	Response.write "0|操作失败,没有权限！"
	Call appEnd()
End If
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If
If CHECKED<>"1" Then CHECKED="0"
If isSame="1"  Then EXA_NAME=SCHOOL
If operateType=0 Then
		conn.begintrans
		autoID=getID("T_EXA_ROOM")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXA_ROOM Where TEST_ID="&test_id&" and EXA_ID="&EXA_ID&" and T_ROOM_ORDER="&T_ROOM_ORDER
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID")=autoID
				rs("EXA_ID")=EXA_ID
				rs("TEST_ID")=TEST_ID
				rs("T_ROOM")=T_ROOM
				rs("T_ROOM_ORDER")=T_ROOM_ORDER
				rs("T_MAX_NUM")=T_MAX_NUM
				rs("CHECKED")=CHECKED
				rs("INFO_TIME")=ToUnixTime(now(),+8)
			rs.update
			rs.close
			for i=1 to  T_MAX_NUM
				conn.execute("insert into T_EXA_ROOM_SEAT_ID(ROOM_ID,TEST_ID,SEAT_ID) values("&autoID&","&test_id&","&i&")")
			next 
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "exam_room_add.asp","添加：系统编号："&autoID&" 考场号名称："&T_ROOM
				Response.write "1|添加成功！"
				conn.CommitTrans 
			Else
				Response.write "0|添加失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|添加失败！记录已经存在."
			conn.RollbackTrans
		End If
ElseIf operateType=1 Then
		If not is_Id(ID) Then
			Response.write "0|操作失败,系统编号错误！"
			Call appEnd()
		End If

		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXA_ROOM Where ID="&ID
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				old_T_MAX_NUM=rs("T_MAX_NUM")
				rs("EXA_ID")=EXA_ID
				rs("T_ROOM")=T_ROOM
				rs("T_ROOM_ORDER")=T_ROOM_ORDER
				rs("T_MAX_NUM")=T_MAX_NUM
				rs("CHECKED")=CHECKED
			rs.update
			rs.close
			If Cint(old_T_MAX_NUM)>Cint(T_MAX_NUM) Then
				'conn.execute("delete from T_EXA_ROOM_SEAT_ID where ROOM_ID="&ID)
				for i=T_MAX_NUM+1 to old_T_MAX_NUM
					conn.execute("delete from T_EXA_ROOM_SEAT_ID where ROOM_ID="&ID&" and SEAT_ID="&i)
				next
			End If
			If Cint(old_T_MAX_NUM)<Cint(T_MAX_NUM) Then
				for i=old_T_MAX_NUM+1 to  T_MAX_NUM
					conn.execute("insert into T_EXA_ROOM_SEAT_ID(ROOM_ID,TEST_ID,SEAT_ID) values("&ID&","&test_id&","&i&")")
				next 
			End if
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "exam_add.asp","修改：系统编号："&id&" 名称："&TITLE
				Response.write "1|修改成功！"
				conn.CommitTrans 
			Else
				Response.write "0|修改失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|修改失败！记录不存在."
			conn.RollbackTrans
		End If
Else
	Response.write "0|添加失败！操作符错误 请联系管理员。"
End If
call CloseConn()
%>
