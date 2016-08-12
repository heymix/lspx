<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

' 
subBox	= Request("subBox")
oStatus = Request("status")
t_max_num = Request("t_max_num")

operateCon=""
'Response.write "1|"&t_max_num&oStatus&subBox&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If test_id="" or  subBox="" or oStatus="" or sys_exam_id=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	If is_test_end(test_id) Then
		Response.write "1|考试已经结束,信息已存档,不允许修改相关信息！"
		appEnd()
	End If

	If oStatus="0" Then
		isMakUP=" and IS_MAKEUP=0"
	End If
	If oStatus="1" Then
		isMakUP=" and IS_MAKEUP=1"
	End If
	set rs=server.createobject("adodb.recordset")
	sql="select USER_ID from V_EXAMINEE_TEST where TRCHECKED=1 and TEST_ID="&test_id&" and SCHOOL_ID in (select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&") and USER_ID not  in (select USER_ID From T_EXA_ROOM_SEAT_ID where TEST_ID="&test_id&")"&isMakUP
	rs.open sql,conn,1,1
		If rs.eof or rs.bof Then
			USER_ID_STR=""
		Else
			do while not (rs.eof or err)	
				USER_ID_STR=USER_ID_STR&rs("USER_ID")&","
				rs.movenext
			loop
		End If
	rs.close
	set rs=nothing
If USER_ID_STR="" Then
	Response.write "0|失败：没有查到插入的用户！"
appEnd()
End If
USER_ID_STR=mid(USER_ID_STR,1,len(USER_ID_STR)-1)
USER_ID_ARR=split(USER_ID_STR,",")


'随机重排分组
leng=UBound(USER_ID_ARR)
randomize  
for i=0 to leng-1  
	 b=int(rnd()*leng)  
	 temp=USER_ID_ARR(b)  
	 USER_ID_ARR(b)=USER_ID_ARR(i)  
	 USER_ID_ARR(i)=temp
Next
'插科记录
	If PURVIEW="0" or PURVIEW="3" Then
		conn.begintrans
		j=0
		for i=1 to t_max_num
			If j<=leng Then
				set rs=server.createobject("adodb.recordset")
				sql="select USER_ID from T_EXA_ROOM_SEAT_ID WHERE USER_ID<>'' and TEST_ID="&test_id&" and ROOM_ID="&subBox&" and SEAT_ID="&i
				'Response.write sql
				rs.open sql,conn,1,3
				If rs.eof or rs.bof Then
					conn.execute("update T_EXA_ROOM_SEAT_ID set USER_ID='"&trim(USER_ID_ARR(j))&"' WHERE TEST_ID="&test_id&" and ROOM_ID="&subBox&" and SEAT_ID="&i)
					j=j+1
				End If
				rs.close
				set rs=nothing
			Else
				exit for
			End If
			
		next
		if Err.Number = 0 Then
			operateLog "exam_room_seat_distribute.asp",operateCon&"教室编号："&subBox&" 考试ID："&test_id
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
