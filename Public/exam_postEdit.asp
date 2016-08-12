<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../Examinee/CheckExamineeOL.asp"-->
<%
ID	 		 = Request("ID")
IS_MAKEUP	 = Request("IS_MAKEUP_bm")
operateType	 = Request("operateType")
subBox	 	 = Request("subBox")
myTEST_ID	 = Request("myTEST_ID")

'response.write "0|"&IS_MAKEUP&"()"&isSame&"()"&SCHOOL&"()"&SCHOOL_ID
'Call appEnd() 
If myTest_id=""  Then 
	Response.write "0|操作失败,参数丢失！"
	Call appEnd()
End If

'判断资料完整性
rs_check_num=0
set rs=conn.execute("select COUNT(*) from T_EXAMINEE where USER_ID='"&EXAMINEE_USER_ID&"' and PHOTO<>''")
	rs_check_num=rs(0)
rs.close
set rs=nothing
If rs_check_num=0 THen
	Response.write "0|操作失败,请完善资料再进行报告！"
	Call appEnd()
End If

If IS_MAKEUP<>"1" Then IS_MAKEUP="0"

If IS_MAKEUP="1" Then
	If subBox=""  Then 
		Response.write "0|操作失败,补考科目没有选择！"
		Call appEnd()
	End If
End If

If operateType=0 Then

		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST_RESULT Where TEST_ID="&myTEST_ID&" and USER_ID='"&EXAMINEE_USER_ID&"'"
		'Response.write  sql
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("USER_ID")	=EXAMINEE_USER_ID
				rs("TEST_ID")	=myTEST_ID
				rs("IS_MAKEUP")	=IS_MAKEUP
				rs("TEST_YEAR")	=year(now())
				rs("IS_MAKEUP")	=IS_MAKEUP
				rs("INFO_TIME")	=ToUnixTime(now(),+8)
			rs.update
			rs.close
			set rs= nothing
			If IS_MAKEUP="0" Then
				set rs=server.createobject("ADODB.Recordset")
				sql="select TTR.TOPIC_ID,TT.IS_PRATICE from T_TEST_TOPIC_RELATION AS TTR Left Join T_TEST_TOPIC AS TT ON TTR.TOPIC_ID=TT.ID where TTR.TEST_ID="&myTEST_ID
				rs.open sql,conn,1,1
				do while not (rs.eof or err)
					If rs("IS_PRATICE")=1 Then
						conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_MAKEUP,RESULT,IS_PASS) values('"&EXAMINEE_USER_ID&"',"&myTEST_ID&","&rs("TOPIC_ID")&","&IS_MAKEUP&",1,1)")
					Else
						conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_MAKEUP,IS_PASS) values('"&EXAMINEE_USER_ID&"',"&myTEST_ID&","&rs("TOPIC_ID")&","&IS_MAKEUP&",1)")
					End If
				rs.movenext
				loop
				rs.close 
				set rs=nothing
				
			Else
				udID=split(subBox,",")
				for i=0 to Ubound(udID)
					If isNumeric(udID(i)) Then
						set rs_topic=conn.execute("select IS_PRATICE from T_TEST_TOPIC WHERE ID="&udID(i))
						IS_PRATICE=rs_topic("IS_PRATICE")
						rs_topic.close
						set rs_topic=nothing
						If IS_PRATICE=1 Then
							conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_MAKEUP,RESULT) values('"&EXAMINEE_USER_ID&"',"&myTEST_ID&","&udID(i)&","&IS_MAKEUP&",1)")
						Else
							conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_MAKEUP) values('"&EXAMINEE_USER_ID&"',"&myTEST_ID&","&udID(i)&","&IS_MAKEUP&")")
						End If
						
					End If
				next
			End If
			
			
			if Err.Number = 0 Then
				operateLog "exam_post.asp","报考：身份证号："&EXAMINEE_USER_ID&" 报名："&myTEST_ID
				Response.write "1|报名成功！"
				conn.CommitTrans 
			Else
				Response.write "0|报名失败！"
				conn.RollbackTrans
			End If
		Else
			Response.write "0|操作失败！已经报名！"
			conn.RollbackTrans
		End If
ElseIf operateType=1 Then
		If not is_Id(myTEST_ID) Then
			Response.write "0|操作失败,系统编号错误！"
			Call appEnd()
		End If
		conn.begintrans
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_TEST_RESULT Where TEST_ID="&myTEST_ID&" and USER_ID='"&EXAMINEE_USER_ID&"'"
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("IS_MAKEUP")	=IS_MAKEUP
			rs.update
			rs.close
			set rs= nothing
			
			conn.execute("delete from T_TOPIC_RESULT where TEST_ID="&myTEST_ID&" and USER_ID='"&EXAMINEE_USER_ID&"'")
			If IS_MAKEUP="0" Then
				conn.execute("insert into T_TOPIC_RESULT(USER_ID,TEST_ID,TOPIC_ID)select "&EXAMINEE_USER_ID&","&myTEST_ID&",TOPIC_ID from T_TEST_TOPIC_RELATION where TEST_ID="&myTEST_ID)
			Else
			
				udID=split(subBox,",")
				for i=0 to Ubound(udID)
					If isNumeric(udID(i)) Then
						conn.execute("insert into T_TOPIC_RESULT (USER_ID,TEST_ID,TOPIC_ID,IS_MAKEUP) values("&EXAMINEE_USER_ID&","&myTEST_ID&","&udID(i)&",1)")
					End If
				next
			End If
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
