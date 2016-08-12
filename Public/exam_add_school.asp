<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
EXA_ID	= Request("EXA_ID")
TEST_CATE = Request("TEST_CATE")
oStatus = Request("status")
operateCon=""
'Response.write "0|"&EXA_ID&"--"&TEST_CATE&"--"&subBox
'Response.End()
'
If PURVIEW="" or  subBox="" or not is_Id(EXA_ID) or not is_Id(TEST_CATE)  or oStatus=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" or PURVIEW="3" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then
				If oStatus="5" Then 
					conn.execute("DELETE T_EXA_SCHOOL_RELATION where EXA_ID="&EXA_ID&" AND TEST_CATE="&TEST_CATE&" AND SCHOOL_ID="&udID(i))
					'operateCon="从考点("&test_id&":"&test_title&")中移除学校"
				End If
				If oStatus="6" Then 
					theSame=conn.execute("select count(*) from T_EXA_SCHOOL_RELATION where TEST_CATE="&TEST_CATE&" AND SCHOOL_ID="&udID(i))(0)			
					If theSame=0 Then
						conn.execute("insert into T_EXA_SCHOOL_RELATION (TEST_CATE,SCHOOL_ID,EXA_ID)values("&TEST_CATE&","&udID(i)&","&EXA_ID&")")
					End If
					operateCon="添加学校到考点("&subBox&":"&EXA_ID&")"
				End If
			End If
		next
		if Err.Number = 0 Then
			operateLog "exam_add_school.asp",operateCon&"："&subBox
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
