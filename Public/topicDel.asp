<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%


subBox	= Request("subBox")
operateCon=""
'Response.write "1|"&subBox&operateType&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW="" or  subBox="" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If
	noOperateID=""
	rs_delCheck_num=0
	If PURVIEW="0" Then
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			If isNumeric(udID(i)) Then 
				set rs_delCheck=conn.execute("SELECT count(*)  FROM T_TEST_TOPIC_RELATION where TOPIC_ID="&udID(i)&"")
			  		rs_delCheck_num=rs_delCheck(0)
			   	rs_delCheck.close
			   	set rs_delCheck=nothing
			   	If rs_delCheck_num=0 Then
					conn.execute("delete  from T_TEST_TOPIC  where ID="&udID(i))
					operateCon="删除"
				Else
						noOperateID=noOperateID&udID(i)&","
				End If
			End If
		next
		if Err.Number = 0 Then
			If noOperateID<>"" Then
				noOperateID="其中"&noOperateID&"未操作成功，原因：科目已经被应用于考试"
			End If
			operateLog "topic.asp",operateCon&"："&subBox&noOperateID
			Response.write "1|操作成功！"&noOperateID
			conn.CommitTrans 
		Else
			Response.write "0|操作失败！"
			conn.RollbackTrans
		End If
	Else
		Response.write "0|操作失败！没有权限"
	End If

call CloseConn()
%>
