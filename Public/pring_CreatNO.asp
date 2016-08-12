<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%

'锁定  status 0 审校 1置顶 2刷新 4删除
subBox	= Request("subBox")
checkType = Request("checkType")
oStatus = Request("status")
test_id	= Request.Cookies("test")("test_id")
sys_exam_id	= Request.Cookies("test")("sys_exam_id")
operateCon=""
'Response.write "0|"&subBox&oStatus&checkType
'Response.End()
'//审核  operateType single multi多选单选 checkType 0 复核 1 反复核
If PURVIEW=""  Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else
	set rs=conn.execute("select CERT_NO,CERT_NO_N from T_TEST where ID="&test_id&"")
	If not (rs.eof or err) Then
	operateCon="生成证书"
		notIsPass=0
		If isNull(rs("CERT_NO_N")) Then
			CERT_START=rs("CERT_NO")
		Else
			CERT_START=rs("CERT_NO_N")
		End If
		CERT_NO=rs("CERT_NO")
	End If
	rs.close
	set rs=nothing
	subBox=""
	set rs=conn.execute("select USER_ID from V_TEST_RESULT where TEST_ID="&test_id&" and SCHOOL_ID in(select SCHOOL_ID From T_EXA_SCHOOL_RELATION where EXA_ID="&sys_exam_id&") order by TEST_NO ASC")
		do while not(rs.eof or err)
			subBox=subBox&","&rs("USER_ID")
		rs.moveNext
		loop
	rs.close
	set rs=nothing
	If len(subBox)>15 Then
		subBox=mid(subBox,2,len(subBox)-1)
	End If
		If PURVIEW<>"" Then
			conn.begintrans
			udID=split(subBox,",")
			If not(isNull(CERT_NO)) Then
				for i=0 to Ubound(udID)
					If checkIDCard(trim(udID(i))) Then
							notIsPass=0
							set rs=conn.execute("select count(USER_ID) from T_TOPIC_RESULT where IS_PASS=0 and USER_ID='"&trim(udID(i))&"' and TEST_ID="&test_id)
							notIsPass=rs(0)
							rs.close
							set rs=nothing
							
						If notIsPass=0 Then
							CERT_START=CERT_START+1
							conn.execute("update T_TEST_RESULT set CRS_NO='"&CERT_START&"' where TEST_ID="&test_id&" and  USER_ID='"&trim(udID(i))&"'")
							conn.execute("update T_TEST set CERT_NO_N='"&CERT_START&"' where ID="&test_id&"")
						End If
					End If
				next
			End If
			if Err.Number = 0 Then
				operateLog "pring_credentials.asp",operateCon&"："&sys_exam_id
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
