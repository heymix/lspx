<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
ID	 		 = Request("ID")
subBox	 	= Request("subBox")
test_id	= Request.Cookies("test")("test_id")

'response.write "0|()"&Request(tt)&"()"&USER_NAME&"()"&PASSWORD
'Call appEnd() 




If PURVIEW="" or  subBox=""   Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
Else

	If PURVIEW="0" Then
		updateStr=""
		conn.begintrans
		udID=split(subBox,",")
		for i=0 to Ubound(udID)
			keyID=split(udID(i),"|")
			If Ubound(keyID)=1 Then
				If isNumeric(keyID(1))  Then
					exam_score=request(trim(keyID(0))&trim(keyID(1)))
					exam_pass=request(trim(keyID(0))&trim(keyID(1))&"_check")
					If exam_pass<>1 Then exam_pass=0
					set rs=server.createobject("ADODB.Recordset") 
					sqlTopic="select * from T_TEST_TOPIC where ID="&keyID(1)
					rs.open sqlTopic,conn,1,3
					If not (rs.eof or rs.bof) Then
					cc=rs("IS_PRATICE")
						If rs("IS_PRATICE")=0 Then
							If exam_score>=60 Then
								exam_pass=1
							Else
								exam_pass=0
							End If
						End If
					End If
					rs.close
					set rs=nothing
					
					
					sql="update T_TOPIC_RESULT set RESULT="&exam_score&",IS_PASS="&exam_pass&" where TEST_ID='"&test_id&"' and USER_ID='"&trim(keyID(0))&"' and TOPIC_ID="&keyID(1)
					'response.write sql
					conn.execute sql,updateNO
					If updateNO=0 Then
						updateStr=updateStr&"未更新的。用户名："&keyID(0)&" 科目："&keyID(1)
					End If
				End If
			End If
			keyID=""
			set keyID=nothing
		next
		if Err.Number = 0 Then
			operateLog "score.asp",operateCon&"："&subBox
			Response.write "1|操作成功！"&cc
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
