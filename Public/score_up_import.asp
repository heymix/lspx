<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
score_files=request("score_files")
'Response.write subBox&"ssss"
'appEnd()
If score_files="" Then
	Response.write "0|失败，上传失败 文件名丢失！"
	appEnd()
End If
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set f=fs.OpenTextFile(Server.MapPath("../UploadFiles/score/"&score_files), 1)  '打开文本文件

conn.begintrans
err_str=""
j=0
do while f.AtEndOfStream = false
	str=f.ReadLine
	splitstr=split(str,",")  '分割一行的字符串
	If j>0 Then
		sql="UPDATE T_TOPIC_RESULT SET  RESULT ="&splitstr(5)&" ,IS_PASS ="&splitstr(6)&" ,IS_SCORE =1 WHERE USER_ID='"&splitstr(0)&"' and  TEST_ID="&test_id&" and TOPIC_ID="&splitstr(3)
		'response.write sql
		conn.execute sql,updateNO
		if updateNO=0 Then
			err_str=err_str&splitstr(0)&","
		End If
	End If
		j=j+1
loop

f.Close
Set f=Nothing
Set fs=Nothing
	if Err.Number = 0 Then
		operateLog "score_up_import.asp",operateCon&"："&subBox
		if err_str<>"" Then err_str= "其中"&err_str&"未成功！"
		Response.write "1|操作成功！"&err_str
		conn.CommitTrans 
	Else
		Response.write "0|操作失败！"
		conn.RollbackTrans
	End If
appEnd()	
%>
