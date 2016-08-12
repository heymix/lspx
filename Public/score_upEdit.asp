<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
subBox=request("subBox")
'Response.write subBox&"ssss"
'appEnd()
If subBox="" Then
	Response.write "0|失败，上传失败 文件名丢失！"
	appEnd()
End If
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set f=fs.OpenTextFile(Server.MapPath("../UploadFiles/score/"&subBox), 1)  '打开文本文件

returnStr="<table  class=""list_table"" width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" >"  '建立表格
j=0
do while f.AtEndOfStream = false
	str=f.ReadLine
	splitstr=split(str,",")  '分割一行的字符串
	returnStr=returnStr&"<tr>"   '建立表格行
	If j=0 Then
		for i =0 to ubound(splitstr)
			returnStr=returnStr&"<th>"   '建立表格列
			returnStr=returnStr& splitstr(i)
			returnStr=returnStr&"</th>"
		next
	Else
		for i =0 to ubound(splitstr)
			returnStr=returnStr&"<td>"   '建立表格列
			returnStr=returnStr& splitstr(i)
			returnStr=returnStr&"</td>"
		next
	End If
	j=j+1	
	returnStr=returnStr&"</tr>"
loop
returnStr=returnStr&"</table>"

f.Close
Set f=Nothing
Set fs=Nothing
If err.Number=0 Then
	Response.write returnStr
Else
	Response.write "0|上传失败|"
End If
	
%>
