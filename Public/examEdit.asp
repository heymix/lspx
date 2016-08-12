<!--#include file="../Inc/conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../inc/CheckUserOL.asp"-->
<%
ID	 		 = Request("ID")
isSame	     = Request("isSame")
operateType	 = Request("operateType")
SCHOOL	 	 = Request("inputString")
SCHOOL_ID	 = Request("SCHOOL_ID")
TEST_CATE	 = Request("TEST_CATE")
CITY_ID	 	 = Request("CITY_ID")
EXA_NAME	 = Request("EXA_NAME")
CHECKED	 	 = Request("CHECKED")
	

'response.write "0|"&PURVIEW&"()"&isSame&"()"&SCHOOL&"()"&SCHOOL_ID
'Call appEnd() 
If PURVIEW="" or  SCHOOL="" or not is_Id(SCHOOL_ID) or not is_Id(TEST_CATE) or TEST_CATE="0" or not is_Id(CITY_ID) or CITY_ID="0" Then 
	Response.write "0|操作失败,参数丢失,检查登录状态！"
	Call appEnd()
End If

'If sqlInjection(opValue) Then
'	Response.write "0|操作失败,名称包含敏感信息！"
'	Call appEnd()
'End If

If PURVIEW<>"0" Then
	Response.write "0|操作失败,没有权限！"
	Call appEnd()
End If
If CHECKED<>"1" Then CHECKED="0"
If isSame="1"  Then EXA_NAME=SCHOOL
If operateType=0 Then
		srd=conn.execute("select count(*) from T_EXA_SCHOOL_RELATION where TEST_CATE="&TEST_CATE&" and SCHOOL_ID="&SCHOOL_ID)(0)
		If srd>0 Then
			Response.write "0|操作失败,考点已经存在，或者所选学校，在其它考点中！"
			Call appEnd()
		End If
		conn.begintrans
		autoID=getID("T_EXA")
		set rs=server.createobject("ADODB.Recordset")
		sql="select * From T_EXA Where TEST_CATE="&TEST_CATE&" and SCHOOL_ID="&SCHOOL_ID
		'Response.write  sql
		rs.open sql,conn,1,3
		If rs.eof or rs.bof Then
			rs.AddNew
				rs("ID")=autoID
				rs("EXA_NAME")=EXA_NAME
				rs("TEST_CATE")	= TEST_CATE
				rs("CITY_ID")	= CITY_ID
				rs("SCHOOL_ID")=SCHOOL_ID
				rs("CHECKED")=CHECKED
				rs("INFO_TIME")=ToUnixTime(now(),+8)
			rs.update
			rs.close
			conn.execute("insert into T_EXA_SCHOOL_RELATION (TEST_CATE,SCHOOL_ID,EXA_ID) values("&TEST_CATE&","&SCHOOL_ID&","&autoID&")")
			set rs= nothing
			if Err.Number = 0 Then
				operateLog "exam_add.asp","添加：系统编号："&autoID&" 考点名称："&EXA_NAME
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
		sql="select * From T_EXA Where ID="&ID
		rs.open sql,conn,1,3
		If not(rs.eof or rs.bof) Then
				rs("EXA_NAME")	= EXA_NAME
				rs("TEST_CATE")	= TEST_CATE
				rs("CITY_ID")	= CITY_ID
				rs("CHECKED")	= CHECKED
			rs.update
			rs.close
		
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
