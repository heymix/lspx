<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--include file="../inc/CheckUserOL.asp"-->
<!--#include file="../inc//Cls_ShowoPage.asp"-->
<%
'-----------------------------------------------------------------------------------------------
DIM startime,endtime
'统计执行时间
startime=timer()

'-----------------------------------------------------------------------------------------------
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>日志文件</title>
<link href="../Css/main.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>

 <script type="text/javascript">
        $(function() {
	
           $("#checkAll").click(function() {
                $('input[name="subBox"]').attr("checked",this.checked); 
            });
            var $subBox = $("input[name='subBox']");
            $subBox.click(function(){
				$("#checkAll").attr("checked",$subBox.length == $("input[name='subBox']:checked").length ? true : false);
            });
			   $("#inverse").click(function () {//反选
                $("#listForm :checkbox").each(function () {
                    $(this).attr("checked", !$(this).attr("checked"));
                });
            });
        });
		

function delPost(type){

	if(!confirm("确认删除数据吗? 不可恢复...")) return false;
	
		var n = $("input:checked").length;
		if (n == 0) {
			alert("未选择数据,无法操作！");
			return false;
		}

		$.post(	"../Public/sysLogDel.asp?operateType="+type, 
				$("#listForm").serialize(), 
				function(data,st){
						
						var resultArr=data.split("|");
						if(resultArr[0]=="1"){
							alert(resultArr[1]);
							window.location.reload();
						}
						else{
							alert(resultArr[1]);
						}	
				});
		

		
}
    </script>


</head>

<body >
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="30"><img src="../Images/tab_03.gif" width="15" height="30" /></td>
        <td width="550" background="../Images/tab_05.gif" class="tdStyle"><img src="../Images/311.gif" width="16" height="16" /> <span class="tbTitle">日志列表</span></td>
        <td background="../Images/tab_05.gif" class="tdStyle tdNavMn"><a href="javascript:void(0)" onClick="window.location.reload();"><img src="../Images/refresh.gif" width="16" height="16" /></a>
          <input id="checkAll" type="checkbox" name="checkbox62" value="checkbox" onBlur="selectAll(this);" />
          全选
          <input id="inverse" type="checkbox" name="inverse" value="checkbox" />
          反选 <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="delPost('0');">删除选中</a></font> <img src="../Images/083.gif" width="14" height="14" /><font><a href="javascript:void(0)" onClick="delPost('1');">删除所有记录</a></font></td>
        <td width="14"><img src="../Images/tab_07.gif" width="14" height="30" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9" background="../Images/tab_12.gif">&nbsp;</td>
        <td bgcolor="#f3ffe3"><form id="listForm" name="listForm" method="post"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#c0de98" >
          <tr>
            <td width="7%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2 STYLE1">选择</div></td>
            <td width="12%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2 STYLE1">时间</div></td>
            <td width="13%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2 STYLE1">操作IP</div></td>
            <td width="18%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2 STYLE1">操作页面</div></td>
            <td width="25%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2 STYLE1">操作功能说明</div></td>
            <td width="17%" height="26" background="../Images/tab_14.gif" class="STYLE1"><div align="center" class="STYLE2">(用户编号)操作用户</div></td>
            </tr>
          

<%
Dim ors
Set ors=new Cls_ShowoPage	'创建对象
With ors
	.Conn=conn			'得到数据库连接对象
	.DbType="AC"
	'数据库类型,AC为access,MSSQL为sqlserver2000,MSSQL_SP为存储过程版,MYSQL为mysql,PGSQL为PostGreSql
	.RecType=0
	'取记录总数方法(0执行count,1自写sql语句取,2固定值)
	.RecSql=0
	'如果RecType＝1则=取记录sql语句，如果是2=数值，等于0=""
	.RecTerm=2
	'取从记录条件是否有变化(0无变化,1有变化,2不设置cookies也就是及时统计，适用于搜索时候)
	.CookieName=""	'如果RecTerm＝2,cookiesname="",否则写cookiesname
	.Order=0			'排序(0顺序1降序),注意要和下面sql里面主键ID的排序对应
	.PageSize=20	'每页记录条数
	.JsUrl="../Inc/"			'showo_page.js的路径
	.Sql="[OPERATE_TIME],[IP],[OPERATE_PAGE],[OPERATE_CON],[User_NAME]$[T_LOG]$1=1 "&findKey&"$OPERATE_TIME desc$OPERATE_TIME" '字段,表,条件(不需要where),排序(不需要需要ORDER BY),主键ID
End With

iRecCount=ors.RecCount()'记录总数
iRs=ors.ResultSet()		'返回ResultSet
If  iRecCount<1 Then%>   
          <tr>
            <td height="18" bgcolor="#FFFFFF"><div align="center" class="STYLE1">[没有查到记录！]</span></div></td>
          </tr>

<%
Else     
    For i=0 To Ubound(iRs,2)
	

	%>    
          <tr>
            <td height="18" bgcolor="#FFFFFF"><div align="center" class="STYLE1">
              <input name="subBox" type="checkbox" class="STYLE2" value="<%=iRs(0,i)%>" />
            </div></td>
            <td height="18" bgcolor="#FFFFFF" class="STYLE2"><div align="center" class="STYLE2 STYLE1"><%=Format_Time(FromUnixTime(iRs(0,i),+8),1)%></div></td>
            <td height="18" bgcolor="#FFFFFF"><div align="center" class="STYLE2 STYLE1"><%=iRs(1,i)%></div></td>
            <td height="18" bgcolor="#FFFFFF"><div align="center" class="STYLE2 STYLE1"><%=iRs(2,i)%></div></td>
            <td height="18" bgcolor="#FFFFFF"><div align="center" class="STYLE2 STYLE1"><%=iRs(3,i)%></div></td>
            <td height="18" bgcolor="#FFFFFF"><div align="center"><a href="javascript:void(0)"><%=iRs(4,i)%></a></div></td>
            </tr>
       
    <%
    Next	
End If
%>
        </table></form></td>
        <td width="9" background="../Images/tab_16.gif">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="29"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15" height="29"><img src="../Images/tab_20.gif" width="15" height="29" /></td>
        <td background="../Images/tab_21.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td  height="29" id="bNavStr"></td>
            </tr>
        </table></td>
        <td width="14"><img src="../Images/tab_22.gif" width="14" height="29" /></td>
      </tr>
    </table></td>
  </tr>
</table>
<!--本页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒-->
</body>
</html>
<%ors.ShowPage()%>
<%
iRs=NULL
ors=NULL
Set ors=NoThing
%>