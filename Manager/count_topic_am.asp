<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="CheckManager.asp"-->
<HTML>
   <HEAD>
      <TITLE>近年查询考试人数</TITLE> 
      
   <link href="../Css/main.css" rel="stylesheet" type="text/css">
   <script language="javascript" src="../Inc/jquery-1.7.1.min.js"></script>
   </HEAD>
   <!-- #INCLUDE FILE="../Inc/FusionCharts.asp" -->
<BODY>
<!--#include file="test_id_check.asp"-->
<%
   'Create an XML data document in a string variable
   'sql="select ID,TNAME from T_TEST_TOPIC_RELATION AS TTR Left Join T_TEST_TOPIC"
   Dim strXML
   strXML = ""
   strXML = strXML & "<graph caption='按考试科目统计平均分' xAxisName='科目' yAxisName='am' decimalPrecision='0' formatNumberScale='0' >"
   strXML = strXML & "<set name='道德修养' value='52' color='AFD8F8' />"
   strXML = strXML & "<set name='法规简论' value='63' color='F6BD0F' />"
   strXML = strXML & "<set name='高等教育心理学' value='74' color='8BBA00' />"
   strXML = strXML & "<set name='高等教育学' value='60' color='FF8E46' />"
   strXML = strXML & "</graph>"

   'Create the chart - Column 3D Chart with data from strXML variable using dataXML method
   Call renderChartHTML("../Charts/FCF_Column3D.swf", "", strXML, "myNext", 800, 500)
%>
</BODY>
</HTML>
