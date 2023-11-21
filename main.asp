<%@ codepage=65001%>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%>
<!--#include file="function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
<style type="text/css">
<!--
.style2 {
	font-size: 10pt;
	font-family: Arial, Helvetica, sans-serif;
	color: #333333;
}
-->
</style>
</head>

<body>
<fieldset>
<LEGEND><span class="style2">Greet with your comming 歡迎光臨</span></LEGEND>
<!--table width="70%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#A8B9CE" bordercolorlight="#000000" bordercolordark="#FFFFFF"-->
<table width="70%" border="0" align="center" cellpadding="4">
  <tr>
    <td width="2%" rowspan="2" class="style2"><img src="img/g1.gif" width="80" height="75" /></td>
    <td width="98%" class="style2"><%=session("username_login")%> 歡迎您的使用</td>
  </tr>
  <tr>
    <td valign="top" class="style2">您上次登入時間為：<%=session("logintime")%></td>
  </tr>
</table>
</fieldset>
</body>

</html>
