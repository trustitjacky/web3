<%@ codepage=65001%>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- Mirrored from www.gatsby.hk/ by HTTrack Website Copier/3.x [XR&CO'2007], Tue, 15 May 2007 10:20:52 GMT -->
<!-- Added by HTTrack --><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
.style3 {
	color: #FF0000;
	font-size: 9pt;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
</head>
<style type="text/css">
<!--
.style2 {font-size: 9pt}
-->
</style>
<TABLE align="center" width='554 cellspacing='0' cellpadding='0' border='0'>
<tr bgcolor="#efb54a"> 
   <td height="20" bgcolor="#FFFFFF"><font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><img src="img/functionmenu.gif" width="554" height="62"></font></td>
</tr></table>
<TABLE align="center" width='554 cellspacing='2' cellpadding='0' border='1'>
<TR><TD valign='top' align='center'>
<a href='mainsetX.asp'>一卡通通借</a></TD>
<TD valign='top' align='center'>
<a href='mainset.asp'>一卡通通還</a></TD></TR>
</TABLE><br>
</html>

