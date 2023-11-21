<%@ codepage=65001%>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%>
<!--#include file="function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
<script src="prototype.js" type="text/javascript"></script>
<script src="pre_logined.js" type="text/javascript"></script>
</head>
<frameset rows="59,*" cols="*" framespacing="2" frameborder="no" border="0" bordercolor="#ECE9D8">
  <frame src="top.asp" name="topFrame" scrolling="No" noresize="noresize" id="topFrame" title="topFrame" />
  <frameset cols="230,*" frameborder="no" border="0" framespacing="0">
    <frame src="leftX.asp" name="leftFrame" scrolling="auto" noresize="noresize" id="leftFrame" title="leftFrame" />
    <frame src="main.asp" name="mainFrame" id="mainFrame" title="mainFrame" />
  </frameset>
</frameset>
<noframes><body>
</body>
</noframes></html>
