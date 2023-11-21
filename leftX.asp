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
body {
	background-color: #EFF3F8;
}
.LEGEND-title {
	font-size: 9pt;
	color: #000033;
	font-family: Arial, Helvetica, sans-serif;
}
a { text-decoration: none ;
	font-family: "新細明體"; 
	font-size: 9pt; 
	color:#663300;
}
-->
</style></head>
<%
SET rs = Server.CreateObject("ADODB.Recordset")
sql="select A_right,Group_id from useraccount where LoginID='"&session("username_login")&"'"
  rs.open sql, adoConn, 2,3
  if not rs.eof then
     rA_right=rs("A_right")
     rGroup_id=rs("Group_id")
  end if 
  RS.CLOSE
%>
<body>
<fieldset>  
<LEGEND class="LEGEND-title">&nbsp;<img src="img/editor.gif" width="16" height="16" />管理選項&nbsp;</LEGEND>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
  <% if rA_right="1" then %>  
  <tr>
    <td>        
      <table width="99%" border="0" cellpadding="1" id="tbl1">            
	  <tr><td><a href="group.asp" target="mainFrame" title="權責單位管理">權責單位管理</a></td></tr>	
          <tr><td><a href="user.asp" target="mainFrame" title="使用者帳號管理">使用者帳號管理</a></td></tr>		
          <% if session("username_login")="setty" or session("username_login")="admin" or session("username_login")="yummy" then %> 
           <tr><td><a href="log.asp" target="mainFrame">LOG查詢</a></td></tr>
          <% end if %>
      </table>
    </td>
  </tr><% end if %>  
  <% if rGroup_id=1 then %> 
  <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">    
        <tr>
          <td><a href="disbutionX.asp" target="mainFrame" title="集中地點收">一卡通集中地點收作業
          </a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">    
        <tr>
          <td><a href="disbution2X.asp" target="mainFrame" title="集中地配送">一卡通集中地配送作業
          </a></td>
        </tr>
      </table>
    </td>
  </tr>
  <% end if %> 


 <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">    
        <tr>
          <td><a href="transationX.asp" target="mainFrame" title="移送資料維護">一卡通通借移送作業
          </a></td>
        </tr>
      </table>
    </td>
  </tr>
       
  <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">    
        <tr>
          <td><a href="circulationX.asp" target="mainFrame" title="通借資料維護">一卡通通借作業
          </a></td>
        </tr>
      </table>
    </td>
  </tr>

  <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">
        <tr>
          <td><a href="rf2X.asp" target="mainFrame" title="移送狀況表">一卡通通借移送狀況表</a></td>
        </tr>
      </table>
    </td>

  <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">
        <tr>
          <td><a href="rf3X.asp" target="mainFrame" title="通借狀況報表">一卡通通借狀況表</a></td>
        </tr>
      </table>
    </td>
  </tr>
<% if session("username_login")="setty" or session("username_login")="admin" or session("username_login")="yummy"  then %> 
 <tr>
    <td>
      <table width="99%" border="0" cellpadding="2">
        <tr>
          <td><a href="rf1X.asp" target="mainFrame" title="通借狀況Excel報表">一卡通通借狀況Excel表</a></td>
        </tr>
      </table>
    </td>
  </tr>
<% end if %>
  <tr>
    <td><a href="menu.asp" target="_parent"><strong><img src="img/20070509090206748.gif" width="9" height="8" border="0" />選單</strong></a></td>
  </tr></br>
  <tr>
    <td><a href="logout.asp" target="_parent"><strong><img src="img/20070509090206748.gif" width="9" height="8" border="0" />登出</strong></a></td>
  </tr></br>
  <tr>
    <td><a href="使用者手冊.pdf" target="mainFrame"><strong><img src="img/20070509090206748.gif" width="9" height="8" border="0" />使用者手冊</strong></a></td>
  </tr>  
</table>
</fieldset>
</body>
</html>
