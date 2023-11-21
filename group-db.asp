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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="function.asp" -->
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
</head>
<%
SET rs = Server.CreateObject("ADODB.Recordset")
sql="select A_right,Group_id from useraccount where LoginID='"&session("username_login")&"'"
  rs.open sql, adoConn, 2,3
  if not rs.eof then
     rA_right=rs("A_right")
     rGroup_id=rs("Group_id")
  end if 
  RS.CLOSE
if rA_right="1" then
%>
<%
op=FunSwitch(request("op"),2)
id=FunSwitch(request("id"),1)

		UserName=request("UserName")
		UserNameE=FunSwitch(request("UserNameE"),2)
		TEL=request("TEL")
		Address=request("Address")
		contacter=FunSwitch(request("contacter"),2)
		A_no=FunSwitch(request("A_no"),2)
		email=FunSwitch(request("email"),2)
		
%>
<% if op<>"del" then %>
<%
if len(UserName)=0 then
%>
<script language="javascript">
		alert ("名稱不可空白");
		window.history.go(-1);
</script>
<%
response.end
end if
%>
<%
if len(contacter)=0 then
%>
<script language="javascript">
		alert ("聯絡人不可空白");
		window.history.go(-1);
</script>
<%
response.end
end if
%>
<%
if len(A_no)=0 then
%>
<script language="javascript">
		alert ("權責單位代碼不可空白");
		window.history.go(-1);
</script>
<%
response.end
end if
%>

<% end if %>
<%
'''''Database''''''''''''''''
if op="add" then
		sql="insert into department(UserName,UserNameE,TEL,Address,contacter,A_no,email)values('"&UserName&"','"&UserNameE&"','"&TEL&"','"&Address&"','"&contacter&"','"&A_no&"','"&email&"')"
end if

if op="upd" then
		sql = "update department set "
		sql = sql & "UserName = '" & UserName & "', "
		sql = sql & "UserNameE = '" & UserNameE & "', "
		sql = sql & "TEL = '" & TEL & "', "
		sql = sql & "Address = '" & Address & "',"
		sql = sql & "contacter = '" & contacter & "', "
		sql = sql & "A_no = '" & A_no & "', "
		sql = sql & "email = '" & email & "' "
		
		sql = sql & "where A_id = " & id & ""
end if

if op="del" then
	sql = "delete from department where A_id = " & id & ""
end if
'w sql
	adoconn.execute(sql)
	response.Redirect "group.asp"
%>
<% end if %>
