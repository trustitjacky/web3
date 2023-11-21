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
		AUserNameE=request("AUserNameE")
		AUserName=FunSwitch(request("AUserName"),2)
		AGroup_id=request("AGroup_id")
		ATel=request("ATel")
		Aemail=FunSwitch(request("Aemail"),2)
		ALoginID=FunSwitch(request("ALoginID"),2)
		ALoginPwd=FunSwitch(request("ALoginPwd"),2)
		A_right=request("A_right")
%>
<% if op<>"del" then %>
<%
if len(AUserName)=0 then
%>
<script language="javascript">
		alert ("姓名不可空白");
		window.history.go(-1);
</script>
<%
response.end
end if
%>
<%
if len(ALoginID)=0 then
%>
<script language="javascript">
		alert ("帳號不可空白");
		window.history.go(-1);
</script>
<%
response.end
end if
%>
<%
if len(ALoginPwd)=0 then
%>
<script language="javascript">
		alert ("密碼不可空白");
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
		''''select 帳號不可以重複
		idchk="select A_id from useraccount where LoginID='"&ALoginID&"'"
		SET rs_list = Server.CreateObject("ADODB.Recordset")
		rs_list.OPEN idchk, adoConn, 3,3
		if not rs_list.eof then
			%>
			<script>
			alert('此編號重複,不可使用');
			window.history.go(-1);
			</script>
			<%
		end if
		rs_list.CLOSE
		sql="insert into useraccount(UserNameE,UserName,Group_id,Tel,email,LoginID,LoginPwd,A_right)values('"&AUserNameE&"','"&AUserName&"','"&AGroup_id&"','"&ATel&"','"&Aemail&"','"&ALoginID&"','"&ALoginPwd&"','"&A_right&"')"
end if

if op="upd" then
		sql = "update useraccount set "
		
		sql = sql & "UserNameE = '" & AUserNameE & "', "
		sql = sql & "UserName = '" & AUserName & "', "
		sql = sql & "Group_id = '" & AGroup_id & "',"
		sql = sql & "Tel = '" & ATel & "', "
		sql = sql & "email = '" & Aemail & "', "
		sql = sql & "LoginID = '" & ALoginID & "', "
		sql = sql & "LoginPwd = '" & ALoginPwd & "', "
		sql = sql & "A_right = " & A_right & " "

		sql = sql & "where A_id = " & id & ""
end if

if op="del" then
	sql = "delete from useraccount where A_id = " & id & ""
end if
'w sql
	adoconn.execute(sql)
	response.Redirect "user.asp"
end if
%>

