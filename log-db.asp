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
   
    date1=request("date1")

    'IF Month(date1) < 10 THEN
    '   intDate = Int(Year(date1) & 0 & Month(date1) & Day(date1))
    'ELSE
    '   intDate = Int(Year(date1) & Month(date1) & Day(date1))
    'END IF
'''''Database''''''''''''''''
        'sql = "DELETE FROM logss where int(logdate)<='"&intDate&"'"
        sql = "DELETE FROM logss where logdate<=#"&date1&"#"
        'sql = "DELETE FROM logss where datediff('d',logdate,"&date1&")>=0"
         adoconn.execute(sql)        
        response.Redirect "log.asp"
end if 
%>