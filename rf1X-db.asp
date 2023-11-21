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
    Startdate=request("StartDate")
    Enddate=request("EndDate")
    IF len(StartDate) > 0 and len(EndDate)>0 THEN
        sql = "DELETE FROM recordX where Send_Date between #"&Startdate&"# and #"&Enddate&"#"
        adoconn.execute(sql)    
        response.Redirect "rf1X.asp"
    ELSE 
         response.write("<script>alert('日期區間未設定，請重設！');location.href='rf1X.asp';</script>")
    END IF  
end if 
%>