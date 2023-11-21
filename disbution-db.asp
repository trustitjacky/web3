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
myPageSize=FunSwitch(CheckStr(request("myPageSize")),1)
kword=FunSwitch(CheckStr(request("kword")),2)
PID=request("PID")
Source_dep=FunSwitch(CheckStr(request("Source_dep")),2)
Target_dep=FunSwitch(CheckStr(request("Target_dep")),2)
chkStatus=request("chkStatus")
ordby=FunSwitch(request("ordby"),2)
ordby=replace(ordby,"%20"," ")
if myPageSize=0 then myPageSize=10 end if
if len(ordby)=0 then ordby="Send_Date desc" end if
page=FunSwitch(CheckStr(request("page")),1)
epage=FunSwitch(CheckStr(request("epage")),1)
search_e_title=FunSwitch(CheckStr(request(SESSION("search_e_title"))),2)

id=FunSwitch(request("id"),1)
'''''Database''''''''''''''''
        sql = "update publishes set "
	sql = sql & "Center_Date = '" & now() & "', "
	sql = sql & "Center_User = '" & session("username_login") & "', "
	sql = sql & "Status = 1 "
	sql = sql & "where P_id = " & id & ""
        adoconn.execute(sql)        
        'record
        sqlstr_list = "SELECT * FROM publishes where P_id = " & id & ""
        SET rs_list = Server.CreateObject("ADODB.Recordset")
        rs_list.OPEN sqlstr_list, adoConn, 3,3 
        if NOT rs_list.eof then
           rP_NO=rs_list("P_NO")
           rP_Name=rs_list("P_Name")
           rSource_id=rs_list("Source_id")
           rTarget_id=rs_list("Target_id")
           if len(rs_list("Update_User"))>0 then
              rUpdate_User=rs_list("Update_User")
           else
              rUpdate_User=" "
           end if 
           rSend_Date=rs_list("Send_Date")
            if len(rs_list("Recive_Date"))>0 then
              rRecive_Date=rs_list("Recive_Date")
           else
              rRecive_Date=" "
           end if 
           rStatus=1
           rCenter_Date=now()
           if len(rs_list("Deliver_Date"))>0 then
              rDeliver_Date=rs_list("Deliver_Date")
           else
              rDeliver_Date=" "
           end if 
           rCenter_User=session("username_login")
           if len(rs_list("Deliver_User"))>0 then
              rDeliver_User=rs_list("Deliver_User")
           else
              rDeliver_User=" "
           end if 
           sql_r = "insert into record(P_NO,P_Name,Source_id,Target_id,Update_User,Send_Date,Recive_Date,Status,Center_Date,Deliver_Date,Center_User,Deliver_User)values('"&rP_NO&"','"&rP_Name&"','"&rSource_id&"','"&rTarget_id&"','"&rUpdate_User&"','"&rSend_Date&"','"&rRecive_Date&"','"&rStatus&"','"&rCenter_Date&"','"&rDeliver_Date&"','"&rCenter_User&"','"&rDeliver_User&"')" 
           adoConn.execute(sql_r)  
        end if  
        rs_list.CLOSE
        response.Redirect "disbution.asp"&"?search_e_title="&SESSION("search_e_title")&"&kword="&Server.UrlEncode(kword)&"&PID="&PID&"&chkStatus="&chkStatus&"&ordby="&ordby&"&myPageSize="&myPageSize&"&page="&page&"&epage="&epage&"&Source_dep="&Server.UrlEncode(Source_dep)&"&Target_dep="&Server.UrlEncode(Target_dep)
%>

