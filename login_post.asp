<%@ codepage=65001%>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%>
<!--#include file="function.asp" -->
<%
username=CheckStr(request("username"))
password=CheckStr(request("password"))

  set rs = Server.CreateObject("ADODB.RecordSet")
  sql="select A_id,Group_id,logintime,A_right from useraccount where LoginID='"&username&"' and LoginPwd='"&password&"' "
  'w sql
  rs.open sql, adoConn, 2,3
  if not rs.eof then
 
    sqlstr_list = "SELECT s_type FROM sys_config where s_id=1"
	SET rs_list = Server.CreateObject("ADODB.Recordset")
	rs_list.OPEN sqlstr_list, adoConn, 3,3
		s_type=FunSwitch(rs_list("s_type"),1)
	rs_list.CLOSE

	
	''''判斷 是否有 已經在線上了 ''''''''''
	'*****開始計算線上人數及名單*****
    if s_type=1 then 		 
        Application.Lock
	OnLineUserAAA = Application("OnLineUser")
	TotalUsersAAA = Application("TotalUsers")
	Application.UnLock

	if TotalUsersAAA>0 then
		For I = 0 To TotalUsersAAA - 1  
			if FunSwitch(OnLineUserAAA(I),2)=FunSwitch(rs("A_id"),2) then
				'response.Redirect "index.asp?t=a"
				response.write OnLineUserAAA(I)&"<br>"
				response.write rs("A_id")&"<br>"
				Item = Application("OnLineUser")(I)
				response.write Abs(Application(Item & "LastAccess") - Timer)
				If Abs(Application(Item & "LastAccess") - Timer) < 60 Then
					response.Redirect "index.asp?t=a"
					response.end
				end if
			end if
		Next
	end if
	
    end if		 
	
	'response.end
	''''判斷 是否有 已經在線上了 ''''''''''
	session("admin_id")=rs("A_id")
	session("Group_id")=rs("Group_id")
	session("username_login")=username
	session("logintime")=rs("logintime")
	session("AAA_right")=rs("A_right")
	
	'''' update logintime
	sql = "update useraccount set logintime='"&now()&"' where A_id="&session("admin_id")&""
        adoconn.execute(sql)
	response.Redirect "menu.asp"
	'w "ok"
  else
  	response.Redirect "index.asp"
	'w "error"
  end if
  rs.close
  adoConn.close
%>

