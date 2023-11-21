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
	background-color: #6699CC;
	background-image: url(img/top-bg.gif);
}
.top_style2 {
	font-size: 10pt;
	color: #FFFFFF;
	line-height:14pt;
	font-family: Arial, Helvetica, sans-serif;
}
-->
</style></head>

<body>
<span class="top_style2">一卡通系統<br/>
<img src="img/manager_user.gif" width="16" height="16" /><%=session("username_login")%> 歡迎您使用</span>
<%
NewUser=session("admin_id")
'*****開始計算線上人數及名單*****
Application.Lock
'*****如果你剛進入，則Application(NewUser & "LastAccess")應該為Empty
If Application(NewUser & "LastAccess") = Empty Then
'*****Application("TotalUsers")記錄線上總人數*****
'*****Application("OnLineUser")(I)記錄線上所有人的姓名之陣列*****?
'*****如果目前線上沒有人則Application("TotalUsers") = 0*****
  If Application("TotalUsers") = Empty Then Application("TotalUsers") = 0
  Redim Temp(Application("TotalUsers") + 1)
'*****把線上名單轉移到Temp()內，再把新進入者加到最後*****
No=0
  For I = 0 To Application("TotalUsers") - 1
    Item = Application("OnLineUser")(I)
    If Item<>NewUser then
      Temp(No) = Item
      No=No+1
    Else
      Application(Item & "LastAccess") = Empty
    End If
  Next
  Temp(No) = NewUser
'*****因為陣列由0開始，所以真正人數為No+1*****
  Application("TotalUsers") = No+1
  Redim Preserve Temp(Application("TotalUsers"))
'*****把線上名單存回Application("OnLineUser")，才能讓所有人看到*****
  Application("OnLineUser") = Temp
End If
'*****下面的程式需要些技巧，
'*****每個訪客記錄一個時間於Application(訪客的姓名 & "LastAccess")內，
'*****如果一直在線上則每更新一次就會重新記錄一次，
'*****與現在的時間相減如果大於60秒，則表示離線了，
'*****所有的記錄都重新記錄，如人數減一．．等。*****
Application(Session("admin_id") & "LastAccess") = Timer

If RefreshTime < 10 Then RefreshTime = 10
IdleTime = RefreshTime * 2

ReDim Temp(Application("TotalUsers"))
No = 0
For I = 0 To Application("TotalUsers") - 1
  Item = Application("OnLineUser")(I)
  If Abs(Application(Item & "LastAccess") - Timer) < IdleTime Then
    Temp(No) = Item
    No = No + 1
  Else
    Application(Item & "LastAccess") = Empty   
  End If
Next

If No <> Application("TotalUsers") Then
  Redim Preserve Temp(No)
  Application("OnLineUser") = Temp
  Application("TotalUsers") = No
End If
'*****Application("OnLineUser")記錄線上所有人的姓名。******
'*****Application("TotalUsers")記錄線上總人數。******
OnLineUser = Application("OnLineUser")
TotalUsers = Application("TotalUsers")

Application.UnLock
%>
</body>

</html>
