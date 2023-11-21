<%
session.codepage="65001"
response.charset = "utf-8"
Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%>
<%
'------------------------------------------------------------------------
'防止SQL injection
'------------------------------------------------------------------------
Function CheckSql()
    Dim sql_injdata  
    SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" 
    SQL_inj = split(SQL_Injdata,"|") 
    If Request.QueryString<>"" Then 
        For Each SQL_Get In Request.QueryString 
            For SQL_Data=0 To Ubound(SQL_inj) 
                if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_DATA))>0 Then 
                    Response.Write "<Script Language='javascript'>{alert('Error');history.back(-1)}</Script>" 
                    Response.end 
                end if 
            next 
        Next 
    End If
    If Request.Form<>"" Then 
        For Each Sql_Post In Request.Form 
            For SQL_Data=0 To Ubound(SQL_inj) 
                if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then 
                    Response.Write "<Script Language='javascript'>{alert('Error');history.back(-1)}    </Script>" 
                    Response.end 
                end if 
            next 
        next 
    end if
End Function

'------------------------------------------------------------------------
'取得使用者IP位置
'------------------------------------------------------------------------
function getIp()
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 
getIp=userip
End function
'------------------------------------------------------------------------
'取得使用者IP位置  第二版
'------------------------------------------------------------------------
Public Function Readusip()
  Dim strIPAddr
  If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
      strIPAddr = Request.ServerVariables("REMOTE_ADDR")
  ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
      strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
  ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
      strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
  Else
      strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  End If
  Readusip = Trim(Mid(strIPAddr, 1, 30))
End Function

'------------------------------------------------------------------------
'比對傳送資料是否來自合法位置 如果是 true 則允許 false 則不允許
'但若 使用 loaction.href 則會出現 false 
'------------------------------------------------------------------------
Function chkFrom() 
    Dim server_v1,server_v2, server1, server2
    chkFrom=False 
    server1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    server2=Cstr(Request.ServerVariables("SERVER_NAME"))
    If Mid(server1,8,len(server2))=server2 Then chkFrom=True 
End Function

'------------------------------------------------------------------------
'偷懶使用 
'------------------------------------------------------------------------
Function w(x) 
    response.write x
End Function

'------------------------------------------------------------------------
'防止SQL injection
'剔除錯誤字元
'------------------------------------------------------------------------
Function CheckStr(byVal ChkStr)
    Dim Str:Str=ChkStr
    Str=Trim(Str)
    If IsNull(Str) Then
        CheckStr = ""
        Exit Function 
    End If
    Dim re
    Set re=new RegExp
    re.IgnoreCase =True
    re.Global=True
    re.Pattern="(\r\n){3,}"
    Str=re.Replace(Str,"$1$1$1")
    Set re=Nothing
	Str = LCase(Str)
    Str = Replace(Str,"'","''")
    Str = Replace(Str, "select", "ｓelect")
    Str = Replace(Str, "join", "ｊoin")
    Str = Replace(Str, "union", "ｕnion")
    Str = Replace(Str, "where", "ｗhere")
    Str = Replace(Str, "insert", "ｉnsert")
    Str = Replace(Str, "delete", "ｄelete")
    Str = Replace(Str, "update", "ｕpdate")
    Str = Replace(Str, "like", "ｌike")
    Str = Replace(Str, "drop", "ｄrop")
    Str = Replace(Str, "create", "ｃreate")
    Str = Replace(Str, "modify", "ｍodify")
    Str = Replace(Str, "rename", "ｒename")
    Str = Replace(Str, "alter", "ａlter")
    Str = Replace(Str, "cast", "ｃast")
	Str = Replace(Str, "eval", "ｅval")
	Str = Replace(Str, "master", "ｍaster")
	Str = Replace(Str, "dbo", "ｄbo")
	
	Str = Replace(Str, "exec", "ｅxec")
	Str = Replace(Str, "create", "ｃreate")
	Str = Replace(Str, "count", "ｃount")
	Str = Replace(Str, "char", "ｃhar")
	Str = Replace(Str, "nchar", "ｎchar")
	Str = Replace(Str, "exists", "ｅxists")
	Str = Replace(Str, "exis", "ｅxis")
	Str = Replace(Str, "script", "ｓcript")
	Str = Replace(Str, "object", "ｏbject")
	Str = Replace(Str, "applet", "ａpplet")
	Str = Replace(Str, "Chr(0)", "")
	Str = Replace(Str, "Chr(13)", "<br>")
	Str = Replace(Str, "<", "")
	Str = Replace(Str, ">", "")
	Str = Replace(Str, "Chr(32)", "&nbsp;")
	Str = Replace(Str, "Chr(9)", "&nbsp;&nbsp;&nbsp;&nbsp;")
	Str = Replace(Str, "Chr(34)", "")
	Str = Replace(Str, "Chr(39)", "&#39;")
	Str = Replace(Str, "Chr(10)", "<br>")
	
    CheckStr=Str
End Function

'------------------------------------------------------------------------
'轉換資料型別 
'Value_ 輸入資料
'Type_  1->數值 2->文字 3->布林
'------------------------------------------------------------------------
function FunSwitch(Value_,Type_)
on error resume next
select case Type_
  case 1
  if isnumeric(Value_) then
   if not isnull(Value_) then
    FunSwitch=clng(Value_) 
   else
    FunSwitch=0
   end if
  else
   FunSwitch=0
  end if
  case 2
   if not isnull(Value_) then
    FunSwitch=cstr(Value_)
   else
    FunSwitch=""
   end if
  case 3
   if not isnull(Value_) and isnumeric(Value_) then
    FunSwitch=cbool(Value_)
   else
    FunSwitch=false
   end if
  case else
   w "Error:TypeError_1"
   
end select
if err.number<>0 then
w "Error:TypeError_2"

err.clear
end if
end function

'------------------------------------------------------------------------
'以下兩個為 文字與Unicode互轉林
'------------------------------------------------------------------------
Function Chr2Unicode(byval str)
	Dim st, t, i
	For i = 1 To Len(str)
		t = Hex(AscW(Mid(str, i, 1)))
		If (Len(t) < 4) Then
			while (Len(t) < 4)
			t = "0" & t
			Wend
		End If
		t = Mid(t, 3) & Left(t, 2)
		st = st & t
	Next
	Chr2Unicode = st
End Function
Function Unicode2Chr(byval str)
	Dim st, t, i
	For i = 1 To Len(str)/4
		t = Mid(str, 4*i-3, 4)
		t = Mid(t, 3) & Left(t, 2)
		t = ChrW(Hex2Dec(t))
		st = st & t
	Next
	Unicode2Chr = st
End Function
'------------------------------------------------------------------------
'會員密碼加密參數
'------------------------------------------------------------------------
Function PWDEncode(strPWD)

    Dim strAllValid 
    Dim strEnCoded  
    Dim chInPWD 
    Dim TempCode 
    Dim i 
    
    chInPWD = ""
    strEnCoded = ""
    strAllValid = "AB2abnoGRSTKLMN89OPQefghtuvwUVW34EFYZXcdHIijqrsklmCDz01pJxy567"
    
    For i = 1 To Len(strPWD)
       
       chInPWD = Mid(strPWD, i, 1)
       TempCode = InStr(strAllValid, chInPWD)
       
       If TempCode < 0 Then
          PWDEncode = "-1"
          Exit Function
       Else
          strEnCoded = strEnCoded & CStr(TempCode)
       End If
    Next 
    
    PWDEncode = strEnCoded
End Function
'------------------------------------------------------------------------
'補零函數
'------------------------------------------------------------------------
Function AddZero(x)
	if not isNull(x) then
		if len(x)=1 then
			AddZero="0"&x
		else
			AddZero=x
		end if
		
	end if
End Function
%>
<%
'基本定義參數
clientIP=Readusip()
myLOCAL_ADDR=Request.servervariables("LOCAL_ADDR ")		'
myPATH_INFO=Request.servervariables("PATH_INFO")		'執行路徑 如果有虛擬目錄時要用此參考
myREQUEST_METHOD=Request.servervariables("REQUEST_METHOD")		'存取方式 是 post 或是 get
mySCRIPT_NAME=Request.servervariables("SCRIPT_NAME")		'執行路徑
mySERVER_NAME=Request.servervariables("SERVER_NAME")		'主機名稱
myCLIENT_LANG=Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")	'客戶端的語言檔案

TimeNow=year(now)&"/"&month(now)&"/"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)
DateNow=year(now)&"/"&month(now)&"/"&day(now)

%>
<%
'---------------亂碼變數 全大寫
function RndCode(Num)
 Str="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
 StrLen=len(Str)
 for i=1 to Num
  Randomize
  RndNum=int(rnd(StrLen)*StrLen)+1
  RndStr=RndStr&mid(Str,RndNum,1)
 next
 RndCode=RndStr
end function
%>
<%
'---------------亂碼變數 全小寫
function RndCodelcase(Num)
 Str="0123456789abcdefghijklmnopqrstuvwxyz"
 StrLen=len(Str)
 for i=1 to Num
  Randomize
  RndNum=int(rnd(StrLen)*StrLen)+1
  RndStr=RndStr&mid(Str,RndNum,1)
 next
 RndCodelcase=RndStr
end function
%>
<%
'--------------寄信-------------
Function SendMail(strSendID,strTitle,strContent,strMail)
   Set mailreg=Server.CreateObject("CDONTS.NewMail")
   mailreg.To=strMail
   mailreg.From="<1@1.com.tw>"
   mailreg.Subject=strTitle
   mailreg.Body = strContent
   mailreg.MailFormat = 0
   mailreg.BodyFormat = 0
   mailreg.Send 
   SET mailreg=nothing 
End Function
%>
<%
'Connect to DB (SQL SERVER)
'''set adoConn = server.createobject("ADODB.Connection")
'''adoConn.Open "Driver={sql server};server=(local);UID=sa;PWD=portal12345;Database=Publish"

'Connect to DB (Access)
'''''dbPath = "data.mdb"
'''''Set adoConn = Server.CreateObject("Adodb.Connection")
'conn.ConnectionString="Driver={Microsoft Access Driver (*.mdb)};Dbq=E:\inetpub\vhosts\windows.software-mate.com\httpdocs\demo\datas\"&dbPath
'''''adoConn.ConnectionString="Driver={Microsoft Access Driver (*.mdb)};Dbq=E:\inetpub\vhosts\windows.software-mate.com\httpdocs\html6\"&dbPath
'''''adoConn.Open

'dim conn 
set adoConn = server.createobject("adodb.connection") 
adoConn.open = "provider=microsoft.jet.oledb.4.0;"&"data source=" & server.mappath("data.mdb") 
%>
<%
'------------------------------
'讀取全形半形字
'------------------------------
Function Readtxt(strline,J,X)
   Y=0
   For i=J to J+X-1
     strword=ASC(mid(strline,i,1))     
     if strword < 128 then
         Y=Y+1
     else 
         Y=Y+2
     end if
     if Y >= X then
        Readtxt = mid(strline,J,i)        
        exit for
     end if
   Next
End Function
%> 
