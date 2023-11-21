<%@ codepage=65001%>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.AddHeader "Pragma", "No-Cache"
%><!--#include file="function.asp" --><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/calendar-win2k-cold-1.css" title="win2k-cold-1" />
  <script type="text/javascript" src="jscalendar-1.0/calendar.js"></script>
  <script type="text/javascript" src="jscalendar-1.0/lang/calendar-en.js"></script>
  <script type="text/javascript" src="jscalendar-1.0/calendar-setup.js"></script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
<style type="text/css">
<!--
.style2 {
	font-size: 10pt;
	font-family: Arial, Helvetica, sans-serif;
	color: #333333;
}
.style7 {font-size: 8pt; color: #999999; font-family: Arial, Helvetica, sans-serif; }
-->
<!--
.style3 {font-size: 9pt; color: #333333; }
-->
</style>
</head>
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>

<body>
<fieldset>
<LEGEND><span class="style2">LOG查詢</span></LEGEND>
<%
Function showBuyerName(x)
	x=FunSwitch(x,1)
	sqlstr_listaa = "SELECT LoginID FROM useraccount where A_id="&int(x)&""
	SET rs_listaa = Server.CreateObject("ADODB.Recordset")
	rs_listaa.OPEN sqlstr_listaa, adoConn, 3,3
		if not rs_listaa.eof then
		p_cname=rs_listaa("LoginID")
		else
		p_cname=""
		end if
	w p_cname
	rs_listaa.CLOSE
	set rs_listaa=nothing
END Function
%>
<br />
<%
'SET rs = Server.CreateObject("ADODB.Recordset")
'sql="select A_right,Group_id from useraccount where LoginID='"&session("username_login")&"'"
'  rs.open sql, adoConn, 2,3
'  if not rs.eof then
'     rA_right=rs("A_right")
'     rGroup_id=rs("Group_id")
'  end if 
'  RS.CLOSE
'if rA_right="1" then
if session("username_login")="setty" or session("username_login")="admin" or session("username_login")="yummy" then 
%>
<form id="form1" name="form1" method="post" action="log-db.asp">

       <tr>
          <span class="style7"><td valign="top" bgcolor="#FFFFFF"><div align="right">刪除</td>
          <input name="date1" type="text" id="date1" value="<% w date()-365 %>" />
		  <button type="button" id="trigger">...</button>前的記錄
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "date1",      
        ifFormat       :    "%Y/%m/%d",       
        showsTime      :    true,            
        button         :    "trigger",   
        singleClick    :    true,           
        step           :    1                
    });
</script>
<!--input name="submit" type="submit" onclick="MM_goToURL('self','log-db.asp');return document.MM_returnValue" value="刪除" /-->
<input name="submit" type="submit" value="刪除" />
	 </span></div></td>
        </tr>
</form>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3">

        <tr>
          <td bgcolor="#D0DFFF"><div align="center" class="style7">USER</div></td>
          <td bgcolor="#D0DFFF"><div align="center" class="style7">Log action </div></td>
          <td bgcolor="#D0DFFF"><div align="center" class="style7">Time</div></td>
        </tr>
  <%
  page=FunSwitch(CheckStr(request("page")),1)
  'seldate=DateAdd("m", -1 , now())
 ' sqlstr_list = "SELECT M_id,Mno,UserNameE,UserName,Occu,Tel,Email FROM member order by M_id desc"
  'sqlstr_list = "SELECT UserName,url,logdate FROM logss where logdate<'"&year(seldate)&"/"&month(seldate)&"/"&day(seldate)&" 00:00:00 ' order by logdate desc"
  sqlstr_list = "SELECT top 500 UserName,url,logdate FROM logss order by logdate desc"
	'w sqlstr_list
	SET rs_list = Server.CreateObject("ADODB.Recordset")
	rs_list.OPEN sqlstr_list, adoConn, 3,3
	'response.write rs_list.recordcount
     IF page = 0 THEN Page = 1 END IF

     rs_list.PageSize = 30           ' 設定每頁顯示 30 筆
	 'w rs_list.recordcount
'response.end
     IF Not rs_list.eof THEN          ' 有資料才執行 
        rs_list.AbsolutePage = page   ' 將資料錄移至 PAGE 頁
     END IF
  %>
  <%
	if not rs_list.eof then
	FOR SH=1 to rs_list.PageSize
   %>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style7">&nbsp;
          <% showBuyerName(rs_list("UserName"))%>
          </span></td>
          <td bgcolor="#FFFFFF"><span class="style7">&nbsp;<%=rs_list("url")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style7">&nbsp;<%=rs_list("logdate")%></span></td>
        </tr>
    <%
		rs_list.MOVENEXT
		IF rs_list.EOF THEN EXIT FOR
		Next
						
		end if
	%>
      </table>
      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td colspan="2" bgcolor="#FFFFFF" class="style3"><%  CALL page_change_10("log.asp",x) %></td>
        </tr>
      </table>
</fieldset><%
end if
FUNCTION page_change_10(pg_name,x)

      X=page
      epage=REQUEST("epage")
      rd=rs_list.RecordCount
      rp=rs_list.PageSize
      
      IF epage="" THEN
         epage=10
      ELSE
         epage=REQUEST("epage")
      END IF

      jpage=epage-9


      RESPONSE.WRITE "<table bgcolor='white' width='100%' border='0' cellspacing='0' cellpadding='0'>"
      RESPONSE.WRITE "  <tr>"
      RESPONSE.WRITE "  <td align='center' valign='bottom' class='page'>"

      'RESPONSE.WRITE jpage&"to"&epage&"to"&rs_list.pagecount


      IF X-1<10 or rs_list.pagecount<=10  THEN
        RESPONSE.WRITE ""
      ELSE
        RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&x="&x&"&page="&jpage-10&"&epage="&epage-10&"' class=""c02""><font size='-1'> 上10頁</a>‧"
      END IF
      next_page=X+1
     
      '----------------------------------
      IF rs_list.pagecount<=10 THEN
         RESPONSE.WRITE ""
      ELSEIF cint(epage)>rs_list.pagecount-1 THEN
         RESPONSE.WRITE ""
      ELSE
            RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&x="&x&"&page="&jpage+10&"&epage="&epage+10&"' class=""c02"">下10頁 </a>&nbsp;&nbsp;&nbsp;&nbsp;"
      END IF

     '----------------------------------

      RESPONSE.WRITE "目前頁次: "&X&"　[ 頁次:"

      FOR J=cint(jpage) to cint(epage)
          IF J<=rs_list.pagecount THEN
           RESPONSE.WRITE " <a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&x="&x&"&epage="&epage&"&page="&cint(J)&"' class=""c02"">"&J&"</A>"
          END IF
      NEXT

      RESPONSE.WRITE " ]"
      'RESPONSE.WRITE "<br><br><br>"
      RESPONSE.WRITE "</td>"
      RESPONSE.WRITE "</tr>"
      RESPONSE.WRITE "</table>"
    END FUNCTION
%>
</body>

</html>
