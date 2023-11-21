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
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<!--#include file="function.asp" -->
<style type="text/css">
<!--
.style2 {font-size: 9pt; }
.style3 {font-size: 9pt; color: #333333; }
-->
</style>
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
	keywords=FunSwitch(CheckStr(request("keywords")),2)
	terms=FunSwitch(CheckStr(request("terms")),2)
%>

<fieldset>
<LEGEND><span class="style3">權責單位管理</span></LEGEND>
<form action="group.asp" method="post" class="style3">搜尋關鍵字
	  <input name="keywords" type="text" id="keywords" value="<%=keywords%>" size="8" maxlength="10" />
      <select name="terms" id="terms">
        <option value="A_no" <% if terms="A_no" then w "selected" end if %>>代碼</option>
        <option value="UserName" <% if terms="UserName" then w "selected" end if %>>中文名稱</option> 
        <option value="UserNameE" <% if terms="UserNameE" then w "selected" end if %>>英文名稱</option>
      </select>
      <input type="submit" name="Submit2" value="搜尋" />
</form>
      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td colspan="2" bgcolor="#FFFFFF"><div align="right">
            <input name="Submit" type="submit" onclick="MM_goToURL('self','group-upd.asp?op=add');return document.MM_returnValue" value="新增權責單位資料" />
        </div></td>
      </tr>
    </table>
      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3">
        <tr>
          <td width="12%" bgcolor="D0DFFF"><div align="center" class="style3">代碼</div></td>
          <td width="11%" bgcolor="D0DFFF"><div align="center" class="style3">
            <div align="center">中文名稱</div>
          </div></td>
          <td width="14%" bgcolor="D0DFFF"><div align="center" class="style3">
            <div align="center">英文名稱</div>
          </div></td>
          <td width="9%" bgcolor="D0DFFF"><div align="center" class="style3">
            <div align="center">連絡人</div>
          </div></td>
          <td width="14%" bgcolor="D0DFFF"><div align="center" class="style3">地址</div></td>
          <td width="16%" bgcolor="D0DFFF"><div align="center" class="style3">電話 </div></td>
          <td width="12%" bgcolor="D0DFFF"><div align="center" class="style3">EMAIL</div></td>
          <td width="12%" bgcolor="D0DFFF"><div align="center" class="style3">管理</div></td>
        </tr>
  <%
  page=FunSwitch(CheckStr(request("page")),1)
  sqlstr_list = "SELECT A_id,A_no,UserNameE,UserName,TEL,Address,contacter,email FROM department "
  if len(keywords)>0 then
  	sqlstr_list=sqlstr_list&" where "&terms&" like '%"&keywords&"%' " 
  end if
  sqlstr_list=sqlstr_list&" order by A_id desc"
	'w sqlstr_list&"<br>"
	SET rs_list = Server.CreateObject("ADODB.Recordset")
	rs_list.OPEN sqlstr_list, adoConn, 3,3
	'response.write rs_list.recordcount
     IF page = 0 THEN Page = 1 END IF

     rs_list.PageSize = 10           ' 設定每頁顯示 10 筆
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
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("A_no")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("UserName")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("UserNameE")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("contacter")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Address")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Tel")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("email")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3">
          <input name="btn_upd<%=rs_list("A_id")%>" type="button" onclick="MM_goToURL('self','group-upd.asp?id=<%=rs_list("A_id")%>&op=upd');return document.MM_returnValue;" value="修" />
		  
            <input type="button" name="btn_del<%=rs_list("A_id")%>" value="刪" onclick="if(window.confirm('確定刪除?')){MM_goToURL('self','group-db.asp?id=<%=rs_list("A_id")%>&op=del');return document.MM_returnValue;}" />
          </span></td>
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
          <td colspan="2" bgcolor="#FFFFFF" class="style3"><%  CALL page_change_10("group.asp",keywords,terms) %></td>
        </tr>
    </table>
	<%
end if
FUNCTION page_change_10(pg_name,keywords,terms)

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


      IF X-1=0 or rs_list.pagecount<=10  THEN
        RESPONSE.WRITE ""
      ELSE
        RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&page="&jpage-10&"&epage="&epage-10&"&keywords="&Server.UrlEncode(keywords)&"&terms="&terms&"' class=""c02""><font size='-1'> 上10頁</a>‧"
      END IF
      next_page=X+1
     
      '----------------------------------
      IF rs_list.pagecount<=10 THEN
         RESPONSE.WRITE ""
      ELSEIF cint(epage)>rs_list.pagecount THEN
         RESPONSE.WRITE ""
      ELSE
            RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&page="&jpage+10&"&epage="&epage+10&"&keywords="&Server.UrlEncode(keywords)&"&terms="&terms&"' class=""c02"">下10頁 </a>&nbsp;&nbsp;&nbsp;&nbsp;"
      END IF

     '----------------------------------

      RESPONSE.WRITE "目前頁次: "&X&"　[ 頁次:"

      FOR J=cint(jpage) to cint(epage)
          IF J<=rs_list.pagecount THEN
           RESPONSE.WRITE " <a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&epage="&epage&"&page="&cint(J)&"&keywords="&Server.UrlEncode(keywords)&"&terms="&terms&"' class=""c02"">"&J&"</A>"
          END IF
      NEXT

      RESPONSE.WRITE " ]"
      'RESPONSE.WRITE "<br><br><br>"
      RESPONSE.WRITE "</td>"
      RESPONSE.WRITE "</tr>"
      RESPONSE.WRITE "</table>"
    END FUNCTION
%>
</fieldset>
<%
sqlLog="insert into logss(UserName,logdate,url)values('"&session("Admin_id")&"','"&TimeNow&"','"&mySCRIPT_NAME&Request.ServerVariables("QUERY_STRING")&"')"
adoconn.execute(sqlLog)
%>