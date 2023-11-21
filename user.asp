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
<STYLE TYPE="text/css">
thead .arrow {font-family: webdings; color: black; padding: 0; font-size: 10px;
height: 11px; width: 10px; overflow: hidden;
margin-bottom: 2; margin-top: -3; padding: 0; padding-top: 0; padding-bottom: 2;}
.style6 {font-size: 10pt; color: #333333; }
<!--
.style2 {font-size: 9pt; }
.style3 {font-size: 9pt; color: #333333; }
-->
</STYLE>
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<script src="prototype.js" type="text/javascript"></script>
<script src="pre_logined.js" type="text/javascript"></script>
<SCRIPT LANGUAGE="JavaScript">
var dom = (document.getElementsByTagName) ? true : false;
var ie5 = (document.getElementsByTagName && document.all) ? true : false;
var arrowUp, arrowDown;

if (ie5 || dom)
initSortTable();

function initSortTable() {
arrowUp = document.createElement("SPAN");
var tn = document.createTextNode("5");
arrowUp.appendChild(tn);
arrowUp.className = "arrow";

arrowDown = document.createElement("SPAN");
var tn = document.createTextNode("6");
arrowDown.appendChild(tn);
arrowDown.className = "arrow";
}



function sortTable(tableNode, nCol, bDesc, sType) {
var tBody = tableNode.tBodies[0];
var trs = tBody.rows;
var trl= trs.length;
var a = new Array();

for (var i = 0; i < trl; i++) {
a[i] = trs[i];
}

var start = new Date;
window.status = "Sorting data...";
a.sort(compareByColumn(nCol,bDesc,sType));
window.status = "Sorting data done";

for (var i = 0; i < trl; i++) {
tBody.appendChild(a[i]);
window.status = "Updating row " + (i + 1) + " of " + trl +
" (Time spent: " + (new Date - start) + "ms)";
}

// check for onsort
if (typeof tableNode.onsort == "string")
tableNode.onsort = new Function("", tableNode.onsort);
if (typeof tableNode.onsort == "function")
tableNode.onsort();
}

function CaseInsensitiveString(s) {
return String(s).toUpperCase();
}

function parseDate(s) {
return Date.parse(s.replace(/\-/g, '/'));
}



function toNumber(s) {
return Number(s.replace(/[^0-9\.]/g, ""));
}

function compareByColumn(nCol, bDescending, sType) {
var c = nCol;
var d = bDescending;

var fTypeCast = String;

if (sType == "Number")
fTypeCast = Number;
else if (sType == "Date")
fTypeCast = parseDate;
else if (sType == "CaseInsensitiveString")
fTypeCast = CaseInsensitiveString;

return function (n1, n2) {
if (fTypeCast(getInnerText(n1.cells[c])) < fTypeCast(getInnerText(n2.cells[c])))
return d ? -1 : +1;
if (fTypeCast(getInnerText(n1.cells[c])) > fTypeCast(getInnerText(n2.cells[c])))
return d ? +1 : -1;
return 0;
};
}

function sortColumnWithHold(e) {
// find table element
var el = ie5 ? e.srcElement : e.target;
var table = getParent(el, "TABLE");

// backup old cursor and onclick
var oldCursor = table.style.cursor;
var oldClick = table.onclick;

// change cursor and onclick
table.style.cursor = "wait";
table.onclick = null;

// the event object is destroyed after this thread but we only need
// the srcElement and/or the target
var fakeEvent = {srcElement : e.srcElement, target : e.target};


window.setTimeout(function () {
sortColumn(fakeEvent);
// once done resore cursor and onclick
table.style.cursor = oldCursor;
table.onclick = oldClick;
}, 100);
}

function sortColumn(e) {
var tmp = e.target ? e.target : e.srcElement;
var tHeadParent = getParent(tmp, "THEAD");
var el = getParent(tmp, "TD");

if (tHeadParent == null)
return;

if (el != null) {
var p = el.parentNode;
var i;

// typecast to Boolean
el._descending = !Boolean(el._descending);

if (tHeadParent.arrow != null) {
if (tHeadParent.arrow.parentNode != el) {
tHeadParent.arrow.parentNode._descending = null; //reset sort order
}
tHeadParent.arrow.parentNode.removeChild(tHeadParent.arrow);
}

if (el._descending)
tHeadParent.arrow = arrowUp.cloneNode(true);
else
tHeadParent.arrow = arrowDown.cloneNode(true);

el.appendChild(tHeadParent.arrow);



// get the index of the td
var cells = p.cells;
var l = cells.length;
for (i = 0; i < l; i++) {
if (cells[i] == el) break;
}

var table = getParent(el, "TABLE");
// can't fail

sortTable(table,i,el._descending, el.getAttribute("type"));
}
}


function getInnerText(el) {
if (ie5) return el.innerText; //Not needed but it is faster

var str = "";

var cs = el.childNodes;
var l = cs.length;
for (var i = 0; i < l; i++) {
switch (cs[i].nodeType) {
case 1: //ELEMENT_NODE
str += getInnerText(cs[i]);
break;
case 3: //TEXT_NODE
str += cs[i].nodeValue;
break;
}

}

return str;
}

function getParent(el, pTagName) {
if (el == null) return null;
else if (el.nodeType == 1 && el.tagName.toLowerCase() == pTagName.toLowerCase())
return el;
else
return getParent(el.parentNode, pTagName);
}
//-->
</SCRIPT>
<!--#include file="function.asp" -->
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
<fieldset>
<LEGEND><span class="style6">使用者帳號管理</span></LEGEND>

<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><div align="right">
      <input name="Submit" type="submit" onclick="MM_goToURL('self','user-upd.asp?op=add');return document.MM_returnValue" value="新增帳號" />
    </div></td>
  </tr>
</table>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3" onclick="sortColumn(event)">
  <thead>
    <tr>
      <td width="12%" bgcolor="#D0DFFF"><span class="style6">使用者帳號</span></td>
      <td width="14%" bgcolor="#D0DFFF"><span class="style6">英文名稱</span></td>
      <td width="15%" bgcolor="#D0DFFF"><span class="style6">姓名</span></td>
      <td width="18%" bgcolor="#D0DFFF"><span class="style6">部門</span></td>
      <td width="13%" bgcolor="#D0DFFF"><span class="style6">電話 </span></td>
      <td width="15%" bgcolor="#D0DFFF"><span class="style6">EMAIL</span></td>
      <td width="15%" bgcolor="#D0DFFF"><span class="style6">系統管理者</span></td>
      <td width="10%" bgcolor="#D0DFFF"><span class="style6">管理</span></td>
    </tr>
  </thead>
  <tbody>
    <%
  page=FunSwitch(CheckStr(request("page")),1)
  sqlstr_list = "SELECT A_id,LoginID,UserNameE,UserName,Tel,email,Group_id,A_right FROM useraccount order by A_id desc"
	'w sqlstr_list
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
      <td bgcolor="#FFFFFF"><span class="style6"><%=rs_list("LoginID")%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><%=rs_list("UserNameE")%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><%=rs_list("UserName")%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><%=showPub(rs_list("Group_id"))%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><%=rs_list("Tel")%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><%=rs_list("email")%></span></td>
      <td bgcolor="#FFFFFF"><span class="style6"><% if rs_list("A_right")="1" then w "是" else w "否" end if %></span></td>
      <td bgcolor="#FFFFFF"><span class="style6">
        <input name="btn_upd<%=rs_list("A_id")%>" type="button" onclick="MM_goToURL('self','user-upd.asp?id=<%=rs_list("A_id")%>&op=upd');return document.MM_returnValue;" value="修" />
        <input type="button" name="btn_del<%=rs_list("A_id")%>" value="刪" onclick="if(window.confirm('確定刪除?')){MM_goToURL('self','user-db.asp?id=<%=rs_list("A_id")%>&op=del');return document.MM_returnValue;}" />
      </span></td>
    </tr>
    <%
		rs_list.MOVENEXT
		IF rs_list.EOF THEN EXIT FOR
		Next
	%>
  </tbody>
  <%					
		end if
	%>
</table>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td colspan="2" bgcolor="#FFFFFF" class="style3"><%  CALL page_change_10("user.asp") %></td>
  </tr>
</table>
<%
FUNCTION showPub(x)
        
	if len(x)>0 then
		sqlstr_listtt = "SELECT UserName FROM department where A_id="&x&""
		SET rs_listtt = Server.CreateObject("ADODB.Recordset")
		rs_listtt.OPEN sqlstr_listtt, adoConn, 3,3
			suma=rs_listtt("UserName")
		rs_listtt.CLOSE
	end if
	showPub=suma
END FUNCTION


FUNCTION page_change_10(pg_name)

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
        RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&page="&jpage-10&"&epage="&epage-10&"' class=""c02""><font size='-1'> 上10頁</a>‧"
      END IF
      next_page=X+1
     
      '----------------------------------
      IF rs_list.pagecount<=10 THEN
         RESPONSE.WRITE ""
      ELSEIF cint(epage)>rs_list.pagecount THEN
         RESPONSE.WRITE ""
      ELSE
            RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&page="&jpage+10&"&epage="&epage+10&"' class=""c02"">下10頁 </a>&nbsp;&nbsp;&nbsp;&nbsp;"
      END IF

     '----------------------------------

      RESPONSE.WRITE "目前頁次: "&X&"　[ 頁次:"

      FOR J=cint(jpage) to cint(epage)
          IF J<=rs_list.pagecount THEN
           RESPONSE.WRITE " <a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&epage="&epage&"&page="&cint(J)&"' class=""c02"">"&J&"</A>"
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
sqlLog="insert into logss(UserName,logdate,url)values('"&session("admin_id")&"','"&TimeNow&"','"&mySCRIPT_NAME&Request.ServerVariables("QUERY_STRING")&"')"
adoconn.execute(sqlLog)
end if
%>