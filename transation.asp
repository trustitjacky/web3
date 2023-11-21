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
%>

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
  '找出所屬部門與權限
  if len(session("username_login"))>0 then
  sqlstr_right = "SELECT A_right,Group_id FROM useraccount where LoginID='"&session("username_login")&"'"
  SET rs_right = Server.CreateObject("ADODB.Recordset")
	rs_right.OPEN sqlstr_right, adoConn, 3,3
          if not rs_right.eof then
 	      GroupID=rs_right("Group_id")
 	      rA_right=rs_right("A_right")
	  end if
	rs_right.CLOSE
  end if
  depName=""	
  if len(GroupID)>0 then 
  sqlstr_dep = "SELECT UserName FROM department where A_id="&GroupID&""
  SET rs_dep = Server.CreateObject("ADODB.Recordset")
	rs_dep.OPEN sqlstr_dep, adoConn, 3,3
          if not rs_dep.eof then
 	      depName=Trim(rs_dep("UserName")) 	      
	  end if
	rs_dep.CLOSE
  end if		
  %>

<fieldset>
<LEGEND><span class="style3">一卡通移送作業</span></LEGEND>
<form action="transation.asp" method="POST">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td class="style3">
          書名關鍵字<input name="kword" id="kword" type="text" value="<%=kword%>" size="6" />           
          圖書登錄號<input name="PID" id="PID" type="text" value="<%=PID%>" size="6" />
          
          通還狀態
              <select name="chkStatus" id="chkStatus">
                <option value="">=請選擇=</option>
                <option value="0">移送中</option>
                <option value="1">集中地已點收</option>
                <option value="2">集中地配送中</option>
                <option value="3">已點收</option>
	      </select>                   
          <% if  rA_right="1" then
	         sqlstr_dep = "SELECT UserName FROM department where 1=1"
                 SET rs_dep = Server.CreateObject("ADODB.Recordset")
	         rs_dep.OPEN sqlstr_dep, adoConn, 3,3
                 if not rs_dep.eof then %>
                    來源圖書館
                    <select name="Source_dep" id="Source_dep">
 	       	        <option value="">=請選擇=</option>
                     <% FOR SH=1 to rs_dep.recordcount %>                
                        <option value="<%=trim(rs_dep("UserName"))%>"><%=Trim(rs_dep("UserName"))%></option>
                     <% rs_dep.MOVENEXT                                
                        IF rs_dep.EOF THEN EXIT FOR 
                        Next %>	             
	            </select>						
          <%     end if
                 rs_dep.CLOSE 
               end if 
          %> 
	  <% 
	         sqlstr_dep = "SELECT UserName FROM department where 1=1"
                 SET rs_dep = Server.CreateObject("ADODB.Recordset")
	         rs_dep.OPEN sqlstr_dep, adoConn, 3,3
                 if not rs_dep.eof then %>
                    目的地圖書館
                    <select name="Target_dep" id="Target_dep">
                        <option value="">=請選擇=</option>      
 	       	     <% FOR SH=1 to rs_dep.recordcount %>                
                        <option value="<%=trim(rs_dep("UserName"))%>"><%=Trim(rs_dep("UserName"))%></option>
                     <% rs_dep.MOVENEXT                                
                        IF rs_dep.EOF THEN EXIT FOR 
                        Next %>	             
	            </select>						
          <%     end if
                 rs_dep.CLOSE                 
          %> <br>	               
      
          排序依據
          <select name="ordby">
            <option value="Send_Date desc" <% if ordby="Send_Date desc" then w "selected" end if %>>依移送日期近到遠</option>
            <option value="Send_Date asc" <% if ordby="Send_Date asc" then w "selected" end if %>>依移送日期遠到近</option>
            <option value="P_NO desc" <% if ordby="P_NO desc" then w "selected" end if %>>依圖書登錄號大到小</option>
            <option value="P_NO asc" <% if ordby="P_NO asc" then w "selected" end if %>>依圖書登錄號小到大</option>
          </select>
          每頁資料數
          <select name="myPageSize">
            <option value="10" <% if myPageSize=10 then w "selected" end if %>>10</option>
            <option value="20" <% if myPageSize=20 then w "selected" end if %>>20</option>
            <option value="50" <% if myPageSize=50 then w "selected" end if %>>50</option>
          </select>
          <input type="submit" name="Submit5" value="執行" /></td>
      </tr>
     </table>
</form><br />

   <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td colspan="2" bgcolor="#FFFFFF"><div align="right">
            <input name="Upload" type="submit" onclick="window.open('upload/upfile2.asp','fileup','width=530,height=50')" value="更新出版品通還資料" />
        </div></td>
      </tr>
    </table>
      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3">
        <tr>
          <td width="12%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">移送日期</div>
          </div></td>
          <td width="10%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">圖書登錄號</div>
          </div></td>
          <td width="20%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">書刊名</div>
          </div></td>
          <td width="10%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">來源圖書館</div>
          </div></td>
          <td width="10%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">目的地圖書館</div>
          </div></td>
          <td width="12%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">點收日期</div>
          </div></td>
          <td width="6%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">狀態</div></td>
          <td width="10%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">點收人</div></td>            
          <!--td width="10%" bgcolor="#D0DFFF"><div align="center" class="style3">
            <div align="center">維護</div></td-->  
        </tr>
  <%
  sqlstr_list = "SELECT P_id,P_NO,P_Name,Source_id,Target_id,Update_User,Send_Date,Recive_Date,Status FROM publishes where 1=1 "
  if len(kword)>0 then
  	sqlstr_list = sqlstr_list &" and P_Name like '%"&kword&"%' "
  end if
  if len(PID)>0 then
  	sqlstr_list = sqlstr_list &" and P_NO ='"&PID&"' "
  end if
  if len(Target_dep)>0 then
  	sqlstr_list = sqlstr_list &" and Target_id ='"&Target_dep&"' "
  end if
  if len(chkStatus)>0 then
  	sqlstr_list = sqlstr_list &" and status = '"&chkStatus&"' "
  end if
  if rA_right<>1 then
        sqlstr_list = sqlstr_list &" and Source_id='"&depName&"' "
  else
        if len(Source_dep)>0 then
           sqlstr_list = sqlstr_list &" and Source_id='"&Source_dep&"' "
        end if
  end if
  sqlstr_list = sqlstr_list &"order by "&ordby
	'w sqlstr_list
  SET rs_list = Server.CreateObject("ADODB.Recordset")
  rs_list.OPEN sqlstr_list, adoConn, 3,3
	'response.write rs_list.recordcount
  IF page = 0 THEN Page = 1 END IF
  rs_list.PageSize = myPageSize           ' 設定每頁顯示 10 筆
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
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Send_Date")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("P_NO")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("P_Name")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Source_id")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Target_id")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Recive_Date")%></span></td>
          <td bgcolor="#FFFFFF"><span class="style3">
              <%if rs_list("Status")=0 then 
                   w "移送中"
                elseif rs_list("Status")=1 then 
                       w "集中地已點收" 
                elseif rs_list("Status")=2 then 
                       w "集中地配送中"  
                else w "已點收" 
                end if %></span></td>
          <td bgcolor="#FFFFFF"><span class="style3"><%=rs_list("Update_User")%></span></td>
          <!--td bgcolor="#FFFFFF"><span class="style3">
          <% 'if rs_list("Status")=2 then %>
          <input name="btn_upd<%=rs_list("P_id")%>" type="button" onclick="MM_goToURL('self','circulation-db.asp?id=<%=rs_list("P_id")%>&search_e_title=<%=SESSION("search_e_title")%>&kword=<%=Server.UrlEncode(kword)%>&PID=<%=PID%>&chkStatus=<%=chkStatus%>&ordby=<%=ordby%>&myPageSize=<%=myPageSize%>&page=<%=page%>&epage=<%=epage%>&Source_dep=<%=Server.UrlEncode(Source_dep)%>&Target_dep=<%=Server.UrlEncode(Target_dep)%>');return document.MM_returnValue"  value="點收" /></span></td-->
          <% 'end if %>
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
          <td colspan="2" bgcolor="#FFFFFF" class="style3"><%  CALL page_change_10("transation.asp",kword,PID,chkStatus,ordby,myPageSize,Source_dep,Target_dep) %></td>
          <td colspan="2" bgcolor="#FFFFFF"><div align="right">
             計<input name="text" type="text" size="8" readonly value="<%=rs_list.recordcount%>" />筆  
          </div></td>      
        </tr>
    </table>
<%
FUNCTION page_change_10(pg_name,kword,PID,chkStatus,ordby,myPageSize,Source_dep,Target_dep)

      X=page
      epage=REQUEST("epage")
      rd=rs_list.RecordCount
      rp=rs_list.PageSize
      
      IF epage=""or epage <10 THEN
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
        RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&kword="&Server.UrlEncode(kword)&"&PID="&PID&"&chkStatus="&chkStatus&"&ordby="&ordby&"&myPageSize="&myPageSize&"&page="&jpage-10&"&epage="&epage-10&"&Source_dep="&Server.UrlEncode(Source_dep)&"&Target_dep="&Server.UrlEncode(Target_dep)&"' class=""c02""><font size='-1'> 上10頁</a>‧"
      END IF
      next_page=X+1
     
      '----------------------------------
      IF rs_list.pagecount<=10 THEN
         RESPONSE.WRITE ""
      ELSEIF cint(epage)>rs_list.pagecount-1 THEN
         RESPONSE.WRITE ""
      ELSE
            RESPONSE.WRITE "<a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&kword="&Server.UrlEncode(kword)&"&PID="&PID&"&chkStatus="&chkStatus&"&ordby="&ordby&"&myPageSize="&myPageSize&"&page="&jpage+10&"&epage="&epage+10&"&Source_dep="&Server.UrlEncode(Source_dep)&"&Target_dep="&Server.UrlEncode(Target_dep)&"' class=""c02""><font size='-1'>下10頁 </a>&nbsp;&nbsp;&nbsp;&nbsp;"
      END IF

     '----------------------------------

      RESPONSE.WRITE "目前頁次: "&X&"　[ 頁次:"

      FOR J=cint(jpage) to cint(epage)
          IF J<=rs_list.pagecount THEN
           RESPONSE.WRITE " <a href='"&pg_name&"?search_e_title="&SESSION("search_e_title")&"&kword="&Server.UrlEncode(kword)&"&PID="&PID&"&chkStatus="&chkStatus&"&ordby="&ordby&"&myPageSize="&myPageSize&"&page="&cint(J)&"&epage="&epage&"&Source_dep="&Server.UrlEncode(Source_dep)&"&Target_dep="&Server.UrlEncode(Target_dep)&"' class=""c02"">"&J&"</A>"
          END IF
      NEXT

      RESPONSE.WRITE " ]"
      'RESPONSE.WRITE "<br><br><br>"
      RESPONSE.WRITE "</td>"
      RESPONSE.WRITE "</tr>"
      RESPONSE.WRITE "</table>"
    END FUNCTION
  
%>
</fieldset><%
sqlLog="insert into logss(UserName,logdate,url)values('"&session("admin_id")&"','"&TimeNow&"','"&mySCRIPT_NAME&Request.ServerVariables("QUERY_STRING")&"')"
adoconn.execute(sqlLog)
%>