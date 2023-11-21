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
<style type="text/css">
<!--
.style9 {font-size: 9pt; color: #333333; }
.style550 {font-size: 10pt;color: #FF0000}
.style5 {font-size: 10pt; color: #333333; }
-->
</style>
<script src="prototype.js" type="text/javascript"></script>
<script src="pre_logined.js" type="text/javascript"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--
function CheckForm()
{
if(document.form1.ALoginID.value=="")
{
alert("使用者帳號未填寫!!");
document.form1.ALoginID.focus();
return false;
}
AA_no_re = /^[^\s]{3,5}$/;
AA_no_reEN = /^[a-zA-Z0-9]+$/;
if(!AA_no_re.test(document.form1.ALoginID.value))
{
alert("使用者帳號長度不符合!!");
document.form1.ALoginID.focus();
return false;
}
if(!AA_no_reEN.test(document.form1.ALoginID.value))
{
alert("使用者帳號僅限制數字與英文!!");
document.form1.ALoginID.focus();
return false;
}


AUserName_re = /^[\u0391-\uFFE5]+$/;
AUserName_reLe = /^[^\s]{2,4}$/;
if(document.form1.AUserName.value=="")
{
alert("中文姓名未填寫!!");
document.form1.AUserName.focus();
return false;
}
/*
if(!AUserName_re.test(document.form1.AUserName.value))
{
alert("中文姓名僅限制填寫中文字!!");
document.form1.AUserName.focus();
return false;
}
if(!AUserName_reLe.test(document.form1.AUserName.value))
{
alert("中文姓名長度不符合!!");
document.form1.AUserName.focus();
return false;
}
*/
ALoginPwd_reLe = /^[^\s]{4,16}$/;
if(document.form1.ALoginPwd.value=="")
{
alert("密碼未填寫!!");
document.form1.ALoginPwd.focus();
return false;
}
if(!ALoginPwd_reLe.test(document.form1.ALoginPwd.value))
{
alert("密碼長度不符合!!");
document.form1.ALoginPwd.focus();
return false;
}
/*
AUserNameE_re = /^[a-zA-Z]+$/;
AUserNameE_reLe = /^[^\s]{2,12}$/;
if(document.form1.AUserNameE.value=="")
{
alert("英文名稱未填寫!!");
document.form1.AUserNameE.focus();
return false;
}
if(!AUserNameE_re.test(document.form1.AUserNameE.value))
{
alert("英文名稱僅限制填寫英文字母!!");
document.form1.AUserNameE.focus();
return false;
}
if(!AUserNameE_reLe.test(document.form1.AUserNameE.value))
{
alert("英文名稱長度不符合!!");
document.form1.AUserNameE.focus();
return false;
}

ATel_re = /^[0-9-]+$/;
ATel_reLe = /^[^\s]{6,12}$/;
if(document.form1.ATel.value=="")
{
alert("電話未填寫!!");
document.form1.ATel.focus();
return false;
}
if(!ATel_re.test(document.form1.ATel.value))
{
alert("電話號碼僅可填寫數字!!");
document.form1.ATel.focus();
return false;
}
if(!ATel_reLe.test(document.form1.ATel.value))
{
alert("電話號碼長度不符合!!");
document.form1.ATel.focus();
return false;
}

//AEmail_re = /^[^\s]+@[^\s]+\.[^\s]+$/;
var pattern = /^([a-zA-Z0-9._-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
//(,"ig");
if(!pattern.test(document.form1.AEmail.value))
{
alert("EMAIL郵件格式錯誤!!");
document.form1.AEmail.focus();
return false;
}
*/

}

//-->
</SCRIPT>
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
op=FunSwitch(request("op"),2)
id=FunSwitch(request("id"),1)
if op="add" then
	txts="新增"
else
	txts="修改"
end if
%>

        <%
	if op="upd" then
	sqlstr_list = "SELECT UserNameE,UserName,Group_id,Tel,email,LoginID,LoginPwd,A_right FROM useraccount where A_id="&id&""
	SET rs_list = Server.CreateObject("ADODB.Recordset")
	rs_list.OPEN sqlstr_list, adoConn, 3,3		
		AUserNameE=rs_list("UserNameE")
		AUserName=rs_list("UserName")
		AGroup_id=rs_list("Group_id")
		ATel=rs_list("Tel")
		Aemail=rs_list("email")
		ALoginID=rs_list("LoginID")
		ALoginPwd=rs_list("LoginPwd")
		A_right=rs_list("A_right")		
	rs_list.CLOSE
	end if
	%>
<fieldset>
<LEGEND><span class="style9">使用者帳號管理</span></LEGEND>
<!--table width="70%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#A8B9CE" bordercolorlight="#000000" bordercolordark="#FFFFFF"-->
<form id="form1" name="form1" method="post" action="user-db.asp" onSubmit="return CheckForm()">
	  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3">
        <tr>
          <td width="18%" bgcolor="#FFFFFF"><span class="style9"><span class="style550">*</span>使用者代碼 </span></td>
          <td width="82%" bgcolor="#FFFFFF"><span class="style9"><input name="ALoginID" type="text" id="ALoginID" value="<%=ALoginID%>" />(登入時使用的帳號)</span></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9"><span class="style550">*</span>密碼</span></td>
          <td width="82%" bgcolor="#FFFFFF"><span class="style9"><input name="ALoginPwd" type="password" id="ALoginPwd" value="<%=ALoginPwd%>" maxlength="16" />(僅允許數字英文,4到16碼之間)</span></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9"><span class="style550">*</span>中文名字</span></td>
          <td bgcolor="#FFFFFF"><input name="AUserName" type="text" id="AUserName" value="<%=AUserName%>" /></td>
        </tr>        
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9">英文名字</span></td>
          <td bgcolor="#FFFFFF"><input name="AUserNameE" type="text" id="AUserNameE" value="<%=AUserNameE%>" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9"><span class="style550">*</span>所屬單位</span></td>
          <td bgcolor="#FFFFFF"><span class="style9">
          <!--input name="AGroup_id" type="text" id="AGroup_id" value="<%=AGroup_id%>" /></td-->		 
		<!--select name="ADep" id="ADep" onchange="chg_admin_dep(this.value)"-->
		<select name="AGroup_id" id="AGroup_id">
                  <%
		  sqlstr_listQ = "SELECT A_id,UserName FROM department order by A_id desc"
		  SET rs_listQ = Server.CreateObject("ADODB.Recordset")
		  rs_listQ.OPEN sqlstr_listQ, adoConn, 3,3
		  if not rs_listQ.eof then		  
		     for iiQ=1 to rs_listQ.recordcount
		  %>
		        <option value="<%=rs_listQ("A_id")%>" <% if FunSwitch(AGroup_id,1)=FunSwitch(rs_listQ("A_id"),1) then w "selected" end if %>><%=rs_listQ("UserName")%></option>
	          <%
		     rs_listQ.MOVENEXT
		     IF rs_listQ.EOF THEN EXIT FOR
		       Next		  
		  else
		  %>
		        <option value="0">無資料</option>
	          <%
		  end if
		  %>
	        </select>
		    </span></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9">電話 </span></td>
          <td bgcolor="#FFFFFF"><input name="ATel" type="text" id="ATel" value="<%=ATel%>" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><span class="style9">EMAIL</span></td>
          <td bgcolor="#FFFFFF"><input name="AEmail" type="text" id="AEmail" value="<%=AEmail%>" /></td>
        </tr>
	<tr>	
          <td bgcolor="#FFFFFF" class="style9">是否為系統管理者</td>
          <td bgcolor="#FFFFFF" class="style9"><span class="style9">
              <select name="A_right" id="A_right">
	        <option value="0">否</option>
                <option value="1" <% if A_right="1" then w "selected" end if %>>是</option>	        
	      </select>
          </span></td>
        </tr>        
      </table>
	  <input name="op" type="hidden" id="op" value="<%=op%>" />
	  <input name="id" type="hidden" id="id" value="<%=id%>" />
      <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td colspan="2" bgcolor="#FFFFFF"><div align="right">
              <input name="Submit" type="submit" value="<%=txts%>" />
          </div></td>
        </tr>
      </table>
</form>
</fieldset>
<%
sqlLog="insert into logss(UserName,logdate,url)values('"&session("admin_id")&"','"&TimeNow&"','"&mySCRIPT_NAME&Request.ServerVariables("QUERY_STRING")&"')"
adoconn.execute(sqlLog)
end if
%>