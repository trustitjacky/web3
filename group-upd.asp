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
.style5 {font-size: 10pt; color: #333333; }
.style6 {
	font-size: 9pt;
	color: #333333;
}
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
op=FunSwitch(request("op"),2)
id=FunSwitch(request("id"),1)
if op="add" then
	txts="新增"
else
	txts="修改"
end if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function CheckForm()
{
if(document.form1.A_no.value=="")
{
alert("代碼未填寫!!");
document.form1.A_no.focus();
return false;
}
AA_no_re = /^[^\s]{3,10}$/;
AA_no_reEN = /^[a-zA-Z0-9]+$/;
if(!AA_no_re.test(document.form1.A_no.value))
{
alert("代碼長度不符合!!");
document.form1.A_no.focus();
return false;
}
if(!AA_no_reEN.test(document.form1.A_no.value))
{
alert("代碼僅限制數字與英文!!");
document.form1.A_no.focus();
return false;
}

AUserName_re = /^[\u0391-\uFFE5]+$/;
AUserName_reLe = /^[^\s]{2,20}$/;
if(document.form1.UserName.value=="")
{
alert("中文名稱未填寫!!");
document.form1.UserName.focus();
return false;
}
/*
if(!AUserName_re.test(document.form1.UserName.value))
{
alert("中文名稱僅限制填寫中文字!!");
document.form1.UserName.focus();
return false;
}
if(!AUserName_reLe.test(document.form1.UserName.value))
{
alert("中文名稱長度不符合!!");
document.form1.UserName.focus();
return false;
}

contacter_reLe = /^[^\s]{2,4}$/;
if(document.form1.contacter.value=="")
{
alert("聯絡人未填寫!!");
document.form1.contacter.focus();
return false;
}

contacter_re = /^[\u0391-\uFFE5]+$/;
if(!contacter_re.test(document.form1.contacter.value))
{
alert("聯絡人僅限制填寫中文字!!");
document.form1.contacter.focus();
return false;
}
if(!contacter_reLe.test(document.form1.contacter.value))
{
alert("聯絡人長度不符合!!");
document.form1.contacter.focus();
return false;
}
*/
/*
AUserNameE_re = /^[a-zA-Z]+$/;
AUserNameE_reLe = /^[^\s]{2,20}$/;
if(document.form1.UserNameE.value=="")
{
alert("英文名稱未填寫!!");
document.form1.UserNameE.focus();
return false;
}
if(!AUserNameE_re.test(document.form1.UserNameE.value))
{
alert("英文名稱僅限制填寫英文字母!!");
document.form1.UserNameE.focus();
return false;
}
if(!AUserNameE_reLe.test(document.form1.UserNameE.value))
{
alert("英文名稱長度不符合!!");
document.form1.UserNameE.focus();
return false;
}

ATel_re = /^[0-9-]+$/;
ATel_reLe = /^[^\s]{6,12}$/;
if(document.form1.TEL.value=="")
{
alert("電話未填寫!!");
document.form1.TEL.focus();
return false;
}
if(!ATel_re.test(document.form1.TEL.value))
{
alert("電話號碼僅可填寫數字!!");
document.form1.TEL.focus();
return false;
}
if(!ATel_reLe.test(document.form1.TEL.value))
{
alert("電話號碼長度不符合!!");
document.form1.TEL.focus();
return false;
}

//AEmail_re = /^[^\s]+@[^\s]+\.[^\s]+$/;
var pattern = /^([a-zA-Z0-9._-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
//(,"ig");
if(!pattern.test(document.form1.email.value))
{
alert("EMAIL郵件格式錯誤!!");
document.form1.email.focus();
return false;

}
*/

}

//-->
</SCRIPT><%
	if op="upd" then
	sqlstr_list = "SELECT A_id,UserName,UserNameE,TEL,Address,contacter,A_no,email FROM department where A_id="&id&""
	SET rs_list = Server.CreateObject("ADODB.Recordset")
	rs_list.OPEN sqlstr_list, adoConn, 3,3
		A_id=rs_list("A_id")
		UserName=rs_list("UserName")
		UserNameE=rs_list("UserNameE")
		TEL=rs_list("TEL")
		Address=rs_list("Address")
		contacter=rs_list("contacter")
		A_no=rs_list("A_no")
		email=rs_list("email")
	rs_list.CLOSE
	end if
	%>

<fieldset>
<LEGEND><span class="style6">權責單位管理</span></LEGEND>
<span class="style6"><a href="group.asp" class="style-topmenu"></a></span><br />
    <form id="form1" name="form1" method="post" action="group-db.asp" onSubmit="return CheckForm()">
	  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="e3e3e3">
        <tr>
          <td width="18%" bgcolor="#FFFFFF" class="style5">代碼 </td>
          <td width="82%" bgcolor="#FFFFFF"><input name="A_no" type="text" id="A_no" value="<%=A_no%>" maxlength="10" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" class="style5">中文名稱</td>
          <td width="82%" bgcolor="#FFFFFF"><input name="UserName" type="text" id="UserName" value="<%=UserName%>" maxlength="60" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" class="style5">英文名稱</td>
          <td bgcolor="#FFFFFF"><input name="UserNameE" type="text" id="UserNameE" value="<%=UserNameE%>" maxlength="60" /></td>
        </tr>

        <tr>
          <td bgcolor="#FFFFFF" class="style5">聯絡人</td>
          <td bgcolor="#FFFFFF"><input name="contacter" type="text" id="contacter" value="<%=contacter%>" maxlength="10" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" class="style5">電話 </td>
          <td bgcolor="#FFFFFF"><input name="TEL" type="text" id="TEL" value="<%=TEL%>" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" class="style5">EMAIL</td>
          <td bgcolor="#FFFFFF"><input name="email" type="text" id="email" value="<%=email%>" maxlength="20" /></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" class="style5">地址</td>
          <td bgcolor="#FFFFFF"><input name="Address" type="text" id="Address" value="<%=Address%>" maxlength="60" /></td>
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