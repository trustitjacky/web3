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
<!-- Mirrored from www.gatsby.hk/ by HTTrack Website Copier/3.x [XR&CO'2007], Tue, 15 May 2007 10:20:52 GMT -->
<!-- Added by HTTrack --><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
.style3 {
	color: #FF0000;
	font-size: 9pt;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>一卡通系統</title>
</head>
<style type="text/css">
<!--
.style2 {font-size: 9pt}
-->
</style>
<script>
function checkReg(){
	x=this.form1.username;
	y=this.form1.password;
	if(x.lenght<=3||y.length<=3||x.value==''||y.value==''){
		alert('請輸入完整帳號密碼');
		return false;
	}else{
		form1.action='login_post.asp';
  		form1.submit();
  		return true;
	}
}
</script>
<%
t=request("t")

%>
<form id="form1" name="form1" method="post" onsubmit="return checkReg();">
<table width="100%" border="0" cellspacing="27">
  <tr>
    <td><table width="530" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td background="img/common/line_pattern1.gif">&nbsp;</td>
      </tr>
      <tr>
        <td></td>
      </tr>
      <tr>
        <td><table width="866" cellspacing="20">
          <tr>
            <td><div align="center"><strong>一卡通系統</strong> <br />
            </div>
                    <% if t="a" then %>
					<div align="center" class="style3">抱歉，您的帳號正在使用；若您確認已經登出，請稍候1分鐘登入謝謝 </div>
                    <% end if %>
                    <table width="500" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                      <tr>
                        <td width="105" height="30" valign="middle" bgcolor="#FFFFFF"><div align="right"><span class="style2">帳號</span></div></td>
                        <td width="395" valign="middle" bgcolor="#FFFFFF"><input name="username" type="text" id="username" /></td>
                      </tr>
                      <tr>
                        <td height="30" valign="middle" bgcolor="#FFFFFF"><div align="right"><span class="style2">密碼</span></div></td>
                        <td valign="middle" bgcolor="#FFFFFF"><input name="password" type="password" id="password" />
                            <input type="submit" name="Submit" value="送出" /></td>
                      </tr>
                  </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
