<%
	Response.Charset="utf-8"
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","no-cache"
	Response.ContentType="text/xml"
%>

<%
	DbPath = SERVER.MapPath("data.mdb")
	Set StrConnect = Server.CreateObject("ADODB.Connection")
	StrConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbPath

	str=Request("LABNMABV")

	IF str<>"" Then
        	Set objName=Server.CreateObject("ADODB.Recordset")
        	sqlCMD="select LABNMABV from laboratory WHERE LABNMABV LIKE '" & str & "%' GROUP BY LABNMABV"
        	objName.Open sqlCMD,StrConnect,2,3
  END IF

	Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>"
	Response.Write "<root>"
	IF str<>"" Then
		WHILE NOT objName.EOF
			Response.Write "<name>"&objName("LABNMABV")&"</name>"
		objName.MoveNext
		WEND
	END IF
	Response.Write "</root>"
%>
