<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1º ano A - Imprimir</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script src="script.js"></script>
</head>

<body onload="print();">
<!--#include file="funcoes.asp"-->
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("primeiro.mdb")&";"

dim tipo, id

tipo = Request.QueryString("tipo")
id = Request.QueryString("id")

if tipo = "comunicados" OR tipo = "caderno" OR tipo = "localizador" AND IsNumeric(id) then

expSQL = "SELECT * FROM " & tipo & " WHERE id=" & id
set rs = cn.execute(expSQL)

%>
<p class="titulo"><%=rs("titulo")%></p>
<p class="texto"><div class='texto'><%=substituirtags(rs("conteudo"))%></div></p>
<%

end if
%>
<p align="center" class="texto"><strong>1º ano A - Ensino Médio | 2008</strong>
<br />
<%=date%> | <%=time%></p>
</body>
</html>
