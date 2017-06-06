<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/default.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>1º ano A</title>
<!-- InstanceEndEditable -->
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script src="script.js"></script>
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>

<body>
<%
Set conexao = Server.CreateObject("ADODB.Connection")
conexao.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("primeiro.mdb")&";"
Set bd = Server.CreateObject ("ADODB.Recordset")
%>
<div id="principal">
  <div id="logo">
    <a accesskey="i" href="default.asp"><img src="imagens/icone_inicio.jpg" alt="Início" vspace="8" border="0" /></a>  </div>
  <div id="alunos">
  <%
  if(session("idusuario") <> "") then
  %><form id="buscar" name="buscar" method="post" action="buscar.asp">
  <table width="470" border="0">
  <tr>
    <td>
    <div align="left">
      <label>
      <input name="buscar" type="text" class="texton" id="buscar" />
      </label>
      <label>
      <select name="tipo" class="texton" id="tipo">
        <option value="tudo">Tudo</option>
        <option value="agenda">Agenda</option>
        <option value="comunicados">Comunicados</option>
        <option value="caderno">Caderno</option>
        <option value="foto">Fotos</option>
        <option value="video">Vídeos</option>
      </select>
      </label>
      <input name="botaobuscar" type="submit" class="texton" id="botaobuscar" value="Buscar" />
    </div>
    </td>
    <td width="50"><div align="right"><a accesskey="s" href="login.asp?acao=logout">Sair</a></div></td>
  </tr>
  <tr>
    <td>
    <div align="left"><a accesskey="c" href="cadastro.asp">Cadastro</a> | <a accesskey="c" href="contato.asp">Contato</a> | <a accesskey="c" href="comentarios.asp">Comentários</a><%
	if session("tipo") = "admin" then
	%> | <a accesskey="u" href="upload.asp">Upload</a><%
	end if
	if session("idusuario") = 7 AND session("tipo") <> "admin" then
	%> | <a accesskey="a" href="login.asp?acao=admin">Admin</a><%
	end if
	%></div>
    </td>
    <td><div align="right"><a accesskey="b" href="blog.asp">Blog</a></div></td>
  </tr>
  <tr>
    <td>
    <div align="left"><%
	bdid = "SELECT * FROM cadastro WHERE id=" & session("idusuario")
	bd.open bdid, conexao
	if(bd("dianasc") = day(date) AND bd("mesnasc") = month(date)) then
		Response.Write("Parabéns, ")
	else
		if(hour(time()) >= 6 AND hour(time()) <= 11 ) then
			Response.Write("Bom Dia, ")
		end if
		if(hour(time()) >= 12 AND hour(time()) <= 18 ) then
			Response.Write("Boa Tarde, ")
		end if
		if(hour(time()) >= 19 AND hour(time()) <= 23 ) then
			Response.Write("Boa Noite, ")
		end if
		if(hour(time()) >=00  AND hour(time()) <= 5 ) then
			Response.Write("Você não dorme ein ")
		end if
	end if
	Response.Write(bd("nome"))
	bd.close
	%></div>
    </td>
    <td><div align="right"><a href="alunos.asp">Alunos</a></div></td>
  </tr>
  </table>
  </form><%
  else
  %><table width="470" border="0">
  <tr>
    <td rowspan="3">
      <div align="center"><a href="alunos.asp"><img src="imagens/alunos.gif" alt="Alunos" width="305" height="85" align="middle" border="0" /></a></div>
    </td>
    <td width="100"><div align="right"><a accesskey="b" href="blog.asp">Blog</a></div></td>
  </tr>
  <tr>
    <td><div align="right"><a accesskey="c" href="cadastro.asp">Cadastrar</a></div></td>
  </tr>
  <tr>
    <td><div align="right"><a accesskey="e" href="login.asp">Entrar</a></div></td>
  </tr>
  </table>
  <%end if%></div>
  <div id="lateral">
    <div id="lateralbox">
      <p>Menu:</p>
      <p class="texto"><a href="atualizacoes.asp">Atualizações</a></p>
      <p class="texto"><a href="agenda.asp">Agenda</a></p>
      <p class="texto"><a href="comunicados.asp">Comunicados</a></p>
      <p class="texto"><a href="horario.asp">Horário</a></p>
      <p class="texto"><a href="galeria.asp">Fotos e Vídeos</a></p>
      <p class="texto"><a href="caderno.asp">Caderno</a></p>
      <p class="texto"><a href="localizador.asp">Localizador</a></p>
    </div>
    <div id="lateralbox">
      <p><%
	if(session("tipo") = "admin") then
	  %><a href="adicionar.asp?tipo=links">[+]</a> <%
	end if
	%>Links:</p><%
		bdid = "SELECT * FROM links"
		bd.open bdid, conexao
		while not bd.EOF%>
      <p class="texto"><a href="<%=bd("site")%>"<%
	  if left(bd("site"), 7) = "http://" then
	  %> target="_blank"<%
	  end if
	  %>><%=bd("titulo")%></a><%
	  if(session("tipo") = "admin") then
	  %><a href="javascript:void(0);" onclick="excluir('links','<%=bd("id")%>');"> [X]</a><%
	  end if
	  %></p><%
		bd.MoveNext
		wend
		bd.close%>
    </div>
  </div>
  <div id="conteudo">
  <!--#include file="funcoes.asp"-->
  <!-- InstanceBeginEditable name="conteudo" -->
<%
if session("tipo") <> "admin" then
	Response.Redirect("default.asp")
end if

Dim acao

acao = Request.QueryString("acao")

Select Case acao

Case "enviar"

On Error Resume Next
Dim objUpload
Set objUpload = Server.CreateObject("Dundas.Upload.2")
objUpload.UseVirtualDir = True
objUpload.UseUniqueNames = False
path = "arquivos/"
objUpload.Save (path)

If Err <> 0 Then

%>
<p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="upload.asp">Upload</a> &nbsp;&gt;&nbsp; Erro</p>
<p class="texto"><b>Houve erro(s) ao carregar o(s) arquivo(s) anexado(s)!</b></p>
<p class="texto">Erro: <%=Err.Number & " - " & Err.Description%></p>
<input type="button" value="Voltar" class="texton" onclick="javascript:history.back(-1)" /><%

Else

%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="upload.asp">Upload</a> &nbsp;&gt;&nbsp; Arquivos</p><%
contador = 0
For Each objUploadedFile in objUpload.Files%>
<p class="texto"><strong><%=objUploadedFile.TagName%></strong></p>
<p class="texto">
Tipo de Arquivo: <%=objUploadedFile.ContentType%>
<br />
Arquivo Original: <%=objUploadedFile.OriginalPath%>
<br />
Link do Arquivo: <%=linkarquivo(objUploadedFile.Path)%>
<br />
Tamanho: <%=objUploadedFile.Size%> bytes
</p>
<%
contador = contador + 1
Next

if contador = 0 then
%><p class="texto"><strong>Nenhum arquivo foi selecionado!</strong></p>
<input type="button" value="Voltar" class="texton" onclick="javascript:history.back(-1)" /><%
elseif contador = 1 then
%><p class="texto"><strong>O arquivo acima foi enviado com sucesso!</strong></p><%
else
%><p class="texto"><strong>Os arquivos acima foram enviados com sucesso!</strong></p><%
end if

End If

Set objUpload = Nothing

Case else

%>
<p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Upload</p>
<p class="titulo">Upload</p>
<form method="post" action="upload.asp?acao=enviar" enctype="multipart/form-data">
<p class="texto">Arquivo 1: <input type="file" name="arquivo1" class="texton" size="50"></p>
<p class="texto">Arquivo 2: <input type="file" name="arquivo2" class="texton" size="50"></p>
<p class="texto">Arquivo 3: <input type="file" name="arquivo3" class="texton" size="50"></p>
<input type="submit" value="Enviar" class="texton">
</form>
<p class="texto">Evite enviar arquivos com acento pois o nome do arquivo quando salvo no servidor irá ficar com caracteres desconfigurados.</p>
<%
End Select

Function linkarquivo (textoconteudo)
textoconteudo = Replace(textoconteudo, "C:\arquivos\sites\primeiroasp","http://clips.sytes.net/primeiro")
textoconteudo = Replace(textoconteudo, "\","/")
linkarquivo = textoconteudo
End function

%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>