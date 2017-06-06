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
    Dim id
    id = Request.QueryString("id")
	origem = Request.QueryString("origem")
	origem = Replace(origem, "%3F","%3F")

	if IsEmpty(id) then
	%>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Vídeos</p>
    <p class="titulo">Galeria de Vídeos</p>
	<%
	bdtabela = "SELECT * FROM video ORDER BY id DESC"
	bd.open bdtabela, conexao
	while not bd.EOF
	%>
    <p class="texto">&#8226; <img src="imagens/icone_video.gif" width="14" height="14" border="0" /> <a href="video.asp?id=<%=bd("id")%>"><%=bd("titulo")%></a><%editarexcluir("video")%></p>
	<%
	bd.MoveNext
	wend
	bd.close
	else
	if (idexiste(id,"video") = "sim") then
	bdid = "SELECT * FROM video LEFT JOIN cadastro ON video.idusuarioadd = cadastro.id WHERE video.id=" & id
	bd.open bdid, conexao
	%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <%=bd("titulo")%></p>
    <p class="titulo"><object width="425" height="344"><param name="movie" value="http://www.youtube.com/v/<%=idlinkvideo(bd("conteudo"))%>&hl=pt-br&fs=1&rel=0&color1=0x3a3a3a&color2=0x999999"></param><param name="allowFullScreen" value="true"></param><embed src="http://www.youtube.com/v/<%=idlinkvideo(bd("conteudo"))%>&hl=pt-br&fs=1&rel=0&color1=0x3a3a3a&color2=0x999999" type="application/x-shockwave-flash" allowfullscreen="true" width="425" height="344"></embed></object></p>
    <p class="titulo" id="conteudoeditarvideo"><%=bd("titulo")%></p>
	<%
		if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
		%><div id="formeditarvideo" style="display:none">
		<form name="video" id="video" method="post" action="editar.asp?tipo=video">
		<input name="titulo" type="text" class="titulo" id="titulo" value="<%=bd("titulo")%>" size="50" maxlength="50" />
		<input type="hidden" name="idvideo" id="idvideo" value="<%=bd("video.id")%>" />
		<p class="texto" align="center"><strong>(<a href="javascript:void(0);" onclick="document.video.submit();">Atualizar</a> - <a href="javascript:void(0);" onclick="editar('video','cancelar');">Cancelar</a>)</strong></p>
		</form>
		</div>
		<div id="menueditarvideo" align="center"><p class="texto"><strong>(<a href="javascript:void(0);" onclick="editar('video','editar');">Editar</a>|<a href="javascript:void(0);" onclick="excluir('video','<%=bd("video.id")%>');">Excluir</a>)</strong></p></div><%
		if (Request.QueryString("acao") = "editar") then
		%><script>editar('video','editar');</script><%
		end if
		end if
		if(Len(idlinkvideo(bd("conteudo"))) <> 11) then
		%>
        <p class="texto">A id do vídeo parece não estar correta. Caso o vídeo não reproduza, verifique com quem adicionou para detalhes.</p>
		<%
		end if
		%><p class="texto"><%
		if origem <> Empty then
		%><a href="<%=origem%>">&laquo; voltar às atualizações</a><%
		else
		%><a href="video.asp">&laquo; voltar para vídeos</a><%
		end if
		%> | <%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%></p><%
	else
	idnaoexiste("galeria de vídeos")
	end if
	end if
	%>
    <%
	Function idlinkvideo (linkdovideo)
	linkdovideo = Replace(linkdovideo, "http://", "")
	linkdovideo = Replace(linkdovideo, "br.youtube.com/watch?v=", "")
	linkdovideo = Replace(linkdovideo, "www.youtube.com/watch?v=", "")
	linkdovideo = Replace(linkdovideo, "&feature=dir", "")
	idlinkvideo = linkdovideo
	End function
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>