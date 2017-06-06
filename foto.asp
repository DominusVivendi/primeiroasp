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
    Dim id, idgaleria
    id = Request.QueryString("id")
    idgaleria = Request.QueryString("idgaleria")
	origem = Request.QueryString("origem")
	origem = Replace(origem, "%3F","%3F")

	if IsEmpty(id) AND IsEmpty(idgaleria) then
	%>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Álbuns</p>
    <p class="titulo">Galeria de Fotos</p>
	<%
	bdtabela = "SELECT * FROM galeriafoto ORDER BY id DESC"
	bd.open bdtabela, conexao
	while not bd.EOF
	%>
    <p class="texto">&#8226; <img src="imagens/icone_foto.gif" width="14" height="14" border="0" /> <a href="foto.asp?idgaleria=<%=bd("id")%>"><%=bd("titulo")%></a><%editarexcluir("foto")%></p>
	<%
	bd.MoveNext
	wend
	bd.close
	else
	if (idexiste(id,"foto") = "sim" OR idexiste(idgaleria,"galeriafoto") = "sim") then
		if(idexiste(idgaleria,"galeriafoto") = "sim" AND IsEmpty(id)) then
		expSQL = "SELECT * FROM galeriafoto LEFT JOIN cadastro ON galeriafoto.idusuarioadd = cadastro.id WHERE galeriafoto.id=" & idgaleria
		bd.open expSQL, conexao
		%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <%=bd("titulo")%></p><%
		%>
        <p class="titulo" id="conteudoeditargaleriafotoo"><%=bd("titulo")%></p>
        <%
		if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
		%><div id="formeditargaleriafotoo" style="display:none">
		<form name="galeriafotoo" id="galeriafotoo" method="post" action="editar.asp?tipo=galeriafoto">
		<input name="titulo" type="text" class="titulo" id="titulo" value="<%=bd("titulo")%>" size="50" maxlength="50" />
		<input type="hidden" name="idgaleriafoto" id="idgaleriafoto" value="<%=bd("galeriafoto.id")%>" />
		<p class="texto" align="center"><strong>(<a href="javascript:void(0);" onclick="document.galeriafotoo.submit();">Atualizar</a> - <a href="javascript:void(0);" onclick="editar('galeriafotoo','cancelar');">Cancelar</a>)</strong></p>
		</form>
		</div>
        <div id="menueditargaleriafotoo" align="center"><p class="texto"><strong>(<a href="javascript:void(0);" onclick="editar('galeriafotoo','editar');">Editar</a>|<a href="javascript:void(0);" onclick="excluir('galeriafoto','<%=bd("galeriafoto.id")%>');">Excluir</a>)</strong></p></div><%
		if(session("tipo") = "aluno" OR session("tipo") = "admin") then
		%><p class="texto"><a href="adicionar.asp?tipo=foto&idgaleria=<%=bd("galeriafoto.id")%>"><strong>Adicionar fotos</strong></a></p><%
		end if
		if (Request.QueryString("acao") = "editar") then
		%><script>editar('galeriafotoo','editar');</script><%
		end if
		end if
		bd.close
		expSQL = "SELECT * FROM foto WHERE idgaleria=" & idgaleria
		bd.CursorLocation = 3
		bd.open expSQL, conexao
		while not bd.EOF
		%>
        <a href="foto.asp?id=<%=bd("id")%>"><img src="<%=bd("conteudo")%>" width="150" height="120" border="0" alt="<%=bd("titulo")%>" /></a>
		<%
		bd.MoveNext
		wend
		bd.close
		expSQL = "SELECT * FROM galeriafoto LEFT JOIN cadastro ON galeriafoto.idusuarioadd = cadastro.id WHERE galeriafoto.id=" & idgaleria
		bd.open expSQL, conexao
		%><p class="texto"><%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%></p><%
		bd.close
		end if
		if(idexiste(id,"foto") = "sim") then
		expSQL = "SELECT * FROM foto LEFT JOIN galeriafoto ON foto.idgaleria = galeriafoto.id WHERE foto.id=" & id
		bd.open expSQL, conexao
		%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <a href="foto.asp?idgaleria=<%=bd("galeriafoto.id")%>"><%=bd("galeriafoto.titulo")%></a> &nbsp;&gt;&nbsp; <%
		if bd("foto.titulo") = Empty then
		%>Foto<%
		else
		%><%=bd("foto.titulo")%><%
        end if
		%></p><%
		idgaleria = bd("idgaleria")
		bd.close
		expSQL = "SELECT * FROM foto WHERE idgaleria=" & idgaleria
		bd.open expSQL, conexao
			verificarid = "nao"
			while not bd.EOF AND verificarid = "nao"
			if( Abs(bd("id")) = Abs(id) ) then
			verificarid = "sim"
			end if
			bd.MoveNext
			if( bd.EOF ) then
			proximafoto = "fim"
			else
			proximafoto = bd("id")
			end if
			wend
		bd.close
		expSQL = "SELECT * FROM foto WHERE idgaleria=" & idgaleria
		bd.CursorLocation = 3
		bd.open expSQL, conexao
		bd.movelast
			verificarid = "nao"
			while not bd.BOF AND verificarid = "nao"
			if( Abs(bd("id")) = Abs(id) ) then
			verificarid = "sim"
			end if
			bd.MovePrevious
			if( bd.BOF ) then
			anteriorfoto = "inicio"
			else
			anteriorfoto = bd("id")
			end if
			wend
		bd.close
		expSQL = "SELECT * FROM foto WHERE id=" & id
		bd.open expSQL, conexao
    		%><p class="titulo"><%
			if(anteriorfoto = "inicio") then
			%><img src="imagens/icone_anterior_desabilitado.gif" border="0" /> <%
			else
			%><a href="foto.asp?id=<%=anteriorfoto%>"><img src="imagens/icone_anterior.gif" border="0" /></a> <%
			end if
			if(proximafoto = "fim") then
			%><img src="imagens/icone_proximo_desabilitado.gif" border="0" /><%
			else
			%><a href="foto.asp?id=<%=proximafoto%>"><img src="imagens/icone_proximo.gif" border="0" /></a><%
			end if
			%></p>
            <p class="titulo"><img src="<%=bd("conteudo")%>" width="450" height="338" border="0" /></p>
            <p class="titulo" id="conteudoeditarfoto"><%=bd("titulo")%></p>
            <%
			if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
			%><div id="formeditarfoto" style="display:none">
            <form name="foto" id="foto" method="post" action="editar.asp?tipo=foto">
            <input name="titulo" type="text" class="titulo" id="titulo" value="<%=bd("titulo")%>" size="50" maxlength="50" />
            <input type="hidden" name="idfoto" id="idfoto" value="<%=bd("id")%>" />
            <p class="texto" align="center"><strong>(<a href="javascript:void(0);" onclick="document.foto.submit();">Atualizar</a> - <a href="javascript:void(0);" onclick="editar('foto','cancelar');">Cancelar</a>)</strong></p>
            </form>
            </div>
<div id="menueditarfoto" align="center"><p class="texto"><strong>(<a href="javascript:void(0);" onclick="editar('foto','editar');">Editar</a>|<a href="javascript:void(0);" onclick="excluir('foto','<%=bd("id")%>');">Excluir</a>)</strong></p></div><%
			end if
			bd.close
			expSQL = "SELECT * FROM foto LEFT JOIN cadastro ON foto.idusuarioadd = cadastro.id WHERE foto.id=" & id
			bd.open expSQL, conexao
			%>
			<p class="texto"><%
			if origem <> Empty then
			%><a href="<%=origem%>">&laquo; voltar às atualizações</a><%
			else
			%><a href="foto.asp?idgaleria=<%=bd("idgaleria")%>">&laquo; voltar para álbum</a><%
			end if
			%> | <%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%></p>
			<script>function keypresed() {
            if ((window.event ? event.keyCode : event.which) == 37 && "<%=anteriorfoto%>" != "inicio")
            {
            document.location.href = "foto.asp?id=<%=anteriorfoto%>"
            }
            if ((window.event ? event.keyCode : event.which) == 39 && "<%=proximafoto%>" != "fim")
            {
            document.location.href = "foto.asp?id=<%=proximafoto%>"
            }
            <%if session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin" then%>
            if ((window.event ? event.keyCode : event.which) == 46)
            {
            excluir('foto','<%=id%>');
            }
            <%end if%>
            }
            </script>
            <script>document.onkeydown = keypresed;</script>
			<%
			bd.close
		end if
	else
	idnaoexiste("galeria de fotos")
	end if
	end if
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>