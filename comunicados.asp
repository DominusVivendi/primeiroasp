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
		%>
    <%
	if IsEmpty(id) then
	%>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Comunicados</p>
    <p class="titulo">Comunicados</p>
	<%
	if(session("tipo") = "aluno" OR session("tipo") = "admin") then
	%><p class="texto"><a href="adicionar.asp?tipo=comunicado"><strong>Novo</strong></a></p><%
	end if
	bdtabela = "SELECT * FROM comunicados ORDER BY id DESC"
	bd.open bdtabela, conexao
	while not bd.EOF
	%>
    <p class="texto">&#8226; <img src="imagens/icone_comunicados.gif" width="14" height="14" border="0" /> <a href="comunicados.asp?id=<%=bd("id")%>"><%=bd("titulo")%></a><%editarexcluir("comunicados")%></p>
	<%
	bd.MoveNext
	wend
	bd.close
	else
	if (idexiste(id,"comunicados") = "sim") then
	bdid = "SELECT * FROM comunicados LEFT JOIN cadastro ON comunicados.idusuarioadd = cadastro.id WHERE comunicados.id=" & id
	bd.open bdid, conexao
	%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="comunicados.asp">Comunicados</a> &nbsp;&gt;&nbsp; <%=bd("titulo")%></p>
    <div id="conteudoeditarcomunicados">
    <p class="titulo"><%=bd("titulo")%></p>
    <p class="texto"><div class='texto'><%=substituirtags(bd("conteudo"))%></div></p>
    </div>
	<%
    if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
    %><div id="formeditarcomunicados" style="display:none">
    <form id="comunicados" name="comunicados" method="post" action="editar.asp?tipo=comunicados" onsubmit="return validacamposbranco('comunicado');">
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="titulo" id="titulo" value="<%=bd("titulo")%>" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Conteúdo:</label><br />
      <textarea name="conteudoa" id="conteudoa" cols="75" rows="20" class="texto"><%=bd("conteudo")%></textarea>
      </label>
    </p>
    <p id="infoconteudoa" class="texto" style="display: none"><strong>digite algum conteúdo</strong></p>
    <input type="hidden" name="idcomunicados" id="idcomunicados" value="<%=bd("comunicados.id")%>" />
    <p align="center">
        <label>
        <input type="submit" name="add" id="add" value="Atualizar" class="texton" />
        </label>
        <label>
        <input type="button" name="cancelar" id="cancelar" value="Cancelar" class="texton" onclick="editar('comunicados','cancelar');validacamposbranco('comunicado');" />
        </label>
    </p>
    </form>
    </div>
    <div id="menueditarcomunicados" align="center"><p class="texto"><strong>(<a href="javascript:void(0);" onclick="editar('comunicados','editar');">Editar</a>|<a href="javascript:void(0);" onclick="excluir('comunicados','<%=bd("comunicados.id")%>');">Excluir</a>)</strong></p></div><%
	if (Request.QueryString("acao") = "editar") then
	%><script>editar('comunicados','editar');</script><%
	end if
    end if
	%><p class="texto"><%
	if origem <> Empty then
	%><a href="<%=origem%>">&laquo; voltar às atualizações</a><%
	else
	%><a href="comunicados.asp">&laquo; voltar para comunicados</a><%
	end if
	%> | <a href="javascript:void(0);" onclick="imprimir('comunicados','<%=bd("comunicados.id")%>');">Imprimir</a> | <%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%></p><%
	else
	idnaoexiste("comunicados")
	end if
	end if
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>