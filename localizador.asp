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
    Dim id, palavrachave, mostrar
    id = Request.QueryString("id")
    palavrachave = Request.QueryString("palavrachave")
	palavrachave = Replace(palavrachave, "'", "")
	palavrachave = Replace(palavrachave, "%", "")
	mostrar = Request.QueryString("mostrar")

	if IsEmpty(id) then
	%>
	  <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Localizador</p>
      <p class="titulo">Localizador</p>
	  <%
		if(session("tipo") = "aluno" OR session("tipo") = "admin") then
		%><p class="texto"><a href="adicionar.asp?tipo=localizador"><strong>Novo</strong></a><%if(session("tipo") = "admin") then
		%> | <a href="localizador.asp?mostrar=todos">Ver todos os registros</a><%end if%></p><%
		end if
	  %>
	  <p class="texto">Digite a palavra chave que foi fornecida para localizar o arquivo/página requerida!</p>
		<form id="form1" name="form1" method="get" action="localizador.asp" onsubmit="return validacamposbranco('buscalocalizador');">
		  <label class="texto">Localizar:
		  <input name="palavrachave" type="text" id="palavrachave" size="40" class="texton" />
		  </label>
		  <label>
		  <input type="submit" value="Localizar" id="botao" class="texton" />
		  </label>
          <p class="texto" id="infopalavrachave" style="display: none"><strong>Digite uma palavra chave</strong></p>
		</form>
		<script>document.form1.palavrachave.focus();</script>
	  <%
	  if palavrachave <> "" OR mostrar = "todos" then
		  %><p class="titulo">Resultados:</p><%
		  qtdresultado = 0
		  if session("tipo") = "" OR session("tipo") = "confirmar" then
		  bdid = "SELECT * FROM localizador WHERE palavrachave LIKE '" & tiraacento(palavrachave) & "' AND publico = 'sim' ORDER BY id DESC"
		  elseif session("tipo") = "admin" AND mostrar = "todos" then
		  bdid = "SELECT * FROM localizador ORDER BY id DESC"
		  else
		  bdid = "SELECT * FROM localizador WHERE palavrachave LIKE '" & tiraacento(palavrachave) & "' ORDER BY id DESC"
		  end if
		  bd.open bdid, conexao
		  while not bd.EOF
		  %>
			<p class="texto">&#8226; <img src="imagens/icone_localizador.gif" width="14" height="14" border="0" /> <a href="localizador.asp?id=<%=bd("id")%>"><%=bd("titulo")%></a><%editarexcluir("localizador")%></p>
		  <%
		  qtdresultado = qtdresultado + 1
		  bd.MoveNext
		  wend
		  bd.close
		  if mostrar = "todos" then
			  if qtdresultado = 0 then
			  %><p class="texto"><strong>Não foi encontrado nenhum registro!</strong></p><%
			  elseif qtdresultado = 1 then
			  %><p class="texto"><strong>Foi encontrado <%=qtdresultado%> registro</strong></p><%
			  else
			  %><p class="texto"><strong>Foram encontrados <%=qtdresultado%> resgistros</strong></p><%
			  end if
		  else
			  if qtdresultado = 0 then
			  %><p class="texto"><strong>Nenhuma palavra parecida com "<%=palavrachave%>" foi encontrada!</strong></p><%
			  elseif qtdresultado = 1 then
			  %><p class="texto"><strong>Foi encontrado <%=qtdresultado%> palavra parecida com "<%=palavrachave%>"</strong></p><%
			  else
			  %><p class="texto"><strong>Foram encontrados <%=qtdresultado%> palavras parecidas com "<%=palavrachave%>"</strong></p><%
			  end if
		  end if
		  end if
	else
		if (idexiste(id,"localizador") = "sim") then
			if session("tipo") = "" OR session("tipo") = "confirmar" then
			bdid = "SELECT * FROM localizador LEFT JOIN cadastro ON localizador.idusuarioadd = cadastro.id WHERE localizador.id=" & id & " AND publico = 'sim'"
			else
			bdid = "SELECT * FROM localizador LEFT JOIN cadastro ON localizador.idusuarioadd = cadastro.id WHERE localizador.id=" & id
			end if
			bd.open bdid, conexao
			if bd.EOF then
				Response.Redirect("login.asp?url=" & Request.ServerVariables("URL") & "%3F" & Request.ServerVariables("QUERY_STRING"))
			else
				%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="localizador.asp">Localizador</a> &nbsp;&gt;&nbsp; <%=bd("titulo")%></p>
				<div id="conteudoeditarlocalizador">
				<p class="titulo"><%=bd("titulo")%></p>
				<p class="texto"><div class='texto'><%=substituirtags(bd("conteudo"))%></div></p>
				</div>
				<%
				if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
					%><div id="formeditarlocalizador" style="display:none">
					<form id="localizador" name="localizador" method="post" action="editar.asp?tipo=localizador" onsubmit="return validacamposbranco('localizador');">
					<p>
					  <label><label class="texto">Palavra - Chave:</label><br />
					  <input name="palavrachave" type="text" class="texto" id="palavrachave" value="<%=bd("palavrachave")%>" size="50" maxlength="50" />
					  </label>
					  <label class="texto">
					  <input type="checkbox" name="publico" id="publico" value="sim"<%if bd("publico") = "sim" then%> checked="checked"<%end if%> /> - publico <a href="javascript:void(0);" onclick="abrirajuda('localizadorpublico');">(?)</a>
					  </label>
					</p>
					<p id="infopalavrachave" class="texto" style="display: none"><strong>digite uma palavrachave</strong></p>
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
					<input type="hidden" name="idlocalizador" id="idlocalizador" value="<%=bd("localizador.id")%>" />
					<p align="center">
						<label>
						<input type="submit" name="add" id="add" value="Atualizar" class="texton" />
						</label>
						<label>
						<input type="button" name="cancelar" id="cancelar" value="Cancelar" class="texton" onclick="editar('localizador','cancelar');validacamposbranco('localizador');" />
						</label>
					</p>
					</form>
					</div>
					<div id="menueditarlocalizador" align="center"><p class="texto"><strong>(<a href="javascript:void(0);" onclick="editar('localizador','editar');">Editar</a>|<a href="javascript:void(0);" onclick="excluir('localizador','<%=bd("localizador.id")%>');">Excluir</a>)</strong></p></div><%
					if (Request.QueryString("acao") = "editar") then
					%><script>editar('localizador','editar');</script><%
					end if
				end if
			end if
		%><p class="texto"><a href="localizador.asp">&laquo; voltar para localizador</a> | <a href="javascript:void(0);" onclick="imprimir('localizador','<%=bd("localizador.id")%>');">Imprimir</a> | <%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%></p><%
		else
		idnaoexiste("localizador")
		end if
	end if
  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>