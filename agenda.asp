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
    Dim dia, mes, id
    dia = Request.QueryString("dia")
    mes = Request.QueryString("mes")
    id = Request.QueryString("id")
	origem = Request.QueryString("origem")
	origem = Replace(origem, "%3F","%3F")

	if IsEmpty(dia) AND IsEmpty(mes) then
		%>
        <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Agenda</p>
		<p class="titulo">Agenda</p>
	<%
		if(session("tipo") = "aluno" OR session("tipo") = "admin") then
		%><p class="texto"><a href="adicionar.asp?tipo=agenda"><strong>Novo</strong></a></p><%
		end if
		bdtabela = "SELECT * FROM agenda WHERE mes >= month(date()) ORDER BY mes, dia, id"
		bd.open bdtabela, conexao
		while not bd.EOF
		  if (bd("mes") = month(date) AND bd("dia") >= day(date)) then
		  proximoseventos
		  else
			  if (bd("mes") > month(date)) then
			  proximoseventos
			  end if
		  end if
		bd.MoveNext
		wend
		bd.close
		%>
		<form id="agenda" name="agenda" method="get" action="agenda.asp" onsubmit="return validacamposbranco('agendabuscar');">
		<p class="texto">Ir para a data: 
		  <label>
		  <select name="dia" id="dia" class="texton">
			<option value="" selected="selected">Dia</option>
			<%Dim diaagenda
		  for diaagenda = 1 to 31
		  %><option value="<%=diaagenda%>"><%=diaagenda%></option>
		  <%next
		  %></select>
		  </label>
		  &nbsp;
		  <label>
		  <select name="mes" id="mes" class="texton">
			<option value="" selected="selected">Mês</option>
			<%Dim mesagenda
			for mesagenda = 1 to 12
			%><option value="<%=mesagenda%>"><%=mesextenso(mesagenda)%></option>
		  <%
			next
			%></select>
		  </label>
		  &nbsp;
		  <label>
		<input type="submit" value="Pesquisar" class="texton" />
		</label>
		</p>
		<label id="infodia" class="texto" style="display: none"><strong>selecione um dia</strong></label>
		<label id="infomes" class="texto" style="display: none"><strong>selecione um mês</strong></label>
		</form>
		<%
	else
		if mes >= 1 AND mes <= 12 AND dia >= 1 AND dia <= 31 then
			expSQL = "SELECT * FROM agenda LEFT JOIN cadastro ON agenda.idusuarioadd = cadastro.id WHERE dia=" & dia & " AND mes=" & mes & " ORDER BY agenda.id"
			bd.open expSQL, conexao
			if NOT bd.EOF then
				%>
                <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="agenda.asp">Agenda</a> &nbsp;&gt;&nbsp; <%=dia%>/<%=mesextenso(mes)%></p>
				<p class="titulo"><%=dia%>/<%=mesextenso(mes)%></p>
				<%
				while not bd.EOF
				%>
				<p class="texto">&#8226; <%=bd("titulo")%>:
				<br />
				<%=substituirtagsagenda(bd("conteudo"))%>
				<br />
				<%=mostradatahoraadd(bd("dataadd"),bd("horaadd"))%><%
				if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
				%> - <strong>(<a href="editar.asp?tipo=agenda&id=<%=bd("agenda.id")%>">Editar</a>|<a href="javascript:void(0);" onclick="excluir('agenda','<%=bd("agenda.id")%>');">Excluir</a>)</strong><%
				end if
				%></p>
	<%
				bd.MoveNext
				wend
			else
				%>
                <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="agenda.asp">Agenda</a> &nbsp;&gt;&nbsp; <%=dia%>/<%=mesextenso(mes)%></p>
				<p class="titulo"><%=dia%>/<%=mesextenso(mes)%></p>
				<p class="texto">&#8226; Não existe nenhum evento para este dia!</p>
				<%
			end if
			%><p class="texto"><%
			if origem <> Empty then
			%><a href="<%=origem%>">&laquo; voltar às atualizações</a><%
			else
			%><a href="agenda.asp">&laquo; voltar para agenda</a><%
			end if
			%></p><%
			bd.close
			%>
			<script>function keypresed() {
            if ((window.event ? event.keyCode : event.which) == 37)
            {
				if (<%=dia%> == 1 && <%=mes%> != 1)
				{
				window.location = "agenda.asp?dia=31&mes=<%=mes - 1%>"
				}
				else
				{
					if (<%=dia%> > 1)
					{
					window.location = "agenda.asp?dia=<%=dia - 1%>&mes=<%=mes%>"
					}
				}
            }
            if ((window.event ? event.keyCode : event.which) == 39)
            {
				if (<%=dia%> == 31 && <%=mes%> != 12)
				{
				window.location = "agenda.asp?dia=1&mes=<%=mes + 1%>"
				}
				else
				{
					if (<%=dia%> < 31)
					{
					window.location = "agenda.asp?dia=<%=dia + 1%>&mes=<%=mes%>"
					}
				}
            }
            }
            </script>
            <script>document.onkeydown = keypresed;</script>
			<%
		else
		%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="agenda.asp">Agenda</a> &nbsp;&gt;&nbsp; <%=dia%>/<%=mes%></p>
        <p class="titulo">Dia / Mês inválidos</p>
		<p class="texto">A data (<%=dia%>/<%=mes%>) não existe.</p>
		<p class="texto"><a href="javascript:history.back();">Voltar</a></p>
		<%
		end if
	end if
	if(idexiste(id,"agenda") = "sim") then
		bdid = "SELECT * FROM agenda WHERE id=" & id
		bd.open bdid, conexao
		if origem = Empty then
		Response.Redirect("agenda.asp?dia=" & bd("dia") & "&mes=" & bd("mes"))
		else
		Response.Redirect("agenda.asp?dia=" & bd("dia") & "&mes=" & bd("mes") & "&origem=" & origem)
		end if
		bd.close
	end if

Sub proximoseventos()
%>
    <p class="texto">&#8226; <img src="imagens/icone_agenda.gif" width="14" height="14" border="0" /> <a href="agenda.asp?dia=<%=bd("dia")%>&mes=<%=bd("mes")%>"><%
		if(bd("dia") = day(date) AND bd("mes") = month(date)) then
		%><strong><%
		end if
	  %><%=bd("dia")%>/<%=mesextenso(bd("mes"))%> - <%=bd("titulo")%><%
		if(bd("dia") = day(date) AND bd("mes") = month(date)) then
		%></strong><%
		end if
	  %></a><%editarexcluir("agenda")%></p>
<%
End Sub
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>