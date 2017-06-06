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
    <div id="conteudo_01">
      <p>Comunicados</p>
      <%
      expSQL = "SELECT TOP 5 * FROM comunicados ORDER BY id DESC"
      bd.open expSQL, conexao
      while not bd.EOF
      %>
      <p class="texto"><img src="imagens/icone_comunicados.gif" width="14" height="14" border="0" /> <a href="comunicados.asp?id=<%=bd("id")%>&origem=default.asp"><strong><%=bd("titulo")%></strong></a>
      <br />foi adicionada ao site <%
	if(day(bd("dataadd")) = day(date) AND month(bd("dataadd")) = month(date) AND year(bd("dataadd")) = year(date)) then
		%>às: </strong><%
		if(len(hour(bd("horaadd"))) = 1) then
		Response.Write("0" & hour(bd("horaadd")))
		else
		Response.Write(hour(bd("horaadd")))
		end if
		%>:<%
		if(len(minute(bd("horaadd"))) = 1) then
		Response.Write("0" & minute(bd("horaadd")))
		else
		Response.Write(minute(bd("horaadd")))
		end if
	else
		%>em: </strong><%
		if(len(day(bd("dataadd"))) = 1) then
		Response.Write("0" & day(bd("dataadd")))
		else
		Response.Write(day(bd("dataadd")))
		end if
		Response.Write("/" & mesextenso(month(bd("dataadd"))))
	end if
	  %></p>
      <%
	  bd.MoveNext
      wend
      bd.close
      %>
      <p class="texto" align="right"><a href="comunicados.asp">ver mais &raquo;</a></p>
    </div>
    <div id="conteudo_02">
      <p>Aniversários</p>
      <%
      expSQL = "SELECT * FROM cadastro WHERE mesnasc >= month(date()) AND tipo<>'" & "confirmar" & "'ORDER BY mesnasc ASC, dianasc ASC"
      bd.open expSQL, conexao
	  contadoraniversario = 0
      while not bd.EOF AND contadoraniversario < 7
      if(bd("mesnasc") = month(date) AND bd("dianasc") >= day(date)) then
      aniversarios
	  else
	  if(bd("mesnasc") > month(date)) then
	  aniversarios
	  end if
	  end if
	  bd.MoveNext
      wend
      bd.close
      %>
    </div>
    <div id="conteudo_03">
      <%
      expSQL = "SELECT * FROM agenda WHERE mes >= month(date()) ORDER BY mes, dia, id"
      bd.open expSQL, conexao
	  contadorevento = 0
      while not bd.EOF AND contadorevento < 10
	  if bd("mes") = month(date) AND bd("dia") = day(date) then
	  if contadorevento = 0 then
	  %><p>Hoje</p><%
	  end if
	  proximoseventos
	  elseif bd("mes") = month(date) AND bd("dia") = day(date) + 1 then
	  if amanhadenovo = Empty then
	  %><p>Amanhã</p><%
	  amanhadenovo = 1
	  end if
	  proximoseventos
	  elseif (bd("mes") = month(date) AND bd("dia") > day(date)) then
      if proximoseventosdenovo = Empty then
	  %><p>Próximos Eventos</p><%
	  proximoseventosdenovo = 1
	  end if
	  proximoseventos
	  else
	  if (bd("mes") > month(date)) then
	  if proximoseventosdenovo = Empty then
	  %><p>Próximos Eventos</p><%
	  proximoseventosdenovo = 1
	  end if
	  proximoseventos
	  end if
	  end if
	  bd.MoveNext
      wend
      bd.close
      %>
      <p class="texto" align="right"><a href="agenda.asp">ver mais &raquo;</a></p>
    </div>
    <%
	Sub proximoseventos()
	
	  %>
      <p class="texto"><img src="imagens/icone_agenda.gif" width="14" height="14" border="0" /> <a href="agenda.asp?dia=<%=bd("dia")%>&mes=<%=bd("mes")%>&origem=default.asp"><%
		if(bd("dia") = day(date) AND bd("mes") = month(date)) then
		%><strong><%
		end if
	  %><%=bd("dia")%>/<%=mesextenso(bd("mes"))%> - <%=bd("titulo")%><%
		if(bd("dia") = day(date) AND bd("mes") = month(date)) then
		%></strong><%
		end if
	  %></a></p>
      <%
	  contadorevento = contadorevento + 1
	
	End Sub
	
	Sub aniversarios()
	
      %>
      <p class="texto">
        <%
		if(bd("orkut") <> "")then
		%>
        <a href="http://www.orkut.com/Profile.aspx?uid=<%=bd("orkut")%>" target="_blank"><img src="imagens/icone_orkut.jpg" alt="Orkut de <%=bd("nome")%>" width="10" height="10" border="0" /></a>
        <%
		end if
		if session("idusuario") = Empty then
		%><a href="mensagem.asp?msg=login"><%
		else
		%><a href="usuario.asp?id=<%=bd("id")%>"><%
		end if
		if(bd("dianasc") = day(date) AND bd("mesnasc") = month(date)) then
		%><strong><%
		end if
		if(len(bd("dianasc")) = 1) then
			Response.Write("0" & bd("dianasc"))
			else
			Response.Write(bd("dianasc"))
		end if
		%>/<%=mesextenso(bd("mesnasc"))%> - <%=bd("nome")%><%
		if(bd("dianasc") = day(date) AND bd("mesnasc") = month(date)) then
		%></strong><%
		end if
		%></a></p>
      <%
	  contadoraniversario = contadoraniversario + 1
	
	End Sub
	
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>