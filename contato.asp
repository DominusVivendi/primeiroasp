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
	if(session("idusuario") = "") then
	Response.Redirect("login.asp")
	end if
	
	Dim id, organizar
	id = Request.QueryString("id")
	organizar = Request.QueryString("organizar")

  if(session("tipo") = "admin") then
  if IsEmpty(id) then
  
  if(organizar = "") then
  	organizar = "contato.id DESC"
  end if
  
  %>
  <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Contato</p>
  <p class="titulo">Mensagens</p>
  <table width="470" border="1" cellpadding="0" cellspacing="0">
  <tr>
  <td width="80" class="texto"><div align="center"><strong><a href="contato.asp?organizar=nome">De</a></strong></div></td>
  <td class="texto"><div align="center"><strong><a href="contato.asp?organizar=titulo">Assunto</a></strong></div></td>
  <td width="50" class="texto"><div align="center"><strong><a href="contato.asp">Data</a></strong></div></td>
  <td width="40" class="texto"><div align="center"><strong>Excluir</strong></div></td>
  </tr>
  <%
  expSQL = "SELECT * FROM contato LEFT JOIN cadastro ON contato.idusuarioadd = cadastro.id ORDER BY " & organizar
  bd.open expSQL, conexao
  u=0
  while not bd.EOF
  %>
  <tr>
  <td width="80" class="texto"><%
		if IsNull(bd("cadastro.id")) then
		%>ñ existe<%
		else
		%><a href="usuario.asp?id=<%=bd("cadastro.id")%>"><%=bd("nome")%></a><%
		end if
		%></td>
  <td class="texto"><a href="contato.asp?id=<%=bd("contato.id")%>"><%=bd("titulo")%></a></td>
  <td width="50" class="texto"><%
		if(len(day(bd("dataadd"))) = 1) then
		Response.Write("0" & day(bd("dataadd")))
		else
		Response.Write(day(bd("dataadd")))
		end if
		Response.Write("/" & mesextenso(month(bd("dataadd"))))
  %></td>
  <td width="40" class="texto"><a href="excluir.asp?tipo=contato&id=<%=bd("contato.id")%>">Excluir</a></td>
  </tr>
  <%
  u=u+1
  bd.MoveNext
  wend
  bd.close
  %>
  </table>
  <p>Foram encontrados <%=u%> mensagens.
  <%
  else
	if(idexiste(id,"contato") = "sim") then
		expSQL = "SELECT * FROM contato LEFT JOIN cadastro ON contato.idusuarioadd = cadastro.id WHERE contato.id=" & id
		bd.open expSQL, conexao
	%>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="contato.asp">Contato</a> &nbsp;&gt;&nbsp; Mensagem</p>
    <p class="texto"><strong>Enviado por:</strong> <%
		if IsNull(bd("cadastro.id")) then
		%>ñ existe<%
		else
		%><a href="usuario.asp?id=<%=bd("cadastro.id")%>"><%=bd("nome")%></a> (<%=bd("email")%>)<%
		end if
		%> - <%=day(bd("dataadd"))%>/<%=mesextenso(month(bd("dataadd")))%>/<%=year(bd("dataadd"))%> - <%=bd("horaadd")%></p>
    <p class="texto"><strong>Assunto:</strong> <%=bd("titulo")%></p>
    <p class="texto"><strong>Mensagem:</strong> <div class='texto'><%=substituirtags(bd("conteudo"))%></div></p>
    <p class="texto"><strong><a href="excluir.asp?tipo=contato&id=<%=bd("contato.id")%>">Excluir</a></strong></p>
	<%
	end if
  end if
  else
  if (request.form("visualizar") = "Visualizar") then
  validacao = "sim"
  else
	  if (request.form("titulo") <> "" AND request.form("conteudoa") <> "") then
	  enviarmensagem
	  end if
  end if
  if (validacao = "sim") then
  titulo = request.form("titulo")
  conteudo = request.form("conteudoa")
  else
  titulo = ""
  conteudo = ""
  end if
  
  %>
    <form id="contato" name="contato" method="post" action="" onsubmit="return validacamposbranco('contato');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Contato</p>
    <p class="titulo">Enviar mensagem</p>
    <table width="470" border="0">
      <tr>
        <td width="75" class="texto">De:</td>
        <td class="texto"><%
        expSQL = "SELECT * FROM cadastro WHERE id=" & session("idusuario")
        bd.open expSQL, conexao
        Response.Write(bd("nome"))
        bd.close
        %></td>
      </tr>
      <tr>
        <td width="75" valign="top" class="texto">Assunto:</td>
        <td><label>
          <input name="titulo" type="text" class="texto" id="titulo" value="<%=titulo%>" size="50" maxlength="50" />
          </label><label id="infotitulo" class="texto" style="display: none"><strong>digite um assunto</strong></label></td>
      </tr>
      <tr>
        <td width="75" valign="top" class="texto">Mensagem:</td>
        <td><label>
          <textarea name="conteudoa" id="conteudoa" cols="60" rows="10" class="texto"><%=conteudo%></textarea>
          </label><label id="infoconteudoa" class="texto" style="display: none"><strong>digite sua mensagem</strong></label></td>
      </tr>
      <tr>
        <td width="75" class="texto">&nbsp;</td>
        <td><label>
        <input type="submit" name="botaocontato" id="botaocontato" value="Enviar" class="texton" />
        </label>
        <label>
        <input type="submit" name="visualizar" id="visualizar" value="Visualizar" class="texton" />
        </label>
        <label>
        <input type="button" name="formatacao" id="formatacao" value="Formatação" class="texton" onclick="abrirformatacao();" />
        </label></td>
      </tr>
    </table>
    </form>
    <script>document.getElementById('titulo').focus();</script>
    <%
	if (request.form("visualizar") = "Visualizar") then
		%>
		<p class="texto"><strong>Assunto:</strong> <%=request.form("titulo")%></p>
		<p class="texto"><strong>Mensagem:</strong> <div class='texto'><%=substituirtags(request.form("conteudoa"))%></div></p>
		<%
	end if
	
	end if
	%>
    <%
  
  Sub enviarmensagem()
  
  bd.Open "contato",conexao,3,3
  bd.AddNew
  bd("titulo") = Request.Form("titulo")
  bd("conteudo") = Request.Form("conteudoa")
  bd("idusuarioadd") = session("idusuario")
  bd("dataadd") = date()
  bd("horaadd") = time()
  bd.Update
  bd.Close
  conexao.Close
  Set bd = Nothing
  Set conexao = Nothing
  
  Response.Redirect("mensagem.asp?msg=contatosucesso")
  
  End Sub
	%>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>