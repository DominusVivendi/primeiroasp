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
  Dim acao
  
  acao = Request.QueryString("acao")
  
  Select Case acao
  
  Case "logout"
  
  Session.Abandon
  Response.Redirect("default.asp")
  
  Case "admin"
  
  if session("idusuario") = 7 then
	  Session("tipo") = "admin"
	  Response.Redirect("default.asp")
  end if
  
  Case else
  
  if(session("idusuario") <> "") then
  Response.Redirect("default.asp")
  else
	email = Request("email")
	senha = Request("senha")
	email = Replace(email, "'", "") 'Proteção contra SQL Injection
	senha = Replace(senha, "'", "")
  %>
  <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Login</p>
  <form id="form1" name="form1" method="post" action="login.asp" onsubmit="return validacamposbranco('login');">
    <table width="400" border="0" align="center">
      <tr>
        <td height="50" colspan="2"><div align="center">Login 1º A</div></td>
      </tr>
      <tr>
        <td width="100"><div align="right" class="texto">E-mail:</div></td>
        <td><input name="email" type="text" class="texton" id="email" value="<%=Request("email")%>" /></td>
      </tr>
      <tr id="infoemail" style="display: none">
        <td></td>
        <td><label class="texto"><strong>digite um email</strong></label></td>
      </tr>
      <%
	  if(request.form("verificarlogin") = "Login" AND email = "") then
	  %>
      <tr>
        <td width="100">&nbsp;</td>
        <td><label class="texto" style="color:#FF0000">Campo obrigatório. Não pode ficar em branco.</label><script>document.form1.email.focus();</script></td>
      </tr>
	  <%
	  end if
	  %>
      <tr>
        <td width="100"><div align="right" class="texto">Senha:</div></td>
        <td><input name="senha" type="password" class="texton" id="senha" /></td>
      </tr>
      <tr id="infosenha" style="display: none">
        <td></td>
        <td><label class="texto"><strong>digite uma senha</strong></label></td>
      </tr>
      <%
	  if(request.form("verificarlogin") = "Login") then
		if(senha = "") then
		%>
      <tr>
        <td width="100">&nbsp;</td>
        <td><label class="texto" style="color:#FF0000">O campo senha não pode ficar em branco!</label></td>
      </tr>
	  	<%
		else
		if(email <> "" AND senha <> "") then
		expSQL="SELECT * from cadastro where email='" & email & "' AND senha='" & senha & "'"
		Set bd = conexao.Execute(expSQL)
		if bd.EOF then
		%>
      <tr>
        <td width="100">&nbsp;</td>
        <td><label class="texto"><strong>e-mail e/ou senha estão incorretos</strong></label><script>document.form1.senha.focus();</script></td>
      </tr>
	  	<%
		else
		  Session("idusuario") = bd("id")
		  Session("tipo") = bd("tipo")
		  if bd("lembretesenha") <> "" then
			if Request.Form("url") = Empty then
			Response.Redirect("default.asp")
			else
			Response.Redirect(Request.Form("url"))
			end if
		  else
		  Response.Redirect("mensagem.asp?msg=faltalembretesenha")
		  end if
		end if
		end if
		end if
		if(email <> "" AND senha = "") then
		%><script>document.form1.senha.focus();</script><%
		end if
	  end if
	  %>
      <tr>
        <td width="100"><div align="right"></div></td>
        <td><input name="verificarlogin" type="submit" class="texton" id="verificarlogin" value="Login" /></td>
      </tr>
      <tr id="infoesquecisenha">
        <td></td>
        <td><label class="texto"><a href="esqueciminhasenha.asp">Esqueci minha senha</a> | <a href="cadastro.asp">Cadastrar</a></label></td>
      </tr>
    </table>
    <input type="hidden" name="url" value="<%
		if(request.form("verificarlogin") = "Login") then
		Response.Write(Request("url"))
		else
			if Request.QueryString("url") = Empty then
			Response.Write(Request.ServerVariables("HTTP_REFERER"))
			else
			Response.Write(Request.QueryString("url"))
			end if
		end if
		%>" />
  </form>
  <%
	if(email = "" AND senha = "") then
	%><script>document.form1.email.focus();</script><%
	end if
	
  end if
  
  End select
  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>