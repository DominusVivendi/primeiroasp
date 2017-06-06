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
  
  Case "adicionar"
  
  if(request.form("enviardados") = "Cadastrar") then
  verificarcamposbranco
  %><a href="javascript:history.back();">Voltar</a><%
  end if
  
  Case "editar"
  
  Dim editarid
  
  if(session("tipo") = "admin") then
  editarid = Request.QueryString("id")
  else
  editarid = session("idusuario")
  end if
  
  bdusuario = "SELECT * FROM cadastro WHERE id=" & editarid
  bd.open bdusuario, conexao
  
  if(session("idusuario") = bd("id") OR session("tipo") = "admin") then
  
  if(session("tipo") = "admin") then
  tipoeditarcadastro = ""
  else
  tipoeditarcadastro = "editar"
  end if
  
  %>
  <form id="form1" name="form1" method="post" action="cadastro.asp?acao=atualizar<%
  if(session("tipo") = "admin") then
  Response.Write("&id=" & bd("id"))
  end if
  %>" onsubmit="return validacamposbranco('<%=tipoeditarcadastro%>cadastro');">
    <p class="titulo">Editar cadastro</p>
    <table width="450" border="0" align="center" class="texto">
      <tr>
        <td><div align="right">Nome: </div></td>
        <td width="280"><label>
          <input name="nome" type="text" class="texton" id="nome" value="<%=bd("nome")%>" size="40" maxlength="16" onkeyup="validacadastro('nome');" />
        </label></td>
      </tr>
      <tr id="infonome" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um nome</strong></label></td>
      </tr>
      <tr>
      <tr>
        <td><div align="right">Sobrenome:</div></td>
        <td><label>
          <input name="sobrenome" type="text" class="texton" id="sobrenome" value="<%=bd("sobrenome")%>" size="40" maxlength="50" onkeyup="validacadastro('sobrenome');" />
        </label></td>
      </tr>
      <tr id="infosobrenome" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um sobrenome</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Data Nascimento:</div></td>
        <td><label>
      <select name="dianasc" id="dianasc" class="texton" onchange="validacadastro('dianasc');">
        <option value="">Dia</option><%
		Dim dianasc
		for dianasc = 1 to 31%>
        <option value="<%=dianasc%>"<%if(bd("dianasc") = dianasc) then %> selected<%end if%>><%=dianasc%></option><%
		next
		%></select>
      </label>
      <label>
      <select name="mesnasc" id="mesnasc" class="texton" onchange="validacadastro('mesnasc');">
        <option value="">Mês</option><%
		Dim mesnasc
		for mesnasc = 1 to 12%>
        <option value="<%=mesnasc%>"<%if(bd("mesnasc") = mesnasc) then %> selected<%end if%>><%=ucase(left(monthname(mesnasc), 1)) & lcase(right(monthname(mesnasc), len(monthname(mesnasc)) - 1))%></option><%
		next
		%></select>
      </label>
      <label>
      <select name="anonasc" id="anonasc" class="texton" onchange="validacadastro('anonasc');">
        <option value="">Ano</option><%
		Dim anonasc
		for anonasc = 1900 to year(date())%>
        <option value="<%=anonasc%>"<%if(bd("anonasc") = anonasc) then %> selected<%end if%>><%=anonasc%></option><%
		next
		%></select>
      </label></td>
      </tr>
      <tr id="infodatanasc" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>selecione a data de nascimento</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">E-mail:</div></td>
        <td><label>
          <input name="email" type="text" class="texton" id="email" value="<%=bd("email")%>" size="50" maxlength="50" onblur="validacadastro('email');" />
        </label></td>
      </tr>
      <tr id="infoemail" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um email</strong></label></td>
      </tr>
      <tr id="infoemailvalido" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um email válido</strong></label></td>
      </tr>
      <input name="emailatual" id="emailatual" type="hidden" value="<%=bd("email")%>" />
      <%
	  if(session("tipo") = "admin") then
	  %>
      <tr>
        <td width="200"><div align="right">Nova senha:</div></td>
        <td><label>
          <input name="senha" type="password" class="texton" id="senha" value="<%=bd("senha")%>" maxlength="50" onblur="validacadastro('senha');" />
        </label></td>
      </tr>
      <tr id="infosenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite uma senha</strong></label></td>
      </tr>
      <tr>
        <td width="200"><div align="right">Confirma nova senha:</div></td>
        <td><label>
          <input name="confirmasenha" type="password" class="texton" id="confirmasenha" value="<%=bd("senha")%>" maxlength="50" onblur="validacadastro('confirmasenha');" />
        </label></td>
      </tr>
      <tr id="infoconfirmasenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite novamente a sua senha</strong></label></td>
      </tr>
      <tr id="infosenhaconfere" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>as senhas não conferem</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Lembrete Senha:</div></td>
        <td><label>
          <input name="lembretesenha" type="text" class="texton" id="lembretesenha" value="<%=bd("lembretesenha")%>" size="50" maxlength="200" onkeyup="validacadastro('lembretesenha');" />
        </label></td>
      </tr>
      <tr id="infolembretesenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um lembrete de senha</strong></label></td>
      </tr>
      <%end if%>
      <tr>
        <td><div align="right">ID Orkut:</div></td>
        <td><label>
          <input name="orkut" type="text" class="texton" id="orkut" value="<%=bd("orkut")%>" size="50" maxlength="100" />
        </label></td>
      </tr>
      <%
			if(session("tipo") = "admin") then
	  %>
      <tr>
        <td><div align="right">Tipo:</div></td>
        <td><label>
          <select name="tipo" class="texton" id="tipo">
            <option value="Tipo">Tipo</option>
            <option value="admin"<%if(bd("tipo") = "admin") then %> selected<%end if%>>Administrador</option>
            <option value="aluno"<%if(bd("tipo")  = "aluno") then %> selected<%end if%>>Aluno</option>
            <option value="outros"<%if(bd("tipo")  = "outros") then %> selected<%end if%>>Outros</option>
            <option value="confirmar"<%if(bd("tipo")  = "confirmar") then %> selected<%end if%>>Confirmar</option>
          </select>
        </label></td>
      </tr>
      <%end if%>
      <tr>
        <td><div align="right">Nº Chamada:</div></td>
        <td><label>
      <select name="nchamada" id="nchamada" class="texton">
        <option value="Nº Chamada">Nº Chamada</option><%
		Dim nchamada
		for nchamada = 1 to 27%>
        <option value="<%=nchamada%>"<%if(bd("nchamada") = nchamada) then %> selected<%end if%>><%=nchamada%></option><%
		next
		%></select>
      <input name="nchamadaatual" type="hidden" id="nchamadaatual" value="<%=bd("nchamada")%>" />
      </label></td>
      </tr>
      <tr>
        <td><div align="right"></div></td>
        <td><label>
          <input name="enviardados" type="submit" class="texton" id="enviardados" value="Atualizar" />
        </label></td>
      </tr>
    </table>
  </form>
  <%
  
  end if
  
  Case "atualizar"
  
  if(request.form("enviardados") = "Atualizar") then
  verificarcamposbranco
  %><a href="javascript:history.back();">Voltar</a><%
  end if
  
  Case "alterarsenha"
  
  if(session("idusuario") <> "") then
  
  Dim alterarsenha
  alterarsenha = session("idusuario")
  Dim validarsenhas
  validarsenhas = "ok"
  
  bdusuario = "SELECT * FROM cadastro WHERE id=" & alterarsenha
  bd.open bdusuario, conexao
  
  if(session("idusuario") = bd("id") OR session("tipo") = "admin") then
  
  %>
  <form id="form1" name="form1" method="post" action="" onsubmit="return validacamposbranco('alterasenha');">
    <p class="titulo">Alterar Senha</p>
    <table width="450" border="0" align="center" class="texto">
      <tr>
        <td width="180"><div align="right">Senha atual:</div></td>
        <td><label>
          <input name="senhaatual" type="password" class="texton" id="senhaatual" maxlength="50" onblur="validaalterasenha('senhaatual');" />
        </label><%
		if(request.form("enviardados") = "Alterar") then
			if(request.form("senhaatual") = "") then
		  	%><br />
		  	<label class="texto">O campo Senha Atual está em branco!</label>
		  	<%
		  	validarsenhas = "falso"
		  	else
			if(request.form("senhaatual") <> bd("senha")) then
		  	%><br />
		  	<label class="texto">A senha atual não está correta!</label>
	  	  <%
		  	validarsenhas = "falso"
		  	end if
			end if
		  end if
		%></td>
      </tr>
      <tr id="infosenhaatual" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite a sua senha atual</strong></label></td>
      </tr>
      <tr>
        <td width="180"><div align="right">Nova senha:</div></td>
        <td><label>
          <input name="senha" type="password" class="texton" id="senha" maxlength="50" onblur="validaalterasenha('senha');" />
        </label><%
		  if(request.form("enviardados") = "Alterar") then
		  	if(request.form("senha") = "") then
		  	%><br />
		  	<label class="texto">O campo Nova Senha está em branco!</label>
		  	<%
		  	validarsenhas = "falso"
			else
			if(request.form("senha") <> request.form("confirmasenha")) then
		  	%><br />
		  	<label class="texto">As novas senhas não conferem!</label>
	  	  <%
		  	validarsenhas = "falso"
		  	end if
			end if
		  end if
		%></td>
      </tr>
      <tr id="infosenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite uma nova senha</strong></label></td>
      </tr>
      <tr>
        <td width="180"><div align="right">Confirma nova senha:</div></td>
        <td><label>
          <input name="confirmasenha" type="password" class="texton" id="confirmasenha" maxlength="50" onblur="validaalterasenha('confirmasenha');" />
        </label><%
		  if(request.form("enviardados") = "Alterar") then
		  	if(request.form("confirmasenha") = "") then
		  	%><br />
		  	<label class="texto">O campo Confirma Nova Senha está em branco!</label>
	  	  <%
		  	validarsenhas = "falso"
			end if
		  end if
		%></td>
      </tr>
      <tr id="infoconfirmasenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite novamente a sua nova senha</strong></label></td>
      </tr>
      <tr id="infosenhaconfere" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>as novas senhas não conferem</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Lembrete Senha <a href="javascript:void(0);" onclick="abrirajuda('lembretesenha');">(?)</a>:</div></td>
        <td><label>
          <input name="lembretesenha" type="text" class="texton" id="lembretesenha" size="50" maxlength="200" onkeyup="validaalterasenha('lembretesenha');" />
        </label><%
		  if(request.form("enviardados") = "Alterar") then
		  	if(request.form("lembretesenha") = "") then
		  	%><br />
		  	<label class="texto">O campo Lembrete Senha está em branco!</label>
	  	  <%
		  	validarsenhas = "falso"
			end if
		  end if
		%></td>
      </tr>
      <tr id="infolembretesenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um lembrete de senha</strong></label></td>
      </tr>
      <tr>
        <td width="180"><div align="right"></div></td>
        <td><label>
          <input name="enviardados" type="submit" class="texton" id="enviardados" value="Alterar" />
        </label></td>
      </tr>
    </table>
  </form>
  <script>
  document.getElementById('senhaatual').focus();
  </script>
  <%
  
  if(request.form("enviardados") = "Alterar" AND validarsenhas = "ok") then
  
  	atualizarsenha
  
  end if
  
  end if
  
  end if
  
  Case "excluir"
  
  excluirid = Request.QueryString("id")
  
  bdid = "SELECT * FROM cadastro where id=" & excluirid
  bd.open bdid, conexao
  
  if(session("tipo") = "admin") then
  
  if request.form("botaoexc") <> "" then
  if request.form("botaoexc") = "Sim" then
  excluircadastro
  else
  Response.Redirect("cadastro.asp")
  end if
  end if
  
  %>
  Tem certeza que deseja excluir o cadastro de <%=bd("nome")%>?
  <form id="form3" name="form3" method="post" action="">
    <label>
      <input type="submit" name="botaoexc" id="botaoexc" value="Sim" class="texton" />
    </label>
    <label>
      <input type="submit" name="botaoexc" id="botaoexc" value="Não" class="texton" />
    </label>
  </form>
  <%
  
  end if
  
  Case else
  
  if(session("idusuario") <> "") then
  if(session("tipo") = "admin") then
  
  Dim organizar
  
  organizar = Request.QueryString("organizar")
  
  if(organizar = "") then
  	organizar = "id"
  end if
  
  %>
  <p class="titulo">Cadastros</p>
  <table width="470" border="1" cellpadding="0" cellspacing="0">
  <tr>
  <td width="100" class="texto"><div align="center"><strong><a href="cadastro.asp?organizar=nome">Nome</a></strong></div></td>
  <td class="texto"><div align="center"><strong><a href="cadastro.asp?organizar=email">E-mail</a></strong></div></td>
  <td width="50" class="texto"><div align="center"><strong><a href="cadastro.asp?organizar=tipo">Tipo</a></strong></div></td>
  <td width="40" class="texto"><div align="center"><strong>Editar</strong></div></td>
  <td width="40" class="texto"><div align="center"><strong>Excluir</strong></div></td>
  </tr>
  <%
  
  bdusuario = "SELECT * FROM cadastro ORDER BY " & organizar
  bd.open bdusuario, conexao
  u=0
  while not bd.EOF
  %>
  <tr>
  <td width="100" class="texto"><a href="usuario.asp?id=<%=bd("id")%>"><%=bd("nome")%></a></td>
  <td class="texto"><a href="usuario.asp?id=<%=bd("id")%>"><%=bd("email")%></a></td>
  <td width="50" class="texto"><%=bd("tipo")%></td>
  <td width="40" class="texto"><a href="cadastro.asp?acao=editar&id=<%=bd("id")%>">Editar</a></td>
  <td width="40" class="texto"><a href="cadastro.asp?acao=excluir&id=<%=bd("id")%>">Excluir</a></td>
  </tr>
  <%
  u=u+1
  bd.MoveNext
  wend
  bd.close
  
  %>
  </table>
  <p>Foram encontrados <%=u%> cadastros.
  <%
  
  else
  
  Response.Redirect("usuario.asp")
  
  end if
  
  else
  
  %>
  <form id="form1" name="form1" method="post" action="cadastro.asp?acao=adicionar" onsubmit="return validacamposbranco('cadastro');">
    <p class="titulo">Cadastro</p>
    <table width="450" border="0" align="center" class="texto">
      <tr>
        <td><div align="right">Nome: </div></td>
        <td width="280"><label>
          <input name="nome" type="text" class="texton" id="nome" size="40" maxlength="16" onkeyup="validacadastro('nome');" />
        </label></td>
      </tr>
      <tr id="infonome" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um nome</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Sobrenome:</div></td>
        <td><label>
          <input name="sobrenome" type="text" class="texton" id="sobrenome" size="40" maxlength="50" onkeyup="validacadastro('sobrenome');" />
        </label></td>
      </tr>
      <tr id="infosobrenome" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um sobrenome</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Data Nascimento:</div></td>
        <td><label>
      <select name="dianasc" id="dianasc" class="texton" onchange="validacadastro('dianasc');" onblur="validacadastro('dianasc');">
        <option value="" selected="selected">Dia</option><%
		Dim dianasca
		for dianasca = 1 to 31%>
        <option value="<%=dianasca%>"><%=dianasca%></option><%
		next
		%></select>
      </label>
      <label>
      <select name="mesnasc" id="mesnasc" class="texton" onchange="validacadastro('mesnasc');" onblur="validacadastro('mesnasc');">
        <option value="" selected="selected">Mês</option><%
		Dim mesnasca
		for mesnasca = 1 to 12%>
        <option value="<%=mesnasca%>"><%=ucase(left(monthname(mesnasca), 1)) & lcase(right(monthname(mesnasca), len(monthname(mesnasca)) - 1))%></option><%
		next
		%></select>
      </label>
      <label>
      <select name="anonasc" id="anonasc" class="texton" onchange="validacadastro('anonasc');" onblur="validacadastro('anonasc');">
        <option value="" selected="selected">Ano</option><%
		Dim anonasca
		for anonasca = 1900 to year(date())%>
        <option value="<%=anonasca%>"><%=anonasca%></option><%
		next
		%></select>
      </label></td>
      </tr>
      <tr id="infodatanasc" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>selecione a data de nascimento</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">E-mail <a href="javascript:void(0);" onclick="abrirajuda('email');">(?)</a>:</div></td>
        <td><label>
          <input name="email" type="text" class="texton" id="email" size="50" maxlength="50" onblur="validacadastro('email');" />
        </label></td>
      </tr>
      <tr id="infoemail" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um email</strong></label></td>
      </tr>
      <tr id="infoemailvalido" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um email válido</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Senha <a href="javascript:void(0);" onclick="abrirajuda('senha');">(?)</a>:</div></td>
        <td><label>
          <input name="senha" type="password" class="texton" id="senha" maxlength="50" onblur="validacadastro('senha');" />
        </label></td>
      </tr>
      <tr id="infosenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite uma senha</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Confirma Senha <a href="javascript:void(0);" onclick="abrirajuda('confirmasenha');">(?)</a>:</div></td>
        <td><label>
          <input name="confirmasenha" type="password" class="texton" id="confirmasenha" maxlength="50" onblur="validacadastro('confirmasenha');" />
        </label></td>
      </tr>
      <tr id="infoconfirmasenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite novamente a sua senha</strong></label></td>
      </tr>
      <tr id="infosenhaconfere" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>as senhas não conferem</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">Lembrete Senha <a href="javascript:void(0);" onclick="abrirajuda('lembretesenha');">(?)</a>:</div></td>
        <td><label>
          <input name="lembretesenha" type="text" class="texton" id="lembretesenha" size="50" maxlength="200" onkeyup="validacadastro('lembretesenha');" />
        </label></td>
      </tr>
      <tr id="infolembretesenha" style="display: none">
        <td></td>
        <td width="280"><label class="texto"><strong>digite um lembrete de senha</strong></label></td>
      </tr>
      <tr>
        <td><div align="right">ID Orkut <a href="javascript:void(0);" onclick="abrirajuda('idorkut');">(?)</a>:</div></td>
        <td><label>
          <input name="orkut" type="text" class="texton" id="orkut" size="50" maxlength="100" />
        </label></td>
      </tr>
      <tr>
        <td><div align="right">Nº Chamada <a href="javascript:void(0);" onclick="abrirajuda('nchamada');">(?)</a>:</div></td>
        <td><label>
      <select name="nchamada" id="nchamada" class="texton">
        <option selected="selected">Nº Chamada</option><%
		Dim nchamadaa
		for nchamadaa = 1 to 27%>
        <option value="<%=nchamadaa%>"><%=nchamadaa%></option><%
		next
		%></select>
      </label></td>
      </tr>
      <tr>
        <td><div align="right"></div></td>
        <td><label>
          <input name="enviardados" type="submit" class="texton" id="enviardados" value="Cadastrar" />
        </label></td>
      </tr>
    </table>
  </form>
  <script>
  document.getElementById('nome').focus();
  </script>
  <%
  end if
  
  End select
  
  Sub verificarcamposbranco()
  
  Dim camposbranco
  camposbranco = "ok"
  if(request.form("nome")="")then
  %>
  <p class="texto">O campo Nome está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("sobrenome")="")then
  %>
  <p class="texto">O campo Sobrenome está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("dianasc") = "" AND request.form("mesnasc") = "" AND request.form("anonasc") = "") then
  %>
  <p class="texto">Os campos da Data de Nascimento não estão selecionados!</p>
  <%
  camposbranco = "falso"
  else
  if(request.form("dianasc")="")then
  %>
  <p class="texto">O campo Dia não está selecionado!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("mesnasc")="")then
  %>
  <p class="texto">O campo Mês não está selecionado!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("anonasc")="")then
  %>
  <p class="texto">O campo Ano não está selecionado!</p>
  <%
  camposbranco = "falso"
  end if
  end if
  if(request.form("email")="")then
  %>
  <p class="texto">O campo E-mail está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("enviardados") = "Cadastrar" OR session("tipo") = "admin") then
  if(request.form("senha")="")then
  %>
  <p class="texto">O campo Senha está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("confirmasenha")="")then
  %>
  <p class="texto">O campo Confirma Senha está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  if(request.form("lembretesenha")="")then
  %>
  <p class="texto">O campo Lembrete de Senha está em branco!</p>
  <%
  camposbranco = "falso"
  end if
  end if
  if(session("tipo") = "admin") then
  if(request.form("tipo")="Tipo")then
  %>
  <p class="texto">O campo Tipo não está selecionado!</p>
  <%
  camposbranco = "falso"
  end if
  end if
  if(session("tipo") = "admin" AND request.form("tipo") = "aluno")then
  if(request.form("nchamada") = "Nº Chamada" OR request.form("nchamada") = "0")then
  %>
  <p class="texto">O campo Nº Chamada não está selecionado!</p>
  <%
  camposbranco = "falso"
  end if
  end if
  if(camposbranco = "ok") then
  verificaemail
  end if
  
  End Sub
  
  Sub verificaemail()
  
  If EmailValido(Trim(Request.Form("email"))) = True Then
  verificaemailexiste
  Else
	%>
	<p class="texto">O email digitado é inválido.</p>
	<%
  End If
  
  End Sub
  
  Function EmailValido(email)
  Set objRegExp = New RegExp
  objRegExp.Pattern = "^[a-z0-9._-]+\@[a-z0-9._-]+\.[a-z]{2,4}$"
  objRegExp.IgnoreCase = True
  EmailValido = objRegExp.Test(email)
  End Function
  
  Sub verificaemailexiste()
  
	  if(request.form("enviardados") = "Cadastrar" OR request.form("enviardados") = "Atualizar") then
	  if(request.form("email") = request.form("emailatual")) then
	  verificaidorkut
	  end if
	  Dim emailexiste
	  emailexiste = "nao"
		bdtabela = "SELECT * FROM cadastro"
		bd.open bdtabela, conexao
		while not bd.EOF
	  if(request.form("email") = bd("email"))then
	  	emailexiste = "sim"
	  end if
		bd.MoveNext
		wend
		bd.close
	  if(emailexiste = "sim") then
	  %>
	  <p class="texto">O E-mail digitado já existe! Por favor verifique se você já não é cadastrado no site ou cadastre-se com outro e-mail.</p>
	  <%
	  end if
		  if(emailexiste = "nao") then
		  	verificaidorkut
		  end if
	  end if
  
  End Sub
  
  Sub verificaidorkut()
  
  if IsNumeric(pegaidorkut(request.form("orkut"))) OR pegaidorkut(request.form("orkut")) = "" then
  if(Len(pegaidorkut(request.form("orkut"))) < 50) then
  verificanchamadaexiste
  else
  %>
  <p class="texto">O ID Orkut não é válido! Possui mais de 50 números.</p>
  <%
  end if
  else
  %>
  <p class="texto">O ID Orkut não é válido! Caso não esteja conseguindo colocar o seu ID Orkut através do link do seu perfil inteiro, copie somente os números que aparecem no final do endereço do seu perfil.</p>
  <%
  end if
    
  End Sub
  
  Sub verificanchamadaexiste()
  
  if(request.form("nchamada") <> "Nº Chamada") then
    if(request.form("enviardados") = "Atualizar" AND Abs(request.form("nchamada")) = Abs(request.form("nchamadaatual"))) then
	verificasenhaconfere
	else
  Dim nchamadaexiste
  nchamadaexiste = "nao"
    bdtabela = "SELECT * FROM cadastro WHERE tipo='"& "aluno" &"'ORDER BY nchamada ASC"
	bd.open bdtabela, conexao
	while not bd.EOF
  if(Abs(request.form("nchamada")) = Abs(bd("nchamada")))then
  nchamadaexiste = "sim"
  end if
    bd.MoveNext
	wend
	bd.close
  if(nchamadaexiste = "sim") then
  %>
  <p class="texto">O número de chamada escolhido já possui um registro!
  <br />
  Confira o seu nº de chamada, e caso esteja certo, comunique ao administrador do site para que o mesmo possa ser arrumado!</p>
  <%
  else
  verificasenhaconfere
  end if
  end if
  else
  verificasenhaconfere
  end if
  
  End Sub
  
  Sub verificasenhaconfere()
  
  if(request.form("senha") <> request.form("confirmasenha")) then
  %>
  <p class="texto">As senhas não conferem!</p>
  <%
  else
  vertipocadastro
  end if
  
  End Sub
  
  Sub vertipocadastro()
  
  if(request.form("enviardados") = "Cadastrar") then
  novocadastro
  end if
  if(request.form("enviardados") = "Atualizar") then
  atualizar
  end if
  
  End Sub
  
  Sub novocadastro()
  
  bd.Open "cadastro",conexao,3,3
  bd.AddNew
  bd("nome") = trim(Request.Form("nome"))
  bd("sobrenome") = trim(Request.Form("sobrenome"))
  bd("dianasc") = Request.Form("dianasc")
  bd("mesnasc") = Request.Form("mesnasc")
  bd("anonasc") = Request.Form("anonasc")
  bd("email") = trim(Request.Form("email"))
  bd("senha") = Request.Form("senha")
  bd("lembretesenha") = trim(Request.Form("lembretesenha"))
  bd("orkut") = pegaidorkut(request.form("orkut"))
  bd("tipo") = "confirmar"
  if(Request.Form("nchamada") = "Nº Chamada") then
  bd("nchamada") = "0"
  else
  bd("nchamada") = Request.Form("nchamada")
  end if
  bd.Update
  bd.Close
  conexao.Close
  Set bd = Nothing
  Set conexao = Nothing
  
  Response.Redirect("mensagem.asp?msg=cadastrosucesso")
  
  End Sub
  
  Sub atualizar()
  
  if(session("tipo") = "admin") then
  editarid = Request.QueryString("id")
  else
  editarid = session("idusuario")
  end if
  campoembranco = "0"
  
  bdid = "update cadastro set nome ='" & trim(request.form("nome")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set sobrenome ='" &trim(request.form("sobrenome")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set dianasc ='" & Server.HTMLEncode(request.form("dianasc")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set mesnasc ='" & Server.HTMLEncode(request.form("mesnasc")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set anonasc ='" & Server.HTMLEncode(request.form("anonasc")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set email ='" & trim(request.form("email")) & "' where id=" & editarid
  bd.Open bdid, conexao
  if(session("tipo") = "admin") then
  bdid = "update cadastro set senha ='" & Server.HTMLEncode(request.form("senha")) & "' where id=" & editarid
  bd.Open bdid, conexao
  bdid = "update cadastro set lembretesenha ='" & trim(request.form("lembretesenha")) & "' where id=" & editarid
  bd.Open bdid, conexao
  end if
  bdid = "update cadastro set orkut ='" & pegaidorkut(request.form("orkut")) & "' where id=" & editarid
  bd.Open bdid, conexao
  if(session("tipo") = "admin") then
  bdid = "update cadastro set tipo ='" & Server.HTMLEncode(request.form("tipo")) & "' where id=" & editarid
  bd.Open bdid, conexao
  end if
  if(request.form("nchamada") = "Nº Chamada") then
  bdid = "update cadastro set nchamada ='" & campoembranco & "' where id=" & editarid
  bd.Open bdid, conexao
  else
  bdid = "update cadastro set nchamada ='" & Server.HTMLEncode(request.form("nchamada")) & "' where id=" & editarid
  bd.Open bdid, conexao
  end if
  
  Response.Redirect("mensagem.asp?msg=atualizarcadastrosucesso")
  
  End Sub
  
  Sub atualizarsenha()
  
  atualizarsenhaa = session("idusuario")
  
  bdid = "update cadastro set senha ='" & Server.HTMLEncode(request.form("senha")) & "' where id=" & atualizarsenhaa
  
  conexao.execute (bdid)
  
  bdid = "update cadastro set lembretesenha ='" & trim(request.form("lembretesenha")) & "' where id=" & atualizarsenhaa
  
  conexao.execute (bdid)
  
  Response.Redirect("mensagem.asp?msg=atualizarsenhasucesso")
  
  End Sub
  
  Sub excluircadastro()
  
  excluirid = Request.QueryString("id")
  
  bdid = "delete from cadastro where id =" & excluirid
  
  conexao.execute (bdid)
  
  Response.Redirect("cadastro.asp")
  
  End Sub
  
  Function pegaidorkut (linkorkut)
  linkorkut = Replace(linkorkut, "http://www.orkut.com/Profile.aspx?uid=","")
  linkorkut = Replace(linkorkut, "http://www.orkut.com.br/Profile.aspx?uid=","")
  linkorkut = Replace(linkorkut, "http://www.orkut.com.br/Main#Profile.aspx?uid=","")
  linkorkut = Replace(linkorkut, "http://www.orkut.com/Main#Profile.aspx?uid=","")
  pegaidorkut = linkorkut
  End Function
  
  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>