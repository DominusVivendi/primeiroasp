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
  Dim id, tipo
  id = Request.QueryString("id")
  tipo = Request.QueryString("tipo")
  
  Select Case tipo
  
  Case "agenda"
  
  if(verpermissao(tipo) = "sim") then
  
  if (request.form("atualizar") = "Atualizar") then
  if (request.form("dia")<>"Dia" and request.form("mes")<>"Mês" and request.form("titulo")<>"" and request.form("conteudoa")<>"")then
  atualizaragenda
  else
  titulo = request.form("titulo")
  conteudo = request.form("conteudoa")
  dia = request.form("dia")
  mes = request.form("mes")
  validacao = "sim"
  %>
  Por favor, preencha os campos corretamente!
  <%
  end if
  end if
  
  %>
    <form id="agenda" name="agenda" method="post" action="">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="agenda.asp">Agenda</a> &nbsp;&gt;&nbsp; Editar - <%=bd("dia")%>/<%=mesextenso(bd("mes"))%></p>
    <p class="titulo">Editar Agenda</p>
    <p>
      <label>
      <select name="dia" id="dia" class="texton">
        <option value="Dia">Dia</option>
        <%Dim diaagenda
	  for diaagenda = 1 to 31
	  %><option value="<%=diaagenda%>"<%if(bd("dia") = diaagenda) then %> selected<%end if%>><%=diaagenda%></option>
        <%next
	  %></select>
      </label>
      &nbsp;
      <label>
      <select name="mes" id="mes" class="texton">
        <option value="Mês">Mês</option>
        <%Dim mesagenda
		for mesagenda = 1 to 12
		%><option value="<%=mesagenda%>"<%if(bd("mes") = mesagenda) then %> selected<%end if%>><%=mesextenso(mesagenda)%></option>
        <%
		next
		%></select>
      </label>
    </p>
    <p>
      <label>
      <input name="titulo" type="text" class="texton" id="titulo" value="<%
	  if(validacao = "sim") then
	  Response.Write(titulo)
	  else
	  Response.Write(bd("titulo"))
      end if%>" size="40" maxlength="50" />
      </label>
    </p>
    <p>
      <label>
      <textarea name="conteudoa" id="conteudoa" cols="50" rows="5" class="texto"><%
	  if(validacao = "sim") then
	  Response.Write(conteudo)
	  else
	  Response.Write(bd("conteudo"))
      end if%></textarea>
      </label>
    </p>
    <label>
    <input type="submit" name="atualizar" id="atualizar" value="Atualizar" class="texton" />
    </label>
    <label>
    <input type="button" name="voltar" id="voltar" value="Voltar" class="texton" onclick="javascript:history.go(-1);" />
    </label>
    </form>
  <%
  
  end if
  
  Case "comunicados"  
  
  if id <> Empty then
  Response.Redirect(tipo & ".asp?acao=editar&id=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("titulo") <> "" AND request.form("conteudoa") <> "") then
	  
	  atualizarcomunicados
	  
	  else
	  
	  Response.Redirect(tipo & ".asp?id=" & id)
	  
	  end if
  
  end if
  
  Case "caderno"  
  
  if id <> Empty then
  Response.Redirect(tipo & ".asp?acao=editar&id=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("titulo") <> "" AND request.form("conteudoa") <> "") then
	  
	  atualizarcaderno
	  
	  else
	  
	  Response.Redirect(tipo & ".asp?id=" & id)
	  
	  end if
  
  end if
  
  Case "galeriafoto"  
  
  if id <> Empty then
  Response.Redirect("foto.asp?acao=editar&idgaleria=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("titulo") <> "") then
	  
	  atualizargaleriafoto
	  
	  else
	  
	  Response.Redirect("foto.asp?idgaleria=" & id)
	  
	  end if
  
  end if
  
  Case "foto"  
  
  if id <> Empty then
  Response.Redirect(tipo & ".asp?acao=editar&id=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("titulo") <> "") then
	  
	  atualizarfoto
	  
	  else
	  
	  Response.Redirect(tipo & ".asp?id=" & id)
	  
	  end if
  
  end if
  
  Case "video"  
  
  if id <> Empty then
  Response.Redirect(tipo & ".asp?acao=editar&id=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("titulo") <> "") then
	  
	  atualizarvideo
	  
	  else
	  
	  Response.Redirect(tipo & ".asp?id=" & id)
	  
	  end if
  
  end if
  
  Case "localizador"
  
  if id <> Empty then
  Response.Redirect(tipo & ".asp?acao=editar&id=" & id)
  end if
  
  id = Request.Form("id" & tipo)
  
  if(verpermissao(tipo) = "sim") then
  
	  if(request.form("palavrachave") <> "" AND request.form("titulo") <> "" AND request.form("conteudoa") <> "") then
	  
	  atualizarlocalizador
	  
	  else
	  
	  Response.Redirect(tipo & ".asp?id=" & id)
	  
	  end if
  
  end if
  
  Case else
  
  End select

function verpermissao (tipo)
  
  Dim permissao
  
  permissao = "nao"
  
  if(idexiste(id,tipo) = "sim") then
  
  bdid = "SELECT * FROM " & tipo & " WHERE id=" & id
  bd.open bdid, conexao
  
  if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
  permissao = "sim"
  end if
  
  end if
  
  verpermissao = permissao
  
end function

Sub atualizaragenda()

bdid = "update agenda set dia ='" & request.form("dia") & "' where id=" & id

conexao.execute (bdid)

bdid = "update agenda set mes ='" & request.form("mes") & "' where id=" & id

conexao.execute (bdid)

bdid = "update agenda set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

bdid = "update agenda set conteudo ='" & request.form("conteudoa") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("agenda.asp?dia=" & request.form("dia") & "&mes=" & request.form("mes"))

End Sub
  
Sub atualizarcomunicados()

bdid = "update comunicados set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

bdid = "update comunicados set conteudo ='" & request.form("conteudoa") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("comunicados.asp?id="& id)

End Sub

Sub atualizarcaderno()

bdid = "update caderno set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

bdid = "update caderno set conteudo ='" & request.form("conteudoa") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("caderno.asp?id="& id)

End Sub  

Sub atualizargaleriafoto()

bdid = "update galeriafoto set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("foto.asp?idgaleria="& id)

End Sub 

Sub atualizarfoto()

bdid = "update foto set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("foto.asp?id="& id)

End Sub 

Sub atualizarvideo()

bdid = "update video set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("video.asp?id="& id)

End Sub  

Sub atualizarlocalizador()

bdid = "update localizador set palavrachave ='" & tiraacento(request.form("palavrachave")) & "' where id=" & id

conexao.execute (bdid)

bdid = "update localizador set publico ='" & request.form("publico") & "' where id=" & id

conexao.execute (bdid)

bdid = "update localizador set titulo ='" & request.form("titulo") & "' where id=" & id

conexao.execute (bdid)

bdid = "update localizador set conteudo ='" & request.form("conteudoa") & "' where id=" & id

conexao.execute (bdid)

Response.Redirect("localizador.asp?id="& id)

End Sub

  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>