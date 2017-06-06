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
  
  excluirdireto
  
  Case "comunicados"
  
  excluirdireto
  
  Case "caderno"
  
  excluirdireto
  
  Case "galeriafoto"
  
  excluirdireto
  
  Case "foto"
  
  excluirdireto
  
  Case "video"
  
  excluirdireto
  
  Case "contato"
  
  if(session("tipo") = "admin") then
  
  	excluirconfirma
  
  end if
  
  Case "localizador"
  
  excluirdireto
  
  Case "links"
  
  if(idexiste(id,tipo) = "sim" AND session("tipo") = "admin") then
	excluir (tipo)  
  end if
  
  Case "comentarios"
  
  excluirdireto
  
  Case "blog"
  
  excluirdireto
  
  Case else
  
  End select
  %>
  <%
  
Sub excluirdireto ()

	if(verpermissao(tipo) = "sim") then
	
		excluir (tipo)
	
	end if

End Sub

Sub excluirconfirma ()

  	if(verpermissao(tipo) = "sim") then
		%>
		<p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="<%=tipo%>.asp"><%=ucase(left(tipo, 1)) & lcase(right(tipo, len(tipo) - 1))%></a> &nbsp;&gt;&nbsp; Excluir</p>
        Tem certeza que deseja excluir: "<%=bd("titulo")%>"?
		<form method="post" action="">
		  <label>
		  <input type="submit" name="botaoexc" id="botaoexc" value="Sim" class="texton" />
		  </label>
		  <label>
		  <input type="button" name="botaoexc" id="botaoexc" value="Não" class="texton" onclick="javascript:history.go(-1);" />
		  </label>
		</form>
		<%
		
		if request.form("botaoexc") = "Sim" then
		excluir (tipo)
		end if
	
	end if

End Sub

Sub excluirgaleriafoto ()

	if(verpermissao(tipo) = "sim") then
		%>
		<p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <a href="foto.asp?idgaleria=<%=id%>"><%=bd("titulo")%></a> &nbsp;&gt;&nbsp; Excluir</p>
        Tem certeza que deseja excluir: "<%=bd("titulo")%>"?
		<form method="post" action="">
		  <label>
		  <input type="submit" name="botaoexc" id="botaoexc" value="Sim" class="texton" />
		  </label>
		  <label>
		  <input type="button" name="botaoexc" id="botaoexc" value="Não" class="texton" onclick="javascript:history.go(-1);" />
		  </label>
		</form>
		<%
		
		if request.form("botaoexc") = "Sim" then			
			expSQL = "delete from foto where idgaleria =" & id
			
			conexao.execute (expSQL)
			
			excluir (tipo)
		end if
	
	end if

End Sub

Sub excluirgaleria ()

	if(verpermissao(tipo) = "sim") then
		if tipo = "foto" then
		bd.close
		expSQL = "SELECT * FROM foto LEFT JOIN galeriafoto ON foto.idgaleria = galeriafoto.id WHERE foto.id=" & id
  		bd.open expSQL, conexao
		%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a>  &nbsp;&gt;&nbsp; <a href="foto.asp?idgaleria=<%=bd("idgaleria")%>"><%=bd("galeriafoto.titulo")%></a> &nbsp;&gt;&nbsp; <a href="foto.asp?id=<%=id%>"><%=bd("foto.titulo")%></a> &nbsp;&gt;&nbsp; Excluir</p><%
		elseif tipo = "video" then
		%><p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <a href="video.asp?id=<%=id%>"><%=bd("titulo")%></a> &nbsp;&gt;&nbsp; Excluir</p><%
		end if
		%>
        Tem certeza que deseja excluir: "<%
		if tipo = "foto" then
		%><%=bd("foto.titulo")%><%
		else
		%><%=bd("titulo")%><%
		end if
		%>"?
		<form method="post" action="">
		  <label>
		  <input type="submit" name="botaoexc" id="botaoexc" value="Sim" class="texton" />
		  </label>
		  <label>
		  <input type="button" name="botaoexc" id="botaoexc" value="Não" class="texton" onclick="javascript:history.go(-1);" />
		  </label>
		</form>
		<%
		
		if request.form("botaoexc") = "Sim" then
		excluir (tipo)
		end if
	
	end if

End Sub

function verpermissao (tipo)
  
  Dim permissao
  
  permissao = "nao"
  
  if(idexiste(id,tipo) = "sim") then
  
  expSQL = "SELECT * FROM " & tipo & " WHERE id=" & id
  bd.open expSQL, conexao
  
  if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
  permissao = "sim"
  end if
  
  end if
  
  verpermissao = permissao
  
end function

function excluir (tipo)

Select Case tipo
Case "foto"
	bd.close
	expSQL = "SELECT * FROM foto WHERE id=" & id
	bd.open expSQL, conexao
	dim idgaleria
	idgaleria = bd("idgaleria")
	bd.close
urlexcluir = "foto.asp?idgaleria=" & idgaleria
Case "galeriafoto"
	expSQL = "delete from foto where idgaleria =" & id
	conexao.execute (expSQL)
urlexcluir = "galeria.asp"
Case "video"
urlexcluir = "galeria.asp"
Case "links"
urlexcluir = "default.asp"
Case Else
urlexcluir = tipo & ".asp"
End Select

expSQL = "delete from " & tipo & " where id =" & id

conexao.execute (expSQL)

Response.Redirect(urlexcluir)

end function  
  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>