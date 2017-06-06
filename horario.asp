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
	<script>
    function mostralegenda(materia,nprof)
    {
		document.form1.legendahorario.value = materia;
		document.form1.nomeprof.value = nprof;
    }
    </script>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; Horário</p>
    <p class="titulo">Horário</p>
    <table width="450" border="1" align="center" class="tabela">
      <tr>
        <td>&nbsp;</td>
        <td>2ª feira</td>
        <td>3ª feira</td>
        <td>4ª feira</td>
        <td>5ª feira</td>
        <td>6ª feira</td>
      </tr>
      <tr>
        <td class="texto">07:30 ás 08:15</td>
        <td onmouseover="javascript:mostralegenda('B - Biologia','Ana Paula');">B</td>
        <td onmouseover="javascript:mostralegenda('Mg - Matemática Geometria','Jadiel');">Mg</td>
        <td onmouseover="javascript:mostralegenda('Ma - Matemática Álgebra','Jadiel');">Ma</td>
        <td onmouseover="javascript:mostralegenda('H - História','Gastão');">H</td>
        <td onmouseover="javascript:mostralegenda('H - História','Gastão');">H</td>
      </tr>
      <tr>
        <td class="texto">08:15 ás 09:00</td>
        <td onmouseover="javascript:mostralegenda('G - Geografia','Paulo');">G</td>
        <td onmouseover="javascript:mostralegenda('EF - Educação Física','Tony');">EF</td>
        <td onmouseover="javascript:mostralegenda('I - Inglês','Fátima');">I</td>
        <td onmouseover="javascript:mostralegenda('F - Física','Bento');">F</td>
        <td onmouseover="javascript:mostralegenda('RED - Redação','Rose');">RED</td>
      </tr>
      <tr>
        <td class="texto">09:00 ás 09:45</td>
        <td onmouseover="javascript:mostralegenda('H - História','Gastão');">H</td>
        <td onmouseover="javascript:mostralegenda('RH - Relações Humanas','Valter');">RH</td>
        <td onmouseover="javascript:mostralegenda('G - Geografia','Paulo');">G</td>
        <td onmouseover="javascript:mostralegenda('F - Física','Bento');">F</td>
        <td onmouseover="javascript:mostralegenda('Q - Química','Eduardo');">Q</td>
      </tr>
      <tr>
        <td class="texto">09:45 ás 10:15</td>
        <td onmouseover="javascript:mostralegenda('Recreio','');">-</td>
        <td onmouseover="javascript:mostralegenda('Recreio','');">-</td>
        <td onmouseover="javascript:mostralegenda('Recreio','');">-</td>
        <td onmouseover="javascript:mostralegenda('Recreio','');">-</td>
        <td onmouseover="javascript:mostralegenda('Recreio','');">-</td>
      </tr>
      <tr>
        <td class="texto">10:15 ás 11:00</td>
        <td onmouseover="javascript:mostralegenda('LIT - Literatura','Rose');">LIT</td>
        <td onmouseover="javascript:mostralegenda('LIT - Literatura','Rose');">LIT</td>
        <td onmouseover="javascript:mostralegenda('Q - Química','Eduardo');">Q</td>
        <td onmouseover="javascript:mostralegenda('B - Biologia','Ana Paula');">B</td>
        <td onmouseover="javascript:mostralegenda('Ma - Matemática Álgebra','Jadiel');">Ma</td>
      </tr>
      <tr>
        <td class="texto">11:00 ás 11:45</td>
        <td onmouseover="javascript:mostralegenda('Q - Química','Eduardo');">Q</td>
        <td onmouseover="javascript:mostralegenda('B - Biologia','Ana Paula');">B</td>
        <td onmouseover="javascript:mostralegenda('SOC - Sociologia','Valter');">SOC</td>
        <td onmouseover="javascript:mostralegenda('Mg - Matemática Geometria','Jadiel');">Mg</td>
        <td onmouseover="javascript:mostralegenda('G - Geografia','Paulo');">G</td>
      </tr>
      <tr>
        <td class="texto">11:45 ás 12:30</td>
        <td onmouseover="javascript:mostralegenda('Ma - Matemática Álgebra','Jadiel');">Ma</td>
        <td onmouseover="javascript:mostralegenda('GRAM - Gramática','Sávio');">GRAM</td>
        <td onmouseover="javascript:mostralegenda('ESP - Espanhol','Aline');">ESP</td>
        <td onmouseover="javascript:mostralegenda('EA - Educação Artística','Adriana');">EA</td>
        <td onmouseover="javascript:mostralegenda('F - Física','Bento');">F</td>
      </tr>
      <tr>
        <td class="texto">12:30 ás 13:15</td>
        <td onmouseover="javascript:mostralegenda('GRAM - Gramática','Sávio');">GRAM</td>
        <td onmouseover="javascript:mostralegenda('-','');">-</td>
        <td onmouseover="javascript:mostralegenda('-','');">-</td>
        <td onmouseover="javascript:mostralegenda('I - Inglês','Fátima');">I</td>
        <td onmouseover="javascript:mostralegenda('-','');">-</td>
      </tr>
    </table>
    <form id="form1" name="form1" method="post" action="">
      <table width="300" border="0" align="center" class="texto">
        <tr>
          <td><div align="right">Matéria:</div></td>
          <td><label>
          <input name="legendahorario" type="text" id="legendahorario" value="Passe o mouse nas matérias" size="30" style="font-family:'Comic Sans MS', 'Times New Roman', Verdana, Arial; border: 0px; background-color: #CCCCCC" readonly="readonly" />
          </label></td>
        </tr>
        <tr>
          <td><div align="right">Professor(a):</div></td>
          <td><label>
          <input name="nomeprof" type="text" id="nomeprof" style="font-family:'Comic Sans MS', 'Times New Roman', Verdana, Arial;  border: 0px; background-color: #CCCCCC" value="" size="30" readonly="readonly" />
          </label></td>
        </tr>
      </table>
    </form>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>