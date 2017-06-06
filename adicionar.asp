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
  if(session("tipo") = "aluno" OR session("tipo") = "admin") then
  else
  Response.Redirect("login.asp")
  end if
  %>
  <%
  Dim tipo
  
  tipo = Request.QueryString("tipo")
  
  Select Case tipo
  
  Case "agenda"
  
  if (request.form("submitform") = "sim" AND request.form("titulo") <> "" AND request.form("conteudoa") <> "" AND request.form("dia") <> "" AND request.form("mes") <> "")then
  adicionaragenda
  end if
  
  %>
    <form id="agenda" name="agenda" method="post" action="" onsubmit="return validacamposbranco('agenda');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="agenda.asp">Agenda</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Agenda</p>
    <p>
      <label>
      <select name="dia" id="dia" class="texton">
        <option value="" selected="selected">Dia</option>
        <%
		Dim diaagenda
	  	for diaagenda = 1 to 31
	  	%><option value="<%=diaagenda%>"><%=diaagenda%></option>
        <%
		next
	  	%></select>
      </label>
      &nbsp;
      <label>
      <select name="mes" id="mes" class="texton">
        <option value="" selected="selected">Mês</option>
        <%
        Dim mesagenda
        for mesagenda = 1 to 12
        %><option value="<%=mesagenda%>"><%=ucase(left(monthname(mesagenda, false), 1)) & lcase(right(monthname(mesagenda, false), len(monthname(mesagenda, false)) - 1))%></option>
        <%
        next
        %></select>
      </label>
    </p>
    <p id="infodia" class="texto" style="display: none"><strong>selecione um dia</strong></p>
    <p id="infomes" class="texto" style="display: none"><strong>selecione um mês</strong></p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="texton" id="titulo" value="" size="40" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Conteúdo:</label><br />
      <textarea name="conteudoa" id="conteudoa" cols="50" rows="5" class="texto"></textarea>
      </label>
    </p>
    <p id="infoconteudoa" class="texto" style="display: none"><strong>digite algum conteúdo</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    </form>
    <script>document.getElementById('dia').focus();</script>
  <%
  
  Case "comunicado"
  
  if (request.form("submitform") = "sim" AND request.form("titulo") <> "" AND request.form("conteudoa") <> "")then
  adicionarcomunicado
  end if
  
  %>
    <form id="comunicado" name="comunicado" method="post" action="" onsubmit="return validacamposbranco('comunicado');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="comunicados.asp">Comunicados</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Comunicado</p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="titulo" id="titulo" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Conteúdo:</label><br />
      <textarea name="conteudoa" id="conteudoa" cols="75" rows="20" class="texto"></textarea>
      </label>
    </p>
    <p id="infoconteudoa" class="texto" style="display: none"><strong>digite algum conteúdo</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    <label>
    <input type="button" name="formatacao" id="formatacao" value="Formatação" class="texton" onclick="abrirformatacao();" />
    </label>
    </form>
    <script>document.getElementById('titulo').focus();</script>
  <%
  
  Case "caderno"
  
  if (request.form("submitform") = "sim" AND request.form("materia") <> "" AND request.form("titulo") <> "" AND request.form("conteudoa") <> "")then
  adicionarcaderno
  end if
  
  %>
    <form id="caderno" name="caderno" method="post" action="" onsubmit="return validacamposbranco('caderno');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="caderno.asp">Caderno</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Caderno</p>
    <p>
      <label>
      <select name="materia" id="materia" class="texton">
        <option value="" selected="selected">Matéria</option>
        <option value="Biologia">Biologia</option>
        <option value="Geografia">Geografia</option>
        <option value="Outros">Outros</option>
        <option value="Sociologia">Sociologia</option>
      </select>
      </label>
    </p>
    <p id="infomateria" class="texto" style="display: none"><strong>escolha uma matéria</strong></p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="titulo" id="titulo" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Conteúdo:</label><br />
      <textarea name="conteudoa" id="conteudoa" cols="75" rows="20" class="texto"></textarea>
      </label>
    </p>
    <p id="infoconteudoa" class="texto" style="display: none"><strong>digite algum conteúdo</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    <label>
    <input type="button" name="formatacao" id="formatacao" value="Formatação" class="texton" onclick="abrirformatacao();" />
    </label>
    </form>
    <script>document.getElementById('materia').focus();</script>
  <%
  
  Case "galeriafoto"
  
  if (request.form("submitform") = "sim" AND request.form("titulogaleria") <> "")then
  adicionargaleriafoto
  end if
  
  %>
    <form id="galeriaafoto" name="galeriaafoto" method="post" action="" onsubmit="return validacamposbranco('galeriafoto');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Galeria de Foto</p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulogaleria" type="text" class="titulo" id="titulogaleria" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulogaleria" class="texto" style="display: none"><strong>digite um título</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    </form>
    <script>document.getElementById('titulogaleria').focus();</script>
  <%
  
  Case "foto"
  
	Dim idgaleria
	idgaleria = Request.QueryString("idgaleria")
	Dim idexistegaleria
	idexistegaleria = "nao"
	
	if IsNumeric(id) then
	bdid = "SELECT * FROM galeriafoto"
	bd.open bdid, conexao
	while not bd.EOF
	if(bd("id") = Abs(idgaleria)) then
	idexistegaleria = "sim"
	end if
	bd.MoveNext
	wend
	bd.close
	end if
  %>
  <%
  if (idexistegaleria = "sim") then
  %>
<script>
function validaextensao()
{
	bSubmit = true;
	fotopreenchida = "";
	fotosembranco = "";
	for(contador=1;contador<=10;contador++)
	{
		enderecodafoto = document.getElementById('urlfoto' + contador).value;
		if (enderecodafoto == "") {
		fotopreenchida = "";
		document.getElementById('infofotoinvalida' + contador).style.display="none";
		}else{
		fotopreenchida = "ok";
		fotosembranco = "falso";
		}
		if (fotopreenchida == "ok") {
			var extensao;
			extensao = enderecodafoto.split("/");
			extensao = extensao[ (extensao.length-1) ].split(".")
			extensao = extensao[ (extensao.length-1) ];
			extensao = extensao.toLowerCase();
			if (extensao != "jpg" && extensao != "gif" && extensao != "png") {
			document.getElementById('infofotoinvalida' + contador).style.display="block";
			bSubmit = false;
			}else{
			document.getElementById('infofotoinvalida' + contador).style.display="none";
			document.getElementById('carregandofotos').style.display="block";
			}
		}
	}
	if (fotosembranco == "falso") {
	document.getElementById('infofotobranco').style.display="none";
	}else{
	document.getElementById('infofotobranco').style.display="block";
	bSubmit = false;
	}
	return bSubmit
}
exibidos = 1
function selecionarmaisfotos()
{
	for(contador=exibidos;contador<=exibidos+3;contador++)
	{
	document.getElementById('foto' + contador).style.display="block";
	}
	exibidos = exibidos + 3
	if (exibidos >= 10) {
	document.getElementById('selecionarmais').style.display="none";
	}
}
</script>
  <form id="enviafoto" name="enviafoto" method="post" action="adicionar.asp?tipo=enviafoto&idgaleria=<%=idgaleria%>" enctype="multipart/form-data" onsubmit="return validaextensao();">
    <%
	expSQL = "SELECT * FROM galeriafoto WHERE id=" & idgaleria
	bd.open expSQL, conexao
	%>
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; <a href="foto.asp?idgaleria=<%=bd("id")%>"><%=bd("titulo")%></a> &nbsp;&gt;&nbsp; Adicionar</p>
    <%
	bd.close
	%>
    <p class="titulo">Adicionar Foto</p>
    <label class="texto">Envie somente imagens JPG, GIF ou PNG e menores que 1 MB.</label>
    <p>
        <%dim i
		for i = 1 to 10
		%><div id="foto<%=i%>" style="display: none"><input type="file" name="urlfoto<%=i%>" class="texton" size="50"></div>
		<label id="infofotoinvalida<%=i%>" class="texto" style="display: none"><strong>imagem inválida - envie um arquivo com extensão JPG, GIF ou PNG</strong></label>
		<%next%>
        <label id="infofotobranco" class="texto" style="display: none"><strong>selecione uma foto</strong></label>
        <label id="selecionarmais" class="texto"><a href="javascript:void(0);" onclick="selecionarmaisfotos();">selecionar mais</a></label>
    </p>
    <div id="carregandofotos" class="texto" style="display: none"><img src="imagens/icone_carregando.gif" border="0" /> carregando ...</div>
    <label>
    <input type="submit" name="add" id="add" value="Enviar fotos" class="texton" />
    </label>
    <p class="texto">Tamanho ideal: 800 x 600 - Evite adicionar fotos menores ou maiores para não distorcer ou demorar para carregar.</p>
  </form>
<script>document.getElementById('foto1').style.display="block";</script>
  <%
  end if
  
  Case "enviafoto"
  
	idgaleria = Request.QueryString("idgaleria")
	idexistegaleria = "nao"
	
	if IsNumeric(id) then
	bdid = "SELECT * FROM galeriafoto"
	bd.open bdid, conexao
	while not bd.EOF
	if(bd("id") = Abs(idgaleria)) then
	idexistegaleria = "sim"
	end if
	bd.MoveNext
	wend
	bd.close
	end if
	
  if(idexistegaleria = "sim") then
	
	On Error Resume Next
	Dim objUpload
	Set objUpload = Server.CreateObject("Dundas.Upload.2")
	objUpload.UseVirtualDir = True
	objUpload.UseUniqueNames = True
	objUpload.MaxFileSize = 1048576
	Path = ("fotos\")
	objUpload.Save (path)
	
	validacao = "ok"
	
	if( Err.Description = "Uploading file size limit exceeded.") then
	%><p class="texto">Selecione uma foto menor! O limite permitido é de 1 MB.</p><%
	validacao = "falso"
	else
	For Each objUploadedFile in objUpload.Files
	formato = right(lcase(objUploadedFile.Path), 3)
	if formato = "jpg" OR formato = "gif" OR formato = "png" then
	else
	%><p class="texto">Selecione uma foto válida! (JPG, GIF, PNG)</p><%
	validacao = "falso"
	end if
	Next
	end if
	if (validacao = "falso") then
		%><input type="button" value="Voltar" class="texton" onclick="javascript:history.back(-1)" /><%
	else
	For Each objUploadedFile in objUpload.Files
	idfotoadicionado = "fotos/" & Replace(objUploadedFile.Path, "C:\arquivos\sites\primeiroasp\fotos\", "")
	adicionarfoto
	Next
	%>
    <form id="foto" name="foto" method="post" action="adicionar.asp?tipo=atualizartitulofoto">
    <p class="titulo">Fotos Adicionadas</p>
    <%
	contador = 1
	For Each objUploadedFile in objUpload.Files
	linkdafoto = "fotos/" & Replace(objUploadedFile.Path, "C:\arquivos\sites\primeiroasp\fotos\", "")
	%><p>
      <label class="texto">
      <img src="<%=linkdafoto%>" width="100" height="100" border="0" />
      Título:
      <input name="titulo<%=contador%>" type="text" class="texton" id="titulo<%=contador%>" value="" size="50" maxlength="50" />
      <input type="hidden" name="conteudo<%=contador%>" id="conteudo<%=contador%>" value="<%=linkdafoto%>" />
      </label>
    </p>
    <%
	contador = contador + 1
	Next
	%>
    <input type="hidden" name="totaldefotos" id="totaldefotos" value="<%=contador-1%>" />
    <input type="hidden" name="idgaleria" id="idgaleria" value="<%=idgaleria%>" />
    <label>
    <input type="submit" name="salvaralteracoes" id="salvaralteracoes" value="Salvar alterações" class="texton" />
    </label>
    </form>
	<%
	end if
	
	Set objUpload = Nothing
  
  end if
  
  Case "atualizartitulofoto"

	totaldefotos = Request.Form("totaldefotos")
	for i = 1 to totaldefotos
	expSQL = "update foto set titulo ='" & request.form("titulo" & i) & "' where conteudo='" & request.form("conteudo" & i) & "'"
	conexao.execute (expSQL)
	next
	
	Response.Redirect("foto.asp?idgaleria=" & Request.Form("idgaleria"))

  Case "video"
  
  if (request.form("submitform") = "sim" AND request.form("titulo") <> "" AND request.form("idvideo") <> "")then
  adicionarvideo
  end if
  
  %>
    <form id="video" name="video" method="post" action="" onsubmit="return validacamposbranco('video');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="galeria.asp">Fotos e Vídeos</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Vídeo</p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="titulo" id="titulo" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Link do Vídeo You Tube:</label><br />
      <input name="idvideo" type="text" id="idvideo" value="" size="45" class="texton" />
      </label>
    </p>
    <p id="infoidvideo" class="texto" style="display: none"><strong>digite um link</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    <p class="texto">Ex de link: http://br.youtube.com/watch?v=ID_DO_VÍDEO ou http://www.youtube.com/watch?v=ID_DO_VÍDEO. Ou se preferir, coloque somente a ID.</p>
    </form>
    <script>document.getElementById('titulo').focus();</script>
  <%
  
  Case "localizador"
  
  if (request.form("submitform") = "sim" AND request.form("palavrachave") <> "" AND request.form("titulo") <> "" AND request.form("conteudoa") <> "")then
  adicionarlocalizador
  end if
  
  %>
    <form id="localizador" name="localizador" method="post" action="" onsubmit="return validacamposbranco('localizador');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="localizador.asp">Localizador</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Localizador</p>
    <p>
      <label><label class="texto">Palavra-Chave:</label><br />
      <input name="palavrachave" type="text" class="texto" id="palavrachave" value="" size="50" maxlength="50" />
      </label>
      <label class="texto">
      <input type="checkbox" name="publico" id="publico" value="sim" checked="checked" /> - publico <a href="javascript:void(0);" onclick="abrirajuda('localizadorpublico');">(?)</a>
      </label>
    </p>
    <p id="infopalavrachave" class="texto" style="display: none"><strong>digite uma palavrachave</strong></p>
    <p>
      <label><label class="texto">Título:</label><br />
      <input name="titulo" type="text" class="titulo" id="titulo" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Conteúdo:</label><br />
      <textarea name="conteudoa" id="conteudoa" cols="75" rows="20" class="texto"></textarea>
      </label>
    </p>
    <p id="infoconteudoa" class="texto" style="display: none"><strong>digite algum conteúdo</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    <label>
    <input type="button" name="formatacao" id="formatacao" value="Formatação" class="texton" onclick="abrirformatacao();" />
    </label>
    </form>
    <script>document.getElementById('palavrachave').focus();</script>
  <%
  
  Case "links"
  
  if(session("tipo") = "admin") then
  
  if (request.form("submitform") = "sim" AND request.form("titulo") <> "" AND request.form("site") <> "")then
  adicionarlinks
  end if
  
  %>
    <form id="links" name="links" method="post" action="" onsubmit="return validacamposbranco('links');">
    <p class="texto"><a href="default.asp">Início</a> &nbsp;&gt;&nbsp; <a href="default.asp">Links</a> &nbsp;&gt;&nbsp; Adicionar</p>
    <p class="titulo">Adicionar Link no Menu</p>
    <p>
      <label><label class="texto">Título do Link:</label><br />
      <input name="titulo" type="text" class="texto" id="titulo" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infotitulo" class="texto" style="display: none"><strong>digite um título</strong></p>
    <p>
      <label><label class="texto">Site:</label><br />
      <input name="site" type="text" class="texto" id="site" value="" size="50" maxlength="50" />
      </label>
    </p>
    <p id="infosite" class="texto" style="display: none"><strong>digite um site</strong></p>
    <input type="hidden" name="submitform" id="submitform" value="sim" />
    <label>
    <input type="submit" name="add" id="add" value="Adicionar" class="texton" />
    </label>
    </form>
    <script>document.getElementById('titulo').focus();</script>
  
  <%
  
  end if
  
  Case else
  
  End select
  
  Sub adicionaragenda()
  
  bd.Open "agenda",conexao,3,3
  bd.AddNew
  bd("tipo") = "agenda"
  bd("dia") = Request.Form("dia")
  bd("mes") = Request.Form("mes")
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
  
  Response.Redirect("agenda.asp")
  
  End Sub
  
  Sub adicionarcomunicado()
  
  bd.Open "comunicados",conexao,3,3
  bd.AddNew
  bd("tipo") = "comunicados"
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
  
  Response.Redirect("comunicados.asp")
  
  End Sub
  
  Sub adicionarcaderno()
  
  bd.Open "caderno",conexao,3,3
  bd.AddNew
  bd("tipo") = "caderno"
  bd("materia") = Request.Form("materia")
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
  
  Response.Redirect("caderno.asp")
  
  End Sub
  
  Sub adicionargaleriafoto()
  
  bd.Open "galeriafoto",conexao,3,3
  bd.AddNew
  bd("tipo") = "foto"
  bd("titulo") = Request.Form("titulogaleria")
  bd("idusuarioadd") = session("idusuario")
  bd("dataadd") = date()
  bd("horaadd") = time()
  bd.Update
  bd.Close
  conexao.Close
  Set bd = Nothing
  Set conexao = Nothing
  
  Response.Redirect("galeria.asp")
  
  End Sub
  
  Sub adicionarfoto()
  
  bd.Open "foto",conexao,3,3
  bd.AddNew
  bd("tipo") = "foto"
  bd("idgaleria") = idgaleria
  bd("conteudo") = idfotoadicionado
  bd("idusuarioadd") = session("idusuario")
  bd("dataadd") = date()
  bd("horaadd") = time()
  bd.Update
  bd.Close
  
  End Sub
  
  Sub adicionarvideo()
  
  bd.Open "video",conexao,3,3
  bd.AddNew
  bd("tipo") = "video"
  bd("titulo") = Request.Form("titulo")
  bd("conteudo") = Request.Form("idvideo")
  bd("idusuarioadd") = session("idusuario")
  bd("dataadd") = date()
  bd("horaadd") = time()
  bd.Update
  bd.Close
  conexao.Close
  Set bd = Nothing
  Set conexao = Nothing
  
  Response.Redirect("galeria.asp")
  
  End Sub
  
  Sub adicionarlocalizador()
  
  bd.Open "localizador",conexao,3,3
  bd.AddNew
  bd("palavrachave") = LCase(tiraacento(Request.Form("palavrachave")))
  bd("publico") = Request.Form("publico")
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
  
  Response.Redirect("localizador.asp")
  
  End Sub
  
  Sub adicionarlinks()
  
  bd.Open "links",conexao,3,3
  bd.AddNew
  bd("titulo") = Request.Form("titulo")
  bd("site") = Request.Form("site")
  bd.Update
  bd.Close
  conexao.Close
  Set bd = Nothing
  Set conexao = Nothing
  
  Response.Redirect("default.asp")
  
  End Sub
  %>
  <!-- InstanceEndEditable -->
  </div>
  <div id="fim">1º ano A - Ensino Médio - Colégio Dominus Vivendi - 2008</div>
</div>
</body>
<!-- InstanceEnd --></html>