<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1º ano A - Formatação</title>
<style type="text/css">
<!--
#formatacao {
	font-family: "Comic Sans MS", "Times New Roman", Verdana, Arial;
	height: 380px;
	width: 310px;
	background-color: #CCCCCC;
	margin: auto;
	padding: 5px;
	border: 5px ridge #FFFFFF;
	overflow: auto;
}
.titulo {
	font-size: 20px;
	text-align: center;
}
.texto {
	font-size: 12px;
}
.texton {
	font-family: "Comic Sans MS", "Times New Roman", Verdana, Arial;
}
a:link {
	color: #000000;
	text-decoration: none;
}
a:visited {
	color: #000000;
	text-decoration: none;
}
a:active {
	color: #000000;
	text-decoration: none;
}
a:hover {
	color: #FF0000;
	text-decoration: none;
}
-->
</style>
</head>

<body>
<%
  if(session("idusuario") = "") then
  %>
  <script>window.close();</script>
  <%
  else
  %>
  <div id="formatacao">
  <p class="titulo">Personalizar conteúdo</p>
  <p class="texto">Digite as tags para obter o resultado personalizado. Não se esqueça de fechar as mesmas.</p>
  <p class="texto">[e] [/e] - Alinhar a esquerda</p>
  <p class="texto">Ex: [e]Texto alinhado a esquerda[/e] = Texto alinhado a esquerda</p>
  <p class="texto">[c] [/c] - Alinhar ao centro (centralizar)</p>
  <p align="center" class="texto">Ex: [c]Texto centralizado[/c] = Texto centralizado</p>
  <p class="texto">[d] [/d] - Alinhar a direita</p>
  <p align="right" class="texto">Ex: [d]Texto alinhado a direita[/d] = Texto alinhado a direita</p>
  <p class="texto">[j] [/j] - Justificar o texto</p>
  <p align="justify" class="texto">Ex: [j]Texto Justificado[/j] = Texto Justificado</p>
  <p class="texto">[i] [/i] - Itálico</p>
  <p class="texto">Ex: [i]Texto em itálico[/i] =<em> Texto em itálico</em></p>
  <p class="texto">[n] [/n] - Negrito</p>
  <p class="texto">Ex: [n]Texto em negrito[/n] =<strong> Texto em negrito</strong></p>
  <p class="texto">[s] [/s] - Sublinhado</p>
  <p class="texto">Ex: [s]Texto sublinhado [/s] = <u>Texto sublinhado</u></p>
  <p class="texto">[l] [/l] [/lt] - Para adicionar links</p>
  <p class="texto">Ex: [l]http://www.google.com.br[/l]Google[/lt] = <a href="http://www.google.com.br" target="_blank">Google</a></p>
  <p class="texto">[img] [/img] - Para adicionar imagens</p>
  <p class="texto">Ex: [img]http://img1.orkut.com/img/i_smile.gif[/img] = <img src="http://img1.orkut.com/img/i_smile.gif" /></p>
  <p class="texto">[vermelho] [/vermelho] - Fonte vermelha</p>
  <p class="texto">Ex: [vermelho]Texto em vermelho[/vermelho] = <font color='red'>Texto em vermelho</font></p>
  <p class="texto">[azul] [/azul] - Fonte azul</p>
  <p class="texto">Ex: [azul]Texto em azul[/azul] = <font color='blue'>Texto em azul</font></p>
  <p class="texto">[verde] [/verde] - Fonte Verde</p>
  <p class="texto">Ex: [verde]Texto em verde[/verde] = <font color='green'>Texto em verde</font></p>
  </div>
  <input type="button" name="fecharformatacao" id="fecharformatacao" value="Fechar" class="texton" onclick="window.close();" />
  <%end if%>
</body>
</html>
