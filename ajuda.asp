<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1º ano A - Ajuda</title>
<style type="text/css">
<!--
#ajuda {
	font-family: "Comic Sans MS", "Times New Roman", Verdana, Arial;
	height: 150px;
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
<div id="ajuda">
<%
	dim tipo
	tipo = Request.QueryString("tipo")
	
	Select Case tipo
	
	Case "email"
	
	%>
	<p class="titulo">E-mail</p>
	<p class="texto">Digite um email válido pois o mesmo será usado para o login no site e para a comunicação entre o site e o usuário.</p>
	<%
	
	Case "senha"
	
	%>
	<p class="titulo">Senha</p>
	<p class="texto">Digite uma senha que será usada para fazer o login no site.</p>
	<%
	
	Case "confirmasenha"
	
	%>
	<p class="titulo">Confirma Senha</p>
	<p class="texto">Digite novamente a sua senha para confirmação.</p>
	<%
	
	Case "lembretesenha"
	
	%>
	<p class="titulo">Lembrete de Senha</p>
	<p class="texto">O Lembrete de Senha é uma frase ou algum texto que te lembre da sua senha, caso venha se esquecer.</p>
	<%
	
	Case "idorkut"
	
	%>
	<p class="titulo">ID Orkut</p>
	<p class="texto">O ID Orkut é a sua identificação do Orkut. Caso queria que ele apareça no site do 1º ano A, coloque a ID. Campo não obrigatório.
    <br />Para saber qual é o seu ID Orkut, acesse seu perfil do Orkut e copie o link do seu perfil que o site irá pegar somente a ID do seu Orkut.</p>
	<%
	
	Case "nchamada"
	
	%>
	<p class="titulo">Número de Chamada</p>
	<p class="texto">O número de chamada somente os alunos necessitam preencher. Caso você não seja aluno do 1º ano A - Colégio Dominus Vivendi deixe este campo em branco.</p>
	<%
	
	Case "localizadorpublico"
	
	%>
	<p class="titulo">Localizador - Público</p>
	<p class="texto">Tornar a página localizador público, quer dizer que todas as pessoas poderão ver esta página. Caso deseje que somente as pessoas logadas no site vejam esta página, não selecione esta caixa.</p>
	<%
	
	Case else
	
	End Select
%>
</div>
</body>
</html>
