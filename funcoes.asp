<%
Function substituirtags (textoconteudo)
textoconteudo = Replace(textoconteudo, "<","&#8249;")
textoconteudo = Replace(textoconteudo, ">","&#8250;")
textoconteudo = Replace(textoconteudo, "[c]","</div><div align='center' class='texto'>")
textoconteudo = Replace(textoconteudo, "[/c]","</div><div class='texto'>")
textoconteudo = Replace(textoconteudo, "[e]","</div><div align='left' class='texto'>")
textoconteudo = Replace(textoconteudo, "[/e]","</div><div class='texto'>")
textoconteudo = Replace(textoconteudo, "[d]","</div><div align='right' class='texto'>")
textoconteudo = Replace(textoconteudo, "[/d]","</div><div class='texto'>")
textoconteudo = Replace(textoconteudo, "[j]","</div><div align='justify' class='texto'>")
textoconteudo = Replace(textoconteudo, "[/j]","</div><div class='texto'>")
textoconteudo = Replace(textoconteudo, "[i]","<em>")
textoconteudo = Replace(textoconteudo, "[/i]","</em>")
textoconteudo = Replace(textoconteudo, "[n]","<strong>")
textoconteudo = Replace(textoconteudo, "[/n]","</strong>")
textoconteudo = Replace(textoconteudo, "[s]","<u>")
textoconteudo = Replace(textoconteudo, "[/s]","</u>")
textoconteudo = Replace(textoconteudo, "[l]","<a target='_blank' href='")
textoconteudo = Replace(textoconteudo, "[/l]","'>")
textoconteudo = Replace(textoconteudo, "[/lt]","</a>")
textoconteudo = Replace(textoconteudo, "[img]","<img src='")
textoconteudo = Replace(textoconteudo, "[/img]","' />")
textoconteudo = Replace(textoconteudo, "[vermelho]","<font color='red'>")
textoconteudo = Replace(textoconteudo, "[/vermelho]","</font>")
textoconteudo = Replace(textoconteudo, "[azul]","<font color='blue'>")
textoconteudo = Replace(textoconteudo, "[/azul]","</font>")
textoconteudo = Replace(textoconteudo, "[verde]","<font color='green'>")
textoconteudo = Replace(textoconteudo, "[/verde]","</font>")
textoconteudo = Replace(textoconteudo, chr(13),"<br />")
substituirtags = textoconteudo
End function

Function substituirtagsagenda (textoconteudo)
textoconteudo = Replace(textoconteudo, "<","&#8249;")
textoconteudo = Replace(textoconteudo, ">","&#8250;")
textoconteudo = Replace(textoconteudo, "[i]","<em>")
textoconteudo = Replace(textoconteudo, "[/i]","</em>")
textoconteudo = Replace(textoconteudo, "[n]","<strong>")
textoconteudo = Replace(textoconteudo, "[/n]","</strong>")
textoconteudo = Replace(textoconteudo, "[s]","<u>")
textoconteudo = Replace(textoconteudo, "[/s]","</u>")
textoconteudo = Replace(textoconteudo, "[l]","<a target='_blank' href='")
textoconteudo = Replace(textoconteudo, "[/l]","'>")
textoconteudo = Replace(textoconteudo, "[/lt]","</a>")
textoconteudo = Replace(textoconteudo, "[img]","<img src='")
textoconteudo = Replace(textoconteudo, "[/img]","' />")
textoconteudo = Replace(textoconteudo, "[vermelho]","<font color='red'>")
textoconteudo = Replace(textoconteudo, "[/vermelho]","</font>")
textoconteudo = Replace(textoconteudo, "[azul]","<font color='blue'>")
textoconteudo = Replace(textoconteudo, "[/azul]","</font>")
textoconteudo = Replace(textoconteudo, "[verde]","<font color='green'>")
textoconteudo = Replace(textoconteudo, "[/verde]","</font>")
textoconteudo = Replace(textoconteudo, chr(13),"<br />")
substituirtagsagenda = textoconteudo
End function

Function tiraacento (palavra)
palavra = LCase(palavra)
palavra = Replace(palavra, "á","a")
palavra = Replace(palavra, "à","a")
palavra = Replace(palavra, "â","a")
palavra = Replace(palavra, "ã","a")
palavra = Replace(palavra, "ä","a")

palavra = Replace(palavra, "é","e")
palavra = Replace(palavra, "è","e")
palavra = Replace(palavra, "ê","e")
palavra = Replace(palavra, "ë","e")

palavra = Replace(palavra, "í","i")
palavra = Replace(palavra, "ì","i")
palavra = Replace(palavra, "î","i")
palavra = Replace(palavra, "ï","i")
palavra = Replace(palavra, "ý","y")

palavra = Replace(palavra, "ó","o")
palavra = Replace(palavra, "ò","o")
palavra = Replace(palavra, "ô","o")
palavra = Replace(palavra, "õ","o")
palavra = Replace(palavra, "ö","o")

palavra = Replace(palavra, "ú","u")
palavra = Replace(palavra, "ù","u")
palavra = Replace(palavra, "û","u")
palavra = Replace(palavra, "ü","u")

palavra = Replace(palavra, "ç","c")

palavra = Replace(palavra, "ñ","n")

tiraacento = palavra
End function

Function mostradatahoraadd (dataadd,horaadd)
%><strong>adicionado <%
if(day(dataadd) = day(date) AND month(dataadd) = month(date) AND year(dataadd) = year(date)) then
	%>às: </strong><%
	if(len(hour(horaadd)) = 1) then
	Response.Write("0" & hour(horaadd))
	else
	Response.Write(hour(horaadd))
	end if
	%>:<%
	if(len(minute(horaadd)) = 1) then
	Response.Write("0" & minute(horaadd))
	else
	Response.Write(minute(horaadd))
	end if
else
	%>em: </strong><%
	if(len(day(dataadd)) = 1) then
	Response.Write("0" & day(dataadd))
	else
	Response.Write(day(dataadd))
	end if
	Response.Write("/" & mesextenso(month(bd("dataadd"))))
end if
%> <strong>por:</strong> <%
	if IsNull(bd("cadastro.id")) then
	%>ñ existe<%
	elseif session("idusuario") = Empty then
	%><%=bd("nome")%><%
	else
	%><a href="usuario.asp?id=<%=bd("cadastro.id")%>"><%=bd("nome")%></a><%
	end if
End Function

Function idexiste (id,tipo)
if IsNumeric(id) AND NOT IsEmpty(id) then
	expSQL = "SELECT * FROM " & tipo & " WHERE id=" & id
	bd.open expSQL, conexao
	if NOT bd.EOF then
		idexiste = "sim"
	end if
	bd.close
end if
End Function

Function editarexcluir (tipo)
if(session("idusuario") = bd("idusuarioadd") OR session("tipo") = "admin") then
	%> - <strong>(<a href="editar.asp?tipo=<%=tipo%>&id=<%=bd("id")%>">Editar</a>|<a href="javascript:void(0);" onclick="excluir('<%=tipo%>','<%=bd("id")%>');">Excluir</a>)</strong><%
end if
End Function

Function idnaoexiste (tipo)
%>
<p class="titulo"><%=ucase(left(tipo, 1)) & lcase(right(tipo, len(tipo) - 1))%></p>
<p class="texto">Id incorreta ou inválida!</p>
<p class="texto">A id que você solicitou parece não ser uma id válida, ou ela pode ter sido excluida do banco de dados.</p>
<p class="texto">Se você chegou a esta página através de um link do site, por favor, avise o fato ocorrido que iremos verificar o erro.</p>
<p class="texto">Obrigado.</p>
<%
End Function

Function mesextenso (mes)
mesextenso = ucase(left(monthname(mes, false), 1)) & lcase(right(monthname(mes, false), len(monthname(mes)) - 1))
End Function
%>