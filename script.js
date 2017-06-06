// JavaScript Document
function abrirformatacao ()
{
	window.open('formatacao.asp','','height=450,width=350,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no');
}
function abrirajuda (tipo)
{
	window.open('ajuda.asp?tipo=' + tipo,'','height=200,width=350,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no');
}
function imprimir (tipo,id)
{
	window.open('imprimir.asp?tipo=' + tipo + '&id=' + id,'','height=450,width=470,toolbar=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,copyhistory=no');
}
function abrirslides (id)
{
	window.open('apresentacaodeslides.asp?acao=iniciar&id=' + id,'','height=670,width=820,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,copyhistory=no');
}
function excluir (tipo,id)
{
	switch (tipo)
	{
		case "agenda": msg = "este evento";
		break;
		case "comunicados": msg = "este comunicado";
		break;
		case "caderno": msg = "este caderno";
		break;
		case "foto": msg = "esta foto";
		break;
		case "galeriafoto": msg = "esta galeria de foto";
		break;
		case "video": msg = "este video";
		break;
		case "localizador": msg = "este localizador";
		break;
		case "links": msg = "este link";
		break;
		default: msg = "";
	}
	if (confirm('Tem certeza de que deseja excluir ' + msg + '?'))
	{
		window.location = 'excluir.asp?tipo=' + tipo + '&id=' + id;
	}
}
function editar (tipo,acao)
{
	if(acao == "editar")
	{
		document.getElementById('formeditar' + tipo).style.display="block";
		document.getElementById('conteudoeditar' + tipo).style.display="none";
		document.getElementById('menueditar' + tipo).style.display="none";
		document.getElementById('titulo').focus();
	}
	if(acao == "cancelar")
	{
		document.getElementById('formeditar' + tipo).style.display="none";
		document.getElementById('conteudoeditar' + tipo).style.display="block";
		document.getElementById('menueditar' + tipo).style.display="block";
		document.getElementById(tipo).reset();
	}
}
function validacamposbranco (tipo)
{
	bSubmit = true;
	if(tipo == "agenda")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
		mostraocultamsg('mes');
		mostraocultamsg('dia');
	}
	if(tipo == "comunicado")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
	}
	if(tipo == "caderno")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
		mostraocultamsg('materia');
	}
	if(tipo == "editarcaderno")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
	}
	if(tipo == "galeriafoto")
	{
		mostraocultamsg('titulogaleria');
	}
	//if(tipo == "titulofoto")
	//{
	//	mostraocultamsg('titulo');
	//}
	if(tipo == "video")
	{
		mostraocultamsg('idvideo');
		mostraocultamsg('titulo');
	}
	if(tipo == "localizador")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
		mostraocultamsg('palavrachave');
	}
	if(tipo == "links")
	{
		mostraocultamsg('site');
		mostraocultamsg('titulo');
	}
	if(tipo == "contato")
	{
		mostraocultamsg('conteudoa');
		mostraocultamsg('titulo');
	}
	if(tipo == "login")
	{
		mostraocultamsg('senha');
		mostraocultamsg('email');
	}
	if(tipo == "cadastro")
	{
		validacadastro('senha');
		validacadastro('email');
		mostraocultamsg('lembretesenha');
		mostraocultamsg('confirmasenha');
		mostraocultamsg('senha');
		mostraocultamsg('email');
		validacadastro('anonasc');
		validacadastro('mesnasc');
		validacadastro('dianasc');
		mostraocultamsg('sobrenome');
		mostraocultamsg('nome');
	}
	if(tipo == "editarcadastro")
	{
		validacadastro('email');
		mostraocultamsg('email');
		validacadastro('anonasc');
		validacadastro('mesnasc');
		validacadastro('dianasc');
		mostraocultamsg('sobrenome');
		mostraocultamsg('nome');
	}
	if(tipo == "alterasenha")
	{
		validaalterasenha('senha');
		mostraocultamsg('lembretesenha');
		mostraocultamsg('confirmasenha');
		mostraocultamsg('senha');
		mostraocultamsg('senhaatual');
	}
	if(tipo == "agendabuscar")
	{
		mostraocultamsg('mes');
		mostraocultamsg('dia');
	}
	if(tipo == "buscalocalizador")
	{
		mostraocultamsg('palavrachave');
	}
	return bSubmit
}
function mostraocultamsg (nomecampo)
{
	if(document.getElementById(nomecampo).value == "")
	{
		document.getElementById('info' + nomecampo).style.display="block";
		document.getElementById(nomecampo).focus();
		bSubmit = false;
	}
	else
	{
		document.getElementById('info' + nomecampo).style.display="none";
	}
}
function validacadastro (nomecampo)
{
	if(nomecampo == "dianasc" || nomecampo == "mesnasc" || nomecampo == "anonasc")
	{
		if ((window.event ? event.keyCode : event.which) != 9)
		{
			if(document.getElementById('dianasc').value == "" || document.getElementById('mesnasc').value == "" || document.getElementById('anonasc').value == "")
			{
				document.getElementById('infodatanasc').style.display="block";
				bSubmit = false;
			}
			else
			{
				document.getElementById('infodatanasc').style.display="none";
			}
		}
	}
	else
	{
		if ((window.event ? event.keyCode : event.which) != 9)
		{
			if(document.getElementById(nomecampo).value == "")
			{
				document.getElementById('info' + nomecampo).style.display="block";
				bSubmit = false;
			}
			else
			{
				document.getElementById('info' + nomecampo).style.display="none";
			}
		}
	}
	if(nomecampo == "confirmasenha" || nomecampo == "senha")
	{
		if(document.getElementById('senha').value != "" && document.getElementById('confirmasenha').value != "")
		{
			validasenha();
		}
		else
		{
			document.getElementById('infosenhaconfere').style.display="none";
		}
	}
	if(nomecampo == "email")
	{
		if(document.getElementById('email').value != "")
		{
			validaemail();
		}
		else
		{
			document.getElementById('infoemailvalido').style.display="none";
		}
	}
}
function validasenha ()
{
	if(document.getElementById('senha').value != document.getElementById('confirmasenha').value)
	{
		document.getElementById('infosenhaconfere').style.display="block";
		bSubmit = false;
	}
	else
	{
		document.getElementById('infosenhaconfere').style.display="none";
	}
}
function validaalterasenha (nomecampo)
{
	if(document.getElementById('senha').value != "" && document.getElementById('confirmasenha').value != "")
	{
		validasenha();
	}
	else
	{
		document.getElementById('infosenhaconfere').style.display="none";
	}
	if ((window.event ? event.keyCode : event.which) != 9)
	{
		if(document.getElementById(nomecampo).value == "")
		{
			document.getElementById('info' + nomecampo).style.display="block";
			bSubmit = false;
		}
		else
		{
			document.getElementById('info' + nomecampo).style.display="none";
		}
	}
}
function validaemail (){
	emailerrado = "nao"
    if(document.form1.email.value.indexOf (' ') != -1){
		emailerrado = "sim";
    }
    if(document.form1.email.value.indexOf ('@') < 1){;
		emailerrado = "sim";
    }
    document.form1.email.value.indexOf ('@')
    if(document.form1.email.value.substring((document.form1.email.value.indexOf ('@') + 1), document.form1.email.value.length).indexOf ('@') >= 0){
		emailerrado = "sim";
    }
    //if(document.form1.email.value.indexOf ('.') < 5){
	//	emailerrado = "sim";
    //}
    if((document.form1.email.value.substring((document.form1.email.value.indexOf ('.') + 1), document.form1.email.value.length).length) < 2){
		emailerrado = "sim";
    }
	if(emailerrado == "sim")
	{
		document.getElementById('infoemailvalido').style.display="block";
		document.getElementById('email').focus();
		bSubmit = false;
	}
	else
	{
		document.getElementById('infoemailvalido').style.display="none";
	}
}
function validaemail2() {
	var obj = eval("document.form1.email");
	var txt = obj.value;
	if ((txt.length != 0) && ((txt.indexOf("@") < 1) || (txt.indexOf('.') < 7)))
	{
		document.getElementById('infoemailvalido').style.display="block";
	}
	else
	{
		document.getElementById('infoemailvalido').style.display="none";
	}
}
function validalocalizador (nomecampo)
{
	bSubmit = true;
	if(document.getElementById(nomecampo).value == "")
	{
		document.getElementById('info' + nomecampo).style.display="block";
		document.getElementById('infopalavrachavediferente').style.display="none";
		bSubmit = false;
	}
	else
	{
		document.getElementById('info' + nomecampo).style.display="none";
		if(document.getElementById('palavrachave').value == "%")
		{
			document.getElementById('infopalavrachavediferente').style.display="block";
			bSubmit = false;
		}
		else
		{
			document.getElementById('infopalavrachavediferente').style.display="none";
		}
	}
	return bSubmit
}
function validaesqueceusenha (nomecampo)
{
	bSubmit = true;
	document.getElementById('emailnaoencontrado').style.display="none";
	if(document.getElementById(nomecampo).value == "")
	{
		document.getElementById('info' + nomecampo).style.display="block";
		bSubmit = false;
	}
	else
	{
		document.getElementById('info' + nomecampo).style.display="none";
	}
	return bSubmit
}
function validalogin (nomecampo)
{
	if ((window.event ? event.keyCode : event.which) != 9)
	{
		if(document.getElementById(nomecampo).value == "")
		{
		document.getElementById('info' + nomecampo).style.display="block";
		}
		else
		{
		document.getElementById('info' + nomecampo).style.display="none";
		}
	}
	if(document.getElementById('email').value == "" || document.getElementById('senha').value == "")
	{
		document.form1.verificarlogin.disabled = true;
	}
	else
	{
		document.form1.verificarlogin.disabled = false;
	}
}
function validacadastro2 (acao,nomecampo)
{
	if(nomecampo == "dianasc" || nomecampo == "mesnasc" || nomecampo == "anonasc")
	{
		if ((window.event ? event.keyCode : event.which) != 9)
		{
			if(document.getElementById(nomecampo).value == "")
			{
				document.getElementById('infodatanasc').style.display="block";
			}
			else
			{
				document.getElementById('infodatanasc').style.display="none";
			}
		}
	}
	else
	{
		if ((window.event ? event.keyCode : event.which) != 9)
		{
			if(document.getElementById(nomecampo).value == "")
			{
				document.getElementById('info' + nomecampo).style.display="block";
			}
			else
			{
				document.getElementById('info' + nomecampo).style.display="none";
			}
		}
	}
	if(nomecampo == "confirmasenha" || nomecampo == "senha")
	{
		if(document.getElementById('senha').value != "" && document.getElementById('confirmasenha').value != "")
		{
			validasenha();
		}
		else
		{
			document.getElementById('infosenhaconfere').style.display="none";
		}
	}
	if(nomecampo == "email" && document.getElementById('email').value != "")
	{
		validaemail();
	}
	else
	{
		document.getElementById('infoemailvalido').style.display="none";
	}
	if(acao == "cadastro")
	{
		if(document.getElementById('nome').value == "" || document.getElementById('sobrenome').value == "" || document.getElementById('senha').value == "" || document.getElementById('confirmasenha').value == "" || document.getElementById('dianasc').value == "" || document.getElementById('mesnasc').value == "" || document.getElementById('anonasc').value == "" || document.getElementById('email').value == "" || emailerrado == "sim" || senhaerrada == "sim")
		{
			document.form1.enviardados.disabled = true;
		}
		else
		{
			document.form1.enviardados.disabled = false;
		}
	}
	if(acao == "editar")
	{
		if(document.getElementById('nome').value == "" || document.getElementById('sobrenome').value == "" || document.getElementById('dianasc').value == "" || document.getElementById('mesnasc').value == "" || document.getElementById('anonasc').value == "" || document.getElementById('email').value == "" || emailerrado == "sim")
		{
			document.form1.enviardados.disabled = true;
		}
		else
		{
			document.form1.enviardados.disabled = false;
		}
	}
}