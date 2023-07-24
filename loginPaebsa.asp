 <% Response.Buffer = True%> 
 <% Server.ScriptTimeOut = 1000000000 %>
           <!--#include file="conexPaebsa.asp"-->
		   <!--#include file="validaFechas.asp"-->
		   <!--#include file="fecha.asp"-->
		   <!--#include file="functions.asp"-->
		   <!--#include file="Fun_Fechas.asp"-->

 
            <%
				if session("usuario") = "" then response.redirect "Cerrar_Ses_Cli.asp"
				if session("contrasena") = "" then response.redirect "Cerrar_Ses_Cli.asp"
				if session("nombre") = "" then response.redirect "Cerrar_Ses_Cli.asp"
 
				SQL = "SELECT Sesion_Activa FROM CATCLIENTES WHERE Id_Cliente = '" & rtrim(session("usuario")) & "'"
				rs.Open SQL
				sesionActiva = rs.fields("Sesion_Activa")				
				rs.close
				if sesionActiva = 0 then response.redirect "Cerrar_Ses_Cli.asp"
            %>
 
            <%

				dim user
				user=session("usuario") 
				pass=session("contrasena") 
				Nombre=session("Nombre")  
            %>
        
            <%
				if(session("statusUsuario")="usuariosecundario") then
			%>
				<script language="Javascript" type="text/javascript">
					location.href="Cerrar_Ses_Cli.asp";
				</script>
			<%
			end if
			%>
			<%
			Application.Lock 
				Application("num_usuario")
			Application.Unlock
			
			'nuevo campo de busqueda 2
			expiracionLogin = DateAdd("n", Session.TimeOut, Now())
			session("banderaTipoUser") = "S"
			
			dim seleccione2
			dim texto2
			dim seleccione
			dim texto
			dim seleccione2DaCon
			dim texto2DaCon
			dim fechaini
			dim fechafin
			dim tipofecha
			dim complFechas
            dim lg
			dim mesActual
			dim hora, minuto, segundo

			Dim tipoUser
			tipoUser = "ADMIN"

			dim nombreUser



			If Hour(now) < 10 Then
				hora = "0"&Hour(now)
			Else
				hora = Hour(now)
			End If

			If Minute(now) < 10 Then
				minuto = "0"&Minute(now)
			Else
				minuto = Minute(now)
			End If

			If Second(now) < 10 Then
				segundo = "0"&Second(now)
			Else
				segundo = Second(now)
			End If

			if Month(now) < 10 Then
				mesActual = "0"&Month(now)
			Else
				mesActual = Month(now)
			End if
 
 
			fechaini=Replace(request.QueryString("datepicker"),"'","")
			fechafin=Replace(request.QueryString("datepickerfinal"),"'","")
			tipofecha=Replace(request.QueryString("tipofecha"),"'","")
			
			complFechas=""
			if(fechaini<>"" or fechafin <>"") then
					if(fechafin <>"") then
						complFechas=" and "&tipofecha&"='"&Replace(fechafin, "-", "") &"' "
					end if
					if(fechaini<>"") then
						complFechas=" and "&Replace(tipofecha,"'","")&"='"&Replace(fechaini, "-", "") &"' "
					end if
					if(fechaini<>"" and fechafin <>"") then
						complFechas=" and "&tipofecha&">='"&Replace(fechaini, "-", "") &"' and "&tipofecha&"<='"&Replace(fechafin, "-", "") &"' "
					end if
			end if

            lg=Request.QueryString("ln")
			 ' response.write lg
				
			'nuevo campo de busqeda
			seleccione2 = request.QueryString("seleccione2")
			'seleccione2 = trim(Replace(seleccione2,"'",""))
			texto2 = request.QueryString("texto2")
			texto2 = trim(Replace(texto2,"'",""))
            texto2= Replace(texto2,"+", " ")
			
			seleccione = request.QueryString("seleccione")	
			'seleccione = trim(Replace(seleccione,"'",""))

			texto = request.QueryString("texto")
			texto = trim(Replace(texto,"'",""))
            texto= Replace(texto,"+", " ")

			if texto="MAIL" or texto="mail" then 
				texto="C001"
			else if texto="FTP" or texto="ftp" then 
				texto="C007"
				else if texto="VAN" or texto="van" then 
					texto="C012"
					else if texto="WEB" or texto="web" then 
						texto="C013"
						else if texto="AS2" or texto="as2" then 
							texto="C014"
							else if texto="SFTP" or texto="sftp" then 
								texto="C015"
								else if texto="Enviado anteriormente" then 
									texto="ERROR15"
									else if texto="Duplicado en transmisión" then 
										texto="ERROR14"
											else if texto="Desconectado" then 
												texto="ERROR13"
												else if texto="No es cliente PAEBSA" then 
													texto="ERROR11"
														else if texto="No es proveedor" then 
															texto="ERROR07"
														end if
												end if
											end if
										end if
									end if
								end if
							end if
						end if
					end if
				end if	
			end if	
							
			fechaactual = date()
			TimeDiff = (DateAdd("m", -2 , fechaactual))
			ayeranio = mid(TimeDiff,1,4)
			ayermes = mid(TimeDiff,6,2)
			ayerdia = mid(TimeDiff,9,2)
			ayerFecha=ayeranio&ayermes&ayerdia

			'Inicia el segundo campo de busqueda
			dim valor2
			valor2=""
			
			if seleccione2<>"" and texto2 <> "" then
				texto2 = Replace(texto2, "/", "")
					if seleccione2 ="archivo" then
						valor2= " and  (Nombre_Archivo like '%"&texto2&"%' or Nombre_Archivo_PDF like '%"&texto2&"%' OR   Nombre_Archivo_CSV like '%"&texto2&"%' or Nombre_Archivo_Txt like '%"&texto2&"%' or Nombre_Archivo_Excel like'%"&texto2&"%')"
					else
						valor2= " and  "&seleccione2&" like '%"&texto2&"%'"
					end if
			else
				valor2 =""
			end if	
			'termina el segundo campo de busqueda
		
			if seleccione<>"" and texto <> "" then
				texto = Replace(texto, "/", "")
				if seleccione ="archivo" then
					buscar= "and  (Nombre_Archivo like '%"&texto&"%' or Nombre_Archivo_PDF like '%"&texto&"%' OR   Nombre_Archivo_CSV like '%"&texto&"%' or Nombre_Archivo_Txt like '%"&texto&"%' or Nombre_Archivo_Excel like'%"&texto&"%')"
				else
					buscar= "and  "&seleccione&" like '%"&texto&"%'" 
				end if
			else
				buscar =""
			end if	
			
			
			tamanopagina=request.QueryString("tamanopagina")
			
			if tamanopagina = "" then
				tamanopagina=25
			end if
					
			paginaabsoluta=request.QueryString("paginaabsoluta")
			if paginaabsoluta="" then
				paginaabsoluta=1
			end if
			
			orden = request.QueryString("orden")
			orden = trim(orden)
			if orden = "" then
				orden="Fecha_Envio_Proveedor"
			end if
			
			alf = request.QueryString("alf")
			alf = trim(alf)
			if alf = "" then
				alf="desc"
			end if
			
            buscar= buscar & valor2 &complFechas
			if buscar<>"" then	
				condicion=2
			else 
				condicion=3
			end if
			user = rtrim(user)
			'Response.write "BUSQUEDA "&buscar
			
			On Error Resume Next
			Set dataUser = GetDataCliente(user)
			Session.Timeout= 30 ' dataUser.Item("session")
			
			' Esta variable indica el periodo inicial 
			' Para consulta de informacion.
			' Los meses se convierten en dias.
			Info_Dias= dataUser.Item("meses")
			
			sql = "Select Id_cliente,Fecha_Envio_Proveedor,Codigo_Cliente,Nombre_Hub,Id_Hub,Codigo_Transaccion,Num_control_dato_docto,Num_Intercambio_Recibido,Hora_Envio_Proveedor,"&_
				  "Fecha_Documento_Edi,Fecha_Canc_Documento_Edi,Consecutivo_Int_Pebsa,Numero_Proveedor_Hub,Status,Nombre_Archivo,Fecha_Consulta_Cliente,Hora_Consulta_Cliente,Identificador_Formato_1,"&_
				  "Nombre_Archivo_PDF,Nombre_Archivo_CSV,Nombre_Archivo_Txt,Nombre_Archivo_Excel,Nombre_Archivo_XML,Fecha_Recepcion_Sistema,Codigo_Tienda,Descripcion_Error,Nombre_Archivo_Etiquetas,Nombre_Archivo_Log  from Vista_DIARIAENVIOSHIST "&_
				  "where Id_Cliente='" & user & "' and Identificador_Canal_1='C013' and  Identificador_Formato_1='F001' " &_
				  "and Fecha_Recepcion_Sistema >=Convert(char(35),DateAdd(d, -"&Info_Dias&" ,getdate()) , 112) AND Fecha_Recepcion_Sistema<=Convert(char(35),getdate() , 112)  "&buscar&" "&_
				  "order by "& orden & " "&alf&",Hora_Envio_Proveedor desc "
			'response.write sql
			
			'cnn.Open 
		    rs.Open sql,cnn,3,1
            cant_paginas=rs.PageCount
		   'ttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttt
		   %>
		   
		   
		   <%
		   
		If (Err.Number <> 0) Then
		
			' Inicia Genera Archivo LOG errores del sistema 
			dim archivo
			
			archivo= request.serverVariables("APPL_PHYSICAL_PATH") & "errores\sistema.txt" 
			
			set confile = createObject("scripting.filesystemobject") 
			set fich = confile.OpenTextFile (archivo,8) 
			'escribo en el archivo 
			fich.WriteLine("Información general del proceso.") 
			fich.WriteLine(""&now()&" - Id Cliente: "&user& " | Codigo Cliente: "&pass&"") 
			fich.WriteLine("Error: "&Err.Number&": "& Err.Description) 
			fich.WriteLine("-----------------------------") 
			fich.WriteLine("") 
			'cerramos el fichero 
			fich.close() 
			response.redirect"Men_Mantenimiento.asp"
			' Termina Genera Archivo LOG errores del sistema  
			
			
%>


<%
		


else
			' Inicia pagina que se muestra para el usuario cuando el proceso es correcto
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Expires" content="0" />
	<meta http-equiv="Pragma" content="no-cache" />
	
    <link href="imagenes2/paebsa.ico" rel="shortcut icon" type="image/vnd.microsoft.icon" />
	
	<script src="jsFromHttp/jquery-1.9.1.js" type="text/javascript"></script>
	<script src="jsFromHttp/jquery-ui.js" type="text/javascript"></script>
	
    <link href="css/loginPaebsa.css" rel="stylesheet" type="text/css" />
    <link href="css/disenioTabla.css" rel="stylesheet" type="text/css" />	
	
    <!--<script type="text/javascript" src="jquery/jquery_validate.js"></script>-->
    <script type="text/javascript" src="jquery/jquery.jMagnify.js" ></script>
    <script type="text/javascript" src="jquery/jqueryTooltip.js" ></script>
	<script type="text/javascript" src="jquery/jquery.cycle.all.2.74.js"></script>
	<script type="text/javascript" src="js/Functions.js"></script>
	
	<link type="text/css" rel="stylesheet" href="jsFromHttp/jquery-ui.css" />
	<link href="css/960.css" rel="stylesheet" media="screen" />
	<link href="css/defaultTheme.css" rel="stylesheet" media="screen" />
	<link href="css/myTheme.css" rel="stylesheet" media="screen" />
	<script src="js/jquery.fixedheadertable.js" type="text/javascript"></script>
    <script src="demo.js" type="text/javascript"></script>
	
	<!-- Traductor de la pagina Espaniol Ingles -->
	<script src="js/translate.js" type="text/javascript"></script>
    <script src="../js/i18next/paebsa/bluebird.min.js" type="text/javascript"></script>
    <script src="../js/i18next/paebsa/i18next.js" type="text/javascript"></script>
    <script src="../js/i18next/paebsa/jquery-i18next.js" type="text/javascript"></script>
    <script src="../js/i18next/paebsa/i18nextXHRBackend.min.js" type="text/javascript"></script>
    <script src="../js/i18next/paebsa/traslatePaebsa.js" type="text/javascript"></script>
		<!-- Traductor de la pagina Espaniol Ingles -->

<!-- Agregar boostrap -->

    <!--<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
    ' <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">
    ' <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>
 -->
	
		
	<title>PAEBSA</title>
	<script type="text/javascript">
		if (window.history) {
				function noBack(){window.history.forward()}
				noBack();
				window.onload=noBack;
				window.onpageshow=function(evt){if(evt.persisted)noBack()}
				window.onunload=function(){void(0)}
		}
	</script>
	
	<script type="text/javascript">
        
        $(document).ready(function () 
		{
		 
			traslate('<%=lg%>', 'loginPaebsa');
            //
			$('.slideshow').cycle({
				fx:      'turnDown', 
				delay:   -4000 
			});
			
			var idCliente=$('#buzon').text().trim();
			if(idCliente=='NESTLE')
				$('#link_desadv').show();
			else
				$('#link_desadv').hide();
				
				
			
			var idCliente=$('#buzon').text().trim();
			if(idCliente=='NESTLE')				
				$('#link_cargaInfo').show();
			else				
				$('#link_cargaInfo').hide();
				
			var arrayName=[];
            var arrayIdProveedor = [];
            var arrayTransaccion = [];
			<%
                    
				Call GetNumberProveedorSpoke(user,"")
                Call GetDataSpoke(user, "Proveedor")
                GetTransaction(user)
			%>
			var object;
            object = { arrayN: arrayName, arrayIP: arrayIdProveedor, arrayT: arrayTransaccion };
            $("#seleccione").change(function (handler) {
                var index = handler.target.selectedIndex;
                disabledSelect($("#seleccione2"), index);
                $("#texto").val("");
                autocompleteTextBox(index, object, $('#texto'), '<%=trim(user)%>', '<%=Info_Dias%>', 'S', '');
                //$("#seleccione option:selected").each(function (e) { console.log(e); });
			});
				
			$("#seleccione2").change(function(handler){
				var index= handler.target.selectedIndex;
                disabledSelect($("#seleccione"), index);
                $("#texto2").val("");
				//cleanFilters($("#formInscripcion"),$("#seleccione2"));
				autocompleteTextBox(index,object,$('#texto2'),'<%=trim(user)%>','<%=Info_Dias%>','S','');
			});

            $("#tipofecha").change(function () {
                $('#datepicker').val("");
                $('#datepickerfinal').val("");
            });

            var objectData = { idCliente: '<%=trim(user)%>', meses: '<%=Info_Dias%>', type: 'S', idUserSec: '', array: object };
            filters($('#formInscripcion'), '<%=texto%>', '<%=texto2%>', '<%=seleccione%>', '<%=seleccione2%>', '<%=orden%>', '<%=alf%>', '<%=tipofecha%>', '<%=fechaini%>', '<%=fechafin%>', '<%=tamanopagina%>',objectData);
            
            $('#second').jMagnify({
				centralEffect: {'color': 'yellow'},
				lat1Effect: {'color': 'orange'},
				lat2Effect: {'color': 'red'},
				lat3Effect: {'color': 'magenta'},
				resetEffect: {'color': '#1E598E'}
            });

          
		});
        

		function openTemplate(cliente,idUsuario){
			var parametros={idCliente: cliente, idUser: idUsuario,language:'<%=lg%>'};
			var propiedades= JSON.stringify({width:'80%', height:'600'});
			browser('AplicacionPaebsa/ASNExcel.aspx?',parametros, propiedades);
		}
		
		function openNoSpots(idCliente, nombre, usuario){
			var parameters={ idClient: idCliente, name: nombre, user: usuario, iduser: 'ADMIN',language:'<%=lg%>' };
			var propiedades= JSON.stringify({width:'80%', height:'600'});
			browser('AplicacionPaebsa/reporteVentasNoSpots.aspx?', parameters, propiedades);
			}
			
		function openBrowser(idCliente, nombre, usuario){
			var parameters={ idClient: idCliente, name: nombre, user: usuario, iduser: 'ADMIN',language:'<%=lg%>' };
			var propiedades= JSON.stringify({width:'80%', height:'600'});
			browser('AplicacionPaebsa/Browser.aspx?', parameters, propiedades);
			

		}
	</script>
	
	
	<script type="text/javascript">
	/* Cuadro de dialogo que se muestra al usuario para descarga de archivos del Portal*/
	/* Tabla de Informacion/ Columna 'Descargar'/ Cuadro de dialogo */
	$(document).ready(function () 
	{
		$("#dialog-form").dialog({
			autoOpen: false,
			height: 150,
			width: 350,
			modal: true,
			resizable: false,
			close: function() {}
		});

		$('.create-user').on('click',function(eEvento){
			$("#dialog-form").dialog("option","position",[event.offsetX="center", event.offsetY="center"]); // esta linea centra el dialog cada vez que abre
			$( "#dialog-form" ).dialog( "open" );
		});
	});
	</script>
		
		
	
	<script type="text/javascript">
		function marcar() {
			obj=arguments[0];
			for(i=1;i<arguments.length;i++)
			{
				marca=arguments[i].replace('fila','');
				marca='c'+marca;
				if (obj.checked)
				{
					document.getElementById(arguments[i]).style.backgroundColor='#E8FF9F';
					document.getElementById(marca).checked=true;
				}
				else
				{
					document.getElementById(arguments[i]).style.backgroundColor='';
					document.getElementById(marca).checked=false;
					document.getElementById('cTodos').checked=false;
				}
			}
		}
	</script>

	<script type="text/javascript">
				function marcarb2() {
				   try{
					obj=arguments[0];
					var urlArchivos="";
					var contadorarcl=0;
						for(i=1;i<arguments.length;i++)
						{
							marca=arguments[i].replace('fila','');
							marca='c'+marca;
								if (obj.checked){
								}
								else{
								//EL +10 INDICA EL NUMERO DE LINKS ANTES DE LOS CHECKBOX, SI SE AUMENTE UN BOTN O LINK AY QUE AUMENTAR LA CIFRA DEL +10
									selChk = document.getElementsByTagName("a")[i+26];							
									if(document.getElementById(marca).checked){
									   if(contadorarcl<20){
										var rCliente="";
										var rArchivo="";
										var url = selChk.href;
 
 										var indexCliente = url.indexOf("?");
										var indexArchivo = url.indexOf("?");

 										indexCliente = url.indexOf("cliente",indexCliente) + "cliente".length;
 										indexArchivo = url.indexOf("archivo",indexArchivo) + "archivo".length;

										if (url.charAt(indexCliente) == "="){
											var rCliente = url.indexOf("&",indexCliente);
											if (rCliente == -1){rCliente=url.length;};
											urlArchivos=urlArchivos+"nCl"+contadorarcl+"="+url.substring(indexCliente + 1,rCliente)+"&";
 										}
										if (url.charAt(indexArchivo) == "="){
											var rArchivo = url.indexOf("&",indexArchivo);
											if (rArchivo == -1){rArchivo=url.length;};
											urlArchivos=urlArchivos+"nAr"+(contadorarcl)+"="+url.substring(indexArchivo + 1,rArchivo)+"&";
 										}
										contadorarcl++;
									   }
									}
							}
						}
					  }catch(e){
					  }finally{
					  
					  		var adjuntarFicheros = confirm ("AVISO: -Para el envio de archivos por E-mail-\nPara enviar los archivos adjuntos como uno solo presione Aceptar,\npara enviarlos por separado presione cancelar.")
							if (adjuntarFicheros){
							location.href="EnviaEmailUsuarioMaestro.asp?"+urlArchivos+"totvar="+contadorarcl+"&adf=S";
							}
							else{
							location.href="EnviaEmailUsuarioMaestro.asp?"+urlArchivos+"totvar="+contadorarcl+"&adf=N";
							}
						
						}

				}
	</script>
 
	

    <script type="text/javascript">
        $(function () {
            var lenguaje ="<%=lg%>";
            
            if (lenguaje== "es"){
            //documentacion: http://www.ajaxshake.com/demo/ES/288/98f18e75/selector-de-fechas-javascript-calendario-gratuito-para-tus-paginas-web-jquery-ui-datepicker.html
                $("#datepicker").datepicker({
                    dateFormat: "yy-mm-dd",
                    monthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"],
                    dayNamesMin: ["Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sa"]
                });
                //documentacion: http://www.ajaxshake.com/demo/ES/288/98f18e75/selector-de-fechas-javascript-calendario-gratuito-para-tus-paginas-web-jquery-ui-datepicker.html
                $("#datepickerfinal").datepicker({
                    dateFormat: "yy-mm-dd",
                    monthNames: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"],
                    dayNamesMin: ["Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sa"]
                });
            }
            if (lenguaje == "en") {
                $("#datepicker").datepicker({ dateFormat: "yy-mm-dd" });
                $("#datepickerfinal").datepicker({ dateFormat: "yy-mm-dd" });
            }
		});
	</script>
	
	<script type="text/javascript">
		function obtenerParametros(parameter){
			// Obtiene la cadena completa de URL
			 var url = location.href;
			 
			 /* Obtiene la posicion donde se encuentra el signo ?, 
			 ahi es donde empiezan los parametros */
			 var index = url.indexOf("?");
			 /* Obtiene la posicion donde termina el nombre del parametro
			 e inicia el signo = */
			 index = url.indexOf(parameter,index) + parameter.length;
			 /* Verifica que efectivamente el valor en la posicion actual 
			 es el signo = */ 
			 if (url.charAt(index) == "="){
				 // Obtiene el valor del parametro
				 var result = url.indexOf("&",index);
				 if (result == -1){result=url.length;};
				 // Despliega el valor del parametro
				 return url.substring(index + 1,result);
			 }
		}
	</script>
	
	<script type="text/javascript">
		function descargaExcel(){
			var fechaini=''+obtenerParametros('datepicker');
			var fechafin=''+obtenerParametros('datepickerfinal');
			var tipofecha=''+obtenerParametros('tipofecha');
			var seleccione2=''+obtenerParametros('seleccione2');
			var texto2=''+obtenerParametros('texto2');
			var seleccione=''+obtenerParametros('seleccione');
			var texto=''+obtenerParametros('texto');
			var orden=''+obtenerParametros('orden');
			var alf=''+obtenerParametros('alf');

			if(fechaini=='undefined'){fechaini='';}
			if(fechafin=='undefined'){fechafin='';}
			if(tipofecha=='undefined'){tipofecha='';}
			if(seleccione2=='undefined'){seleccione2='';}
			if(texto2=='undefined'){texto2='';}
			if(seleccione=='undefined'){seleccione='';}
			if(texto=='undefined'){texto='';}
			if(orden=='undefined'){orden='';}
			if(alf=='undefined'){alf='';}

			
			location.href='ExportaTablaAExcel.asp?'+'datepicker='+fechaini+'&datepickerfinal='+fechafin+'&tipofecha='+tipofecha+'&seleccione2='+seleccione2+'&texto2='+texto2+'&seleccione='+seleccione+'&texto='+texto+'&orden='+orden+'&alf='+alf+"&ln=<%=lg%>";
		}
	</script>

	<script type="text/javascript">
		function reprocesoarchivos(){
		try{
				obj=arguments[0];
				var urlArchivos="";
				var contadorarcl=0;
					for(i=1;i<arguments.length;i++)
					{
						marca=arguments[i].replace('fila','');
						marca='c'+marca;
							if (obj.checked){
							alert(obj.value);
							}
							else{
								if(document.getElementById(marca).checked){
									if(contadorarcl<20){
										urlArchivos=urlArchivos+""+(document.getElementById(marca).value)+"&";
										contadorarcl++;
									}
								}
							}
					}
					}catch(e){
					}finally{
						location.href="ReprocesoArchivos.asp?ln=<%=lg%>&idc="+"<%=trim(user)%>"+"&coc="+"<%=trim(pass)%>"+"&"+urlArchivos+"totvar="+contadorarcl;
					}
		}
	</script>

	<script type="text/javascript">
		function generarPDFs(){
			try
			{
				obj=arguments[0];
				var urlArchivos="";
				var contadorarcl=0;
				for(i=1;i<arguments.length;i++)
				{
					marca=arguments[i].replace('fila','');
					marca='c'+marca;

						if (obj.checked){
							alert(obj.value);
						}
						else{
							if(document.getElementById(marca).checked){
								if(contadorarcl<20){
									urlArchivos=urlArchivos+""+(document.getElementById(marca).value)+"&";
									contadorarcl++;
									
								}
							}
						}
				}
			}
			catch(e)
			{
			}
			finally
			{
				location.href="GenerarPDF.asp?ln=<%=lg%>&idc="+"<%=trim(user)%>"+"&coc="+"<%=trim(pass)%>"+"&"+urlArchivos+"totvar="+contadorarcl;
			}
		}
	</script>

	<link rel="stylesheet" href="css/EstiloEnviaFacturas.css" type="text/css" />

	
	<script type="text/javascript">
		$(document).ready(function() {
			$(".mainCompose").hide();
			$('.loader').hide();
			$('#errortxt').hide();
			$('.compose').click(function() {
				$('.mainCompose').slideToggle();
			});
			$('.sendbtn').click(function(e) {
				e.preventDefault();
				$('.sendbtn').hide();
				$('.loader').show();
				if($('#mymsg').val() == "") {
					$('#errortxt').show();
					$('.sendbtn').show();
					$('.loader').hide();
				}
				else {
					$('sendbtn').hide();
					$('.loader').show();
					$('#errortxt').hide();
					var formQueryString = $('#sendprivatemsg').serialize(); // form data for ajax input
					finalSend();    		
				}
				// possibly include Ajax calls here to external PHP
				function finalSend() {
					$('.mainCompose').delay(1000).slideToggle('slow', function() {
						$('#composeicon').addClass('sent').removeClass('compose').hide();
					
						// hide original link and display confirmation icon
						$('#composebtn').append('<img src="img/check-sent.png" />');
					});
				}
			});
		});
	</script>	

	<script type="text/javascript">
		$(document).ready(function() {
			$(".mainEDI").hide();
			$('.loader').hide();
			$('#errortxt').hide();
			$('.composeEdi').click(function() {
				$('.mainEDI').slideToggle();
			});
			$('.sendbtn').click(function(e) {
				e.preventDefault();
				$('.sendbtn').hide();
				$('.loader').show();
				
				if($('#mymsg').val() == "") {
					$('#errortxt').show();
					$('.sendbtn').show();
					$('.loader').hide();
				}
				else {
					$('sendbtn').hide();
					$('.loader').show();
					$('#errortxt').hide();
					
					var formQueryString = $('#sendprivatemsgEdi').serialize(); // form data for ajax input
					finalSend();    		
				}
				function finalSend() {
					$('.mainEDI').delay(1000).slideToggle('slow', function() {
						$('#composeiconEdi').addClass('sent').removeClass('composeEdi').hide();
					
						$('#btnEdi').append('<img src="img/check-sent.png" />');
					});
				}
			});
		});
	</script>	

	<script type="text/javascript" language="javascript">  
		 var iStart = 0;
		 var iMinute = <%=Session.Timeout%>; //Obtengo el tiempo de session permitida
		 function showTimer() 
		 {
			lessMinutes(); 
		 } 
		 
		 function lessMinutes()
		 {
			 //Busco mi elemento que uso para mostrar los minutos que le quedan (minutos y segundos)
			 obj = document.getElementById('TimeLeft'); 
			 if (iStart == 0) 
			 {
				 iStart = 60 
				 iMinute -= 1; 
			 }
			 
			 iStart = iStart - 1;

			 var modulo=iStart%2;
			 if(iMinute<=2 && modulo==0)
			 {
					document.getElementById("msgSesion").style.color="#FF0000";
			 }
			 else
			 {
					document.getElementById("msgSesion").style.color="#B40431";
			 }
			 //Si minuto y segundo = 0 ya expiró la sesion 
			if (iMinute==0 && iStart==0) 
			{
				obj.innerText = " - Su sesion ha expirado -";
                var mensaje = $.t("alertas.terminaSession");
                alert(mensaje);//"Su sesion ha expirado, usted sera redireccionado a la pagina principal.\nPAEBSA"
				 $.ajax
				(
					{
						type: "POST",
						url: "AplicacionPaebsa/Procesos.asmx/cerrarSesionMaestro",
						data: "{idCliente: '" + '<%=rtrim(session("usuario"))%>' + "', tipoUsr: 'S'}",
						contentType: "application/json; charset=utf-8",
						async: true,
						dataType: "json",
						success: function (data, status) 
						{
							var respuesta = data.d;
							console.log(respuesta);
							location.href = "Cerrar_Ses_Cli.asp"
							
						},
						failure: function (xhr, status, error) 
						{
							console.log("Error");
							console.log(xhr);
						}
					}
				);
			}
			
			if (iStart < 10)
				obj.innerText = iMinute.toString() + ':0' + iStart.toString();
			 else
				obj.innerText = iMinute.toString() + ':' + iStart.toString();
			//actualizo mi método cada segundo  
			 window.setTimeout("lessMinutes();",999)
		 }
	</script>

 
 
	<script type='text/javascript'>
        function validarMaximoArchivos() {
                var text = "";
				var $fileUpload = $("input[type='file']");
            if (parseInt($fileUpload.get(0).files.length) == 0) {
                    text = $.t("alertas.envioFacturas.sinArchivos");
				    alert(text);//"No hay archivos XML seleccionados para subir"
				 return false;
				}
				
            if (parseInt($fileUpload.get(0).files.length) > 50) {
                text = $.t("alertas.envioFacturas.limite");
				 alert(text);//"Solo se permite subir 50 archivos XML por cada carga."
				 return false;
				}
				var files = $("input[type='file']").get(0).files;
				for (i = 0; i < files.length; i++)
				{
                    if (files[i].size > 3145728) {
                        text = $.t("alertas.envioFacturas.peso");
				    alert(text +files[i].name+" * "+files[i].size/1024/1024);
				   return false;
				   }
				   //pdfs
				   var extensionPdf=files[i].name.split('.').pop();
				   extensionPdf=extensionPdf.toUpperCase();
				   if(extensionPdf=="PDF")
				   {
						var xmlYpdf=false;
						var nombreActualPDF=files[i].name.toUpperCase();
						var totalArchivos = $("input[type='file']").get(0).files;
						for (j = 0; j < totalArchivos.length; j++)
						{
							var nombreActual=totalArchivos[j].name.toUpperCase();
							if(nombreActualPDF.replace(".PDF",".XML")==nombreActual){
								xmlYpdf=true;
							}
						}
				   
                       if (!xmlYpdf) {
                           text = $.t("alertas.envioFacturas.pdf");
							alert(text);//"Si desea enviar archivos .pdf también deberá proporcionar su archivo .xml con el mismo nombre que el pdf"
							return false;
						}   
				   }
				}
				return true;
		}
	</script>
	
	<script type='text/javascript'>
        function validarMaximoArchivosAddenda() {
			var text = "";
			var fileInput = document.getElementById("fuFiles");
			var files = fileInput.files;
			var maximo = 50;
	
			if (files.length > maximo) {
				document.getElementById("fuFiles").innerHTML = "";
				 text = $.t("alertas.envioFacturas.limite");
				 alert(text);//"Solo se permite subir 50 archivos XML por cada carga."
				return false;
			}

			return true;
		}
	</script>
	
	<!-- Validacion de archivos EDI-->
	<script type='text/javascript'>
		function validarMaximoArchivosEdis(){
            var inp = document.getElementById('archivoEdi');
            var mensaje = "";
            if (parseInt(inp.files.length) == 0) {
                    mensaje = $.t("alertas.envioASN.sinArchivos");
                    alert(mensaje);//"No hay archivos EDI seleccionados para subir"
				    return false;
				}
                if (parseInt(inp.files.length) > 10) {
                    mensaje = $.t("alertas.envioASN.limite");
                    alert(mensaje);//"Solo se permite subir 10 archivos EDI por cada carga."
				 return false;
				}
				var files = inp.files;
				for (i = 0; i < inp.files.length; i++)
				{
                    if (inp.files.item(i).size > 3145728) {
                        mensaje = $.t("alertas.envioASN.peso");
                        alert(mensaje + inp.files.item(i).name + " * " + (inp.files.item(i).size / 1024 / 1024) + " MB");
				   return false;
				   }
				   //pdfs
				   var extensionPdf=files[i].name.split('.').pop();
				   extensionPdf=extensionPdf.toUpperCase();
				   
				   if(extensionPdf!="EDI")
                   {
                       mensaje = $.t("alertas.envioASN.formato");
                       alert(mensaje);//"Los archivos a enviar deben estar en formato EDI"
						return false;
				   }
				}
				return true;
		}
	</script>
	
	<script type="text/jscript">
		function fechaactualPantalla() 
		{         
			// GET CURRENT DATE
			var date = new Date();
			 
			// GET YYYY, MM AND DD FROM THE DATE OBJECT
			var yyyy = date.getFullYear().toString();
			var mm = (date.getMonth()+1).toString();
			var dd  = date.getDate().toString();
			 
			// CONVERT mm AND dd INTO chars
			var mmChars = mm.split('');
			var ddChars = dd.split('');
			 
			// CONCAT THE STRINGS IN YYYY-MM-DD FORMAT
            var datestring = yyyy + '-' + (mmChars[1] ? mm : "0" + mmChars[0]) + '-' + (ddChars[1] ? dd : "0" + ddChars[0]);
           // var stringDate = $.t("sistema.fecha");
			$('#today').html(datestring);
		}
	</script>
	
	<script type="text/javascript">
	   function submitForm()
	   {
		   document.reportesLogs.target = "myActionWin";
		   window.open("","myActionWin","height=650px,width=750px,toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,modal=yes");
		   document.reportesLogs.submit();
	   }
	</script>
	
   <script type="text/javascript">
	   function submitReportesExcel()
	   {
		   document.reportesExcel.target = "myActionWin";
		   window.open("","myActionWin","height=650px,width=750px,toolbar=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,modal=yes");
		   document.reportesExcel.submit();
	   }
   </script>

   <script language="javascript" type="text/javascript">
		function ventanaHistorial() 
		{
		//window.showModalDialog('HistorialProveedores.asp', '', 'status:1; resizable:1; dialogWidth:900px; dialogHeight:750px; dialogTop=50px; dialogLeft:100px')
			window.open("HistorialProveedores.asp", "_blank", "toolbar=no, scrollbars=yes, resizable=yes, top=50, left=50, width=1200, height=700");
		}
	</script>
	
	<script>
		$(document).ready(function()
		{					
			var hoy = new Date();
			localStorage['idCliente'] = '<%=trim(user)%>';
			var bandera = '<%=session("statusUsuario")%>';
			if(bandera=="usuariomaestro")
				localStorage['tipoUser'] = 'Spokes';
			else if(bandera=="Administrador")
				localStorage['tipoUser'] = 'Administrador';	
			localStorage['fechaActual'] = '<%=Year(now)&"/"&mesActual&"/"&Day(now)%>';	
			localStorage['horaActual'] = '<%=hora&":"&minuto&":"&segundo%>'+ '.' + hoy.getMilliseconds() +'';
			localStorage['nombreUser'] = $('#nombreUser').text().trim();
			
			//console.log(localStorage['nombreUser']);
		});	
	</script>

	<script>
		localStorage['tipoCliente'] = "S";
		localStorage['banderaTipo'] = "E";
	</script>

</head>
<body> 
	<div class="texto recibo">
		<div class="encabezado">   			
			<div class="serviciosTitulo2"><h2 data-i18n="sistema.titulo"><strong>Servicio de Bur&oacute Electr&oacutenico "Proveedores"</strong></h2></div>
     
			<div class="contenido-encabezado">
				<div class="imagen"><img src="imagenes2/servicioConsulta.png" alt="PAEBSA"/></div>				
				<div class="serviciosTexto" style="color:#FAFAFA;"><strong data-i18n="sistema.user">Bienvenido: </strong><label id="nombreUser"><%=Nombre%></label>
					<br/><strong data-i18n="sistema.buzon">Usuario/Buz&oacute;n: </strong><label id="buzon"><%=user%></label>
					<br /><strong data-i18n="sistema.encabezado">Informaci&oacuten recibida </strong>
					<br /><strong><label data-i18n="sistema.fecha"></label><label id="today"/><script type="text/javascript">fechaactualPantalla();</script> </strong>
				</div>
			</div>
		</div>	
	</div>
	<!--<div class="block" id="block"></div>-->



<!--<nav class="navbar naxvbar-expand-sm bg-dark navbar-dark">
  <div class="container-fluid">
    <a class="navbar-brand" href="#">Nuevas funciones del portal</a>
    <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#collapsibleNavbar">
      <span class="navbar-toggler-icon"></span>
    </button>
    <div class="collapse navbar-collapse" id="collapsibleNavbar">
      <ul class="navbar-nav">
        <li class="nav-item">
          <a class="nav-link" href="#">Link</a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="#">Link</a>
        </li>
        <li class="nav-item">
          <a class="nav-link" href="#">Link</a>
        </li>
      </ul>
    </div>
  </div>
</nav>-->





<nav class="navbar navbar-expand-md navbar-light">
  <div class="container-fluid">
    <div class="collapse navbar-collapse" id="main-menu">
      <ul id="menu-menuprincipalbootstrap-5" class="navbar-nav me-auto mb-2 mb-md-0 dropdown-menu-start">
        <li id="menu-item-38" class="menu-item menu-item-type-post_type menu-item-object-page 
                menu-item-has-children dropdown nav-item nav-item-38">
                <a href="#" class="nav-link  dropdown-toggle" data-bs-toggle="dropdown" 
                aria-haspopup="true" aria-expanded="false">Nuevas funciones del portal</a>
            <ul class="dropdown-menu  depth_0">
            
			

				<li id="composebtn">
							<a href="#"  class="compose" id="composeicon" data-i18n="menu.factura.titulo"> &raquo;Env&iacuteo de facturas a clientes</a>
								<div class="mainCompose">
									<div class="calloutUp">
										<div class="calloutUp2"></div>
									</div>	
									<div id="msgform" class="msgEnvio" width="700px">
										<form id="sendprivatemsg" class="UsuariosCss" action="EnvioXML/ValidaXML.aspx" method="post" enctype="multipart/form-data">
											<label data-i18n="menu.factura.xml">Factura XML/EDI</label>
											<input type="file" name="archivo[]" accept="text/xml,.edi" size="70" multiple value="Examine"/>
											<br /><br />
											<label style="color:#B40404;" data-i18n="menu.factura.archivos">N&uacutemero m&aacuteximo de archivos por carga: 50</label>
											<br /><br/>
											<label style="color:#0B4C5F;" data-i18n="menu.factura.aviso">IMPORTANTE: Para enviar facturas con addenda resguardo de Walmart vaya a "Captura de Addendas-> Addendas de Wal-Mart-> Addenda Resguardo"</label>
											<br /><br />
											<%
												sqlProveedorMerza = "select rtrim(id_cliente)id_Cliente, Codigo_Cliente,Codigo_Transaccion_Produccion,RFCSpoke,RFCHub from CATSPOKESHUBS where Codigo_Cliente='"&trim(pass)&"' and Id_Cliente='"&trim(user)&"' and Codigo_Transaccion_Produccion='INVOIC' and RFCHub='ADU800131T10'"
												set rsProveedor=server.createobject("ADODB.Recordset") 						
												rsProveedor.Open sqlProveedorMerza,cnn,3,1	
												if rsProveedor.EOF then
												else
												response.write "<label  style='color:#B40404;' data-i18n='menu.factura.avisoMerza'> SI ERES PROVEEDOR DE MERZA, FAVOR DE SUBIR FACTURAS CON LA ADDENDA SOLICITADA</label></a><br/><br/>"
												end if

											%>
											<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
											<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
											<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
											<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>
											<div style="padding-bottom: 25px;">
												<div class="input" style="float:right;">
													<input class="button_opt prtText" onclick="return validarMaximoArchivos()" type="submit" id="Submit1" value="Enviar facturas" data-i18n="[value]menu.factura.boton" />
												</div>
											</div>
											<br /><br />
										</form>
									  </div>
								</div>
								<!-- termina cuadro de dialogo de facturas -->
						</li>
						
						<!-- Modulo_Genera_Addenda_Nube -->
						<li  class="has-sub dropdown-menu"><a href="#" data-i18n="[html]menu.generarAddendaAutomatica.titulo"> &raquo; Generar Addenda en la nube</a>
							<ul>
								<li>
								<% Call CargaDeAddendaGenerica(pass,user,Nombre,"loginPaebsa.asp?ln="&lg) %><br />
								</li>			
							</ul>
						</li>
						<!-- Modulo_Genera_Addenda_Nube -->
						
						
						<!-- Modulo_Envio_EDI_Clientes -->
						<li id="btnEdi">
						<!-- Inicia cuadro de dialogo de archivos EDI -->
							<a href="#" class="composeEdi" id="composeiconEdi" data-i18n="menu.asn.titulo"> &raquo;Env&iacuteo de archivos DESADV</a>
								<div class="mainEDI">
									<div class="calloutUp">
										<div class="calloutUp2"></div>
									</div>	
									<div id="msgformEDI" class="msgEnvio" width="700px">
										<form id="sendprivatemsgEdi" class="UsuariosCss" action="AplicacionPaebsa/ValidaXML.ashx" method="post" enctype="multipart/form-data">
											<label data-i18n="menu.asn.envio">Archivos ASN(.edi) </label>
											<input type="file" name="archivoEdi[]" id="archivoEdi" accept="text/edi" size="70" multiple />
											<br /><br />
											<label style="color:#B40404;" data-i18n="menu.asn.aviso">N&uacutemero m&aacuteximo de archivos por carga: 10</label>
											<br />
											<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
											<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
											<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
											<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>									  
											<div style="padding-bottom: 25px;">
												<div class="input" style="float:right;">
													<input class="button_opt prtText" onclick="return validarMaximoArchivosEdis()" type="submit" id="btnenviafac" value="Enviar archivos" data-i18n="[value]menu.asn.boton"/>
												</div>
											</div>
											<br /><br />
										</form>
									</div>
								</div>
							<!-- termina cuadro de dialogo de archivos EDI -->	
						</li>
						<!-- Modulo_Envio_EDI_Clientes -->
						
						<li id="link_cargaInfo"><a href="#" onClick="openBrowser('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="menu.cargaInformacion">&raquo;Carga de informaci&oacute;n</a></li>
						<li><a href="#" onClick="openNoSpots('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="menu.cargaNoSpots">&raquo;Carga de NO SPOTS</a></li>
						
						<li class="has-sub">
						<a href="#" data-i18n="menu.administrarCuenta.titulo">&raquo;Administre su cuenta</a>
							<ul class=" dropdown-menu"> 
								<li><a href="RegistroUsuarios.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.usuarios">&raquo;Administre sus usuarios</a></li>
								<!--Sistema viejo de contraseñas-->
								<li><a href="CambioPassword.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.contrasena"> &raquo;Cambiar contraseña</a></li>
								<!--Sistema nuevo de contraseñas>>>Quitar comentario y comentar línea de arriba-->
								<!--<li><a href="AplicacionPaebsa/ReestablecerContrasena.aspx?tipoUsr=M&pagina=loginPaebsa.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.contrasena"> &raquo;Cambiar contraseña</a></li>-->
								<li><a href="#" id="modal"  onClick="ventanaHistorial();" data-i18n="menu.administrarCuenta.historial">&raquo;Historial de usuarios</a></li>
							</ul>
						</li>
						
						
						
						<li><a href="loginPaebsa.asp?ln=<%=lg%>" data-i18n="menu.general">&raquo;Consulta general</a></li>
	 
						<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
						<li id="link_desadv">
							<a href="#" onClick="openTemplate('<%=trim(user)%>','ADMIN')" data-i18n="menu.template">&raquo;Captura de confirmaci&oacute;n para los templates de Walmart/Sahuayo (DESAV)</a>
						</li>
						<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
	 
						<!-- Menu Colgate -->
						<li class="has-sub" style="display:<% if trim(user)="CPA7503043P1" then response.write "block" else response.write "none" end if %>"><a title="" href="#" data-i18n="[html]menu.colgate.titulo">&raquo;Usuarios Colgate</a>
							<ul>
								<!-- Inicia Link SemiEdi-->	
								<%
									Call semiEDI(trim(user),trim(pass),trim(Nombre),"loginPaebsa.asp?ln="&lg)
								%>
								<!-- Termina Link SemiEdi-->						 
							 </ul>
						</li>
						<!-- Menu Colgate -->
						<!-- Link de SemiEdiColgate-->
						<!-- Link de Facturas express -->		
						
						<!-- Fin link -->
						<li><a href="InfoReceivedSupplier.asp?ln=<%=lg%>" data-i18n="menu.enviada">&raquo;Informaci&oacuten enviada a clientes </a></li>
						<!-- Link de Facturas express -->
						<ul>
						<li>
						<%
							Call facturaExpress(pass,user,Nombre)
						%>	
						</li>
						<!-- Link de facturas express -->

						<!-- Modulo_Nestle_Casa_Ley -->
						<li>
						<%
							Call Modulo_Nestle_Casa_Ley(user,pass,Nombre, lg)
						%>	
						</li>	
						<!-- Modulo_Nestle_Casa_Ley -->
						<!-- Modulo_ARS_Nestle -->
						<li>
						<%
							Call Modulo_ARS_Nestle(user, pass, Nombre, lg, "SPOKE")
						%>	
						</li>	
						<!-- Modulo_ARS_Nestle -->
						
					
						<li>
						<%
							Call reporteBitacoras(pass,user,Nombre)%>
						</li>
						<li>
						<% Call reporteExcel(pass,user,Nombre) %>
						</li>
						<!-- Fin link -->
						</ul>
						<!-- Inicia Link de generacion de archivo ASN ALMGARCIA -->
								<li  style="display:<% if trim(user)="CIVSA" or trim(user)="MXG1505" or (trim(user) = "MXG1397" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXGU435" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXG2004" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) then response.write "block" else response.write "none" end if %>">
								<%
										Call ASNAlmGarcia(trim(pass),trim(user),trim(Nombre),"loginPaebsa.asp?ln="&lg)
										
								%> 
							</li>
								<!-- Termina Link de generacion de archivo ASN ALMGARCIA -->
								<!-- Inicia Link de administrar brokers  ALMGARCIA -->
								<li  style="display:<% if trim(user)="CIVSA" then response.write "block" else response.write "none" end if %>">
								<%
										Call Brokers(trim(pass), trim(user),trim(Nombre))
										
								%> 
								
							</li>

							<li>
								<%
									Call CargaFragua(pass,user,Nombre)		
								%>
							</li>
							
							<li>
								<%
									Call CargaProductosExcel(rtrim(Nombre),"loginPaebsa.asp?ln="&lg, rtrim(user), rtrim(tipoUser), pass)		
								%>
							</li>

							<li>
								<%
									Call CargaImssExcel(rtrim(user), rtrim(pass), rtrim(Nombre), "", "loginPaebsa.asp?ln="&lg)		
								%>
							</li>
							<li>
								<%
									Call CargaTiendas(rtrim(user), rtrim(Nombre), "ADMIN", "loginPaebsa.asp?ln="&lg)		
								%>
							</li>

						<!-- Termina Link de administrar brokers  ALMGARCIA -->
						<li class="has-sub"><a href="#" data-i18n="menu.addendas">&raquo; Captura de Adendas</a>
							<ul>
								<!-- Link de Facturas Walmart -->
								<li class="has-sub"><a href="#" data-i18n="menu.adendaWalmart">&raquo;Addendas de Wal-Mart</a>
								<ul>
									<li>
									<%
										Call AddendaWalmartEdi(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>
									</li>
									<li>
									<%
										Call addendaWalmartResguardo(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>
									</li>
								</ul>
								</li>
								<!-- Link de facturas Walmart -->								
								<!-- Link de envio de facturas con addenda de amazon-->
								<li>
									<%
										Call addendaAmazon(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
										<!-- Fin link -->							
										<!-- Link de envio de facturas con addenda de BB&B-->
								<li>
									<%							
										Call  addendaEdiBBB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
										<!-- Fin link -->
										<!-- Link de envio de facturas con addenda de almacenes Garcia-->
								<li>
									<%							
										Call addendaAlmacenesGarcia(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
								<!-- Fin link -->
								<!-- Inicia Addenda de MERZA -->
								<li>
									<%
										Call addendaMerza(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)	
									%>
								</li>
								<!-- Termina Addenda de Merza -->
								<!-- Inicia Addenda de Corvi -->
								<li>
								<%
									Call addendaCorvi(pass,user,"","loginPaebsa.asp?ln="&lg)		
								%>
								</li>
								<li>
								<%
									Call addendaChedraui(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
								%>
								</li>
								<li>
								<%
									Call AddendaHEB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
								%>
								</li>
								
								<!--Inicia Addenda SORIANA ADMIN -->
								 <!--.functions.asp-->
							    <li class="has-sub"><a href="#" >Addendas SORIANA</a>
									<ul>
										<li>
											<%
												Call Addenda_SORIANA_Ext(pass,user,"",Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")
											%>
										</li>
									</ul>
								</li>        
								<!--Termina Addenda SORIANA ADMIN-->
								
							</ul>
						</li>
							<!-- Termina Captura de Addenda -->
							
					</ul>	
		        </ul>			
		




			
			
			
         </li>
      </ul>
    </div>
  </div>
</nav>


















	<div class="content_loading"  id="content_loading"></div>
	<iframe id="iframe" style="display:none;"></iframe>
	<div class="contenidoGral">
        <div class="links">
		
            <div class="session">
		        <div id="msgSesion" style="color:#B40431;">
			        <label data-i18n="sistema.sesion"> Su sesi&oacute;n expira en: </label><span id="TimeLeft"></span>
			        <script type="text/javascript" language="javascript"> 
				        showTimer();
			        </script>
		        </div>
	        </div>
			
	        <div class="enlaces">
		        <a target="_blank" data-i18n="[html]sistema.manual;[href]sistema.hrefManual">Manual de usuario</a>  |  <a id="btnCerrarSesion" href="Cerrar_Ses_Cli.asp" data-i18n="[html]sistema.cerrarSesion">Cerrar sesi&oacute;n</a>
	        </div>
			<!-- Inicia Cambio de Password -->
			<%
			' Inicia - Se muestra la notificacion al usuario, cuando solo le quedan 10 para que expira su contraseña.
			
			' Consulta Generica SQL
			dim Con_Sql
			' Fecha Actual
			dim Fec_Act
			' Fecha Cambio de password 
			dim Fec_Cam_Pass
			' Dias de diferencia entre fecha 
			dim Dia_Dif_Fec
			' Periodo de expiracion Contrasenia
			dim Per_Exp_Con
			' Dias para que expire la contrasenia
			dim Dia_Exp_Con
			
			' Inicia - Consulta SQL, para obtener la fecha del ultimo cambio realizado del Password
			Con_Sql = " select "&_ 
					  " LEFT(Fecha_Ultimo_Cambio_Pwd,4)+'-'+SUBSTRING(Fecha_Ultimo_Cambio_Pwd,5,2)+'-'+RIGHT(Fecha_Ultimo_Cambio_Pwd,2)  as Fecha_Ultimo_Cambio_Pwd, Periodo_Expiracion_Pwd "&_ 
					  " from CATCLIENTES where Id_Cliente='"&trim(user)&"'"
					  
			set RS_Gen =server.createobject("ADODB.Recordset") 						
			RS_Gen.Open Con_Sql, cnn 
				Fec_Cam_Pass = trim((RS_Gen.fields("Fecha_Ultimo_Cambio_Pwd")  & " "))
				Per_Exp_Con= trim((RS_Gen.fields("Periodo_Expiracion_Pwd")  & " "))
			RS_Gen.close
			' Termina - Consulta SQL, para obtener la fecha del ultimo cambio realizado del Password
	 
			Fec_Act = Fecha_Formato_AnioMesDia("YYYY-MM-DD")
			Dia_Dif_Fec=DateDiff("d", Fec_Cam_Pass, Fec_Act)
			Dia_Exp_Con = CInt(Per_Exp_Con) -CInt(Dia_Dif_Fec)
	 
			 if CInt(Dia_Exp_Con) <= 10 then
				'Sistema de contraseñas viejo
				response.write "<div class='alertContrasena'>"&_
								"<img src='imagenes/Cam_Pass.png'/>"&_ 
								"<label class='TextoCambioPass' id='cPass' data-i18n='expira' data-i18n-options={'dias':'"&Dia_Exp_Con&"'}></label>"&_
								"<span class='TextoCambioPass'>"&Dia_Exp_Con&"</span> "&_
								"<a href='CambioPassword.asp?ln="&lg&"' data-i18n='[html]cambiar'> Cambiar ahora</a> "&_
								"<img src='imagenes/Cam_Pass.png'/>"&_
								"</div>"
				
				response.write "<script> (function() { "&_
								"setInterval(function(){  var el = document.getElementById('cPass'); if(el.className==='TextoCambioPass'){ el.className = 'TextoCambioPass on';}else{ el.className = 'TextoCambioPass'; } },700);	})(); "&_
								"</script>"
			 end if
				' Termina - Se muestra la notificacion al usuario, cuando solo le quedan 10 para que expira su contraseña.
			%>
			<!-- Inicia Cambio de Password -->
	</div>
    
	<% 
	If cant_paginas = 0 and condicion<>2 Then
	' Si la cantidad de paginas es igual a cero y el cliente
	' no tiene informacion registrada.
	%>
	<!-- Inicia - Sin informacion -->
	<div id="templatemo_outer_wrapper_sp" style="margin-top:1px;">
		<div id="templatemo_wrapper_sp"><!-- end of templatemo header -->
			<div class="filtros-busqueda">
				<h2><strong data-i18n="filtros.tituloFiltros"> Filtros de b&uacute;squeda</strong></h2><br/>
				<form name="formulario" action="loginPaebsa.asp" id="formInscripcion" method="get">
						 <select  name="seleccione" id="seleccione" class="select-text select-opt" >
						   <option value="" selected="selected" data-i18n="filtros.seleccione.seleccion">Seleccione</option>
						   <option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
						   <option value="Numero_Proveedor_Hub" data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
						   <option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
						   <option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute;n </option>
						   <option value="Status" data-i18n="filtros.seleccione.estado">Estado</option>
						   <option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
						 </select><a class="tooltip" title="[!]Importante[/!]Seleccione una opci&oacute;n" data-i18n="[title]filtros.seleccione.infoSeleccion"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
						 <input  class="captura busqueda"  name="texto" type="text"  id="texto" size="15" placeholder="Valor obligatorio" data-i18n="[placeholder]filtros.seleccione.captura" />
						 <a class="tooltip" title="[!]Importante[/!]Por favor escriba un texto" data-i18n="[title]filtros.seleccione.infoCaptura"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
						<!-- Campo de busqueda alternativo -->
						<br/><br/>
						 <select  name="seleccione2" id="seleccione2" class="select-text select-opt"  >
						   <option value="" selected="selected" data-i18n="filtros.seleccione.seleccion">Seleccione</option>
						   <option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
						   <option value="Numero_Proveedor_Hub" data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
						   <option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
						   <option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute; n </option>
						   <option value="Status" data-i18n="filtros.seleccione.estado">Estado</option>
						   <option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
						   <!--<option value="Fecha_Recepcion_Sistema">Fecha documento</option>-->
						   <!--<option value="Fecha_Canc_Documento_Edi">Fecha cancelaci&oacuten documento</option>-->
						 </select><a class="tooltip" title="[!]Opcional[/!]Este es un campo opcional para agregar otro valor de b&uacute;squeda" data-i18n="[title]filtros.seleccione.infoSeleccionDos"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
						 <input  class="captura busqueda"  name="texto2" type="text"  id="texto2" size="15" placeholder="Valor opcional" data-i18n="[placeholder]filtros.seleccione.capturaDos" />
					  <a class="tooltip" data-i18n="[title]filtros.seleccione.infoCapturaDos"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
						<!-- Fin del campo de busqueda alternativo -->
						<br/><br/> 
						<select name="orden"  id="orden" class="select-text select-opt" >
							<option value="" data-i18n="filtros.ordenar.resultados">Ordenar resultados por</option>
							<option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
							<option value="Numero_Proveedor_Hub" data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
							<option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
							<option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute;n </option>
							<option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
							<option value="Fecha_Envio_Sistema" data-i18n="filtros.ordenar.fecha">Fecha documento</option>
							<option value="Fecha_Canc_Documento_Edi" data-i18n="filtros.ordenar.fechaCancelacion">Fecha cancelaci&oacute;n documento </option>
						   <!--<option value="Consecutivo_Int_Pebsa">Consecutivo int PAEBSA</option>-->
						 </select><a class="tooltip" title="[!]Importante[/!]Seleccione el orden" data-i18n="[title]filtros.ordenar.info"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripción de Nombre" /></a>
						 <select name="alf"  id="alf" class="select-text select-opt" >
							<option value="desc" data-i18n="filtros.ordenar.descendente">Orden descendente</option>
							<option value="asc" data-i18n="filtros.ordenar.ascendente">Orden ascendente</option>
						 </select>
						 <br/><br/>
						 <select name="tipofecha"  id="tipofecha" class="select-text select-opt" >
						   <option value="Fecha_Recepcion_Sistema" data-i18n="filtros.ordenar.fecha">Fecha documento</option>
						   <option value="Fecha_Canc_Documento_Edi" data-i18n="filtros.ordenar.fechaCancelacion">Fecha cancelaci&oacute;n documento </option>
							<option value="Fecha_Consulta_Cliente" data-i18n="filtros.ordenar.fechaConsulta">Fecha consulta</option>
						 </select>
						 <input  class="captura fecha" placeholder="Fecha inicial" data-i18n="[placeholder]filtros.fecha.fechaInicial"  type="text" id="datepicker" name="datepicker" />
						 <input  class="captura fecha" placeholder="Fecha final" data-i18n="[placeholder]filtros.fecha.fechaFinal"  type="text" id="datepickerfinal" name="datepickerfinal" />
						 <a class="tooltip" data-i18n="[title]filtros.fecha.info" ><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
						<br/><br/> 
						<!-- inicia nuevo campo de registros por pagina -->
						<select  name="tamanopagina" id="tamanopagina" class="select-text select-opt">
							<option value="25" selected="selected" data-i18n="filtros.pagina.numero">N&uacute;mero de registros por p&aacutegina </option>
							<option value="25">25</option>
							<option value="50">50</option>
							<option value="75">75</option>
							<option value="100">100</option>
							<option value="200">200</option>
						 </select><a class="tooltip" title="[!]Opcional[/!]Este campo es el n&uacutemero de registros a visualizar por p&aacute;gina (25 por default)" data-i18n="[title]filtros.pagina.info"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
						<!-- termina el nuevo campo de paginas por busqueda-->
						<br/><br/>
						<div class="input"><input class="button_opt prtText" name="Submit" disabled type="submit" value="Buscar" id="btnBuscar" data-i18n="[value]filtros.botones.buscar"/></div>
						<div class="input"><input class="button_opt prtText" name="button" disabled type="button"  value="Restablecer" id="btnRestablecer" data-i18n="[value]filtros.botones.restablecer"/></div>
					</form>
			</div>
			
			<div class="bitacora-datos">
				<h2><strong data-i18n="filtros.tituloBitacoras">Informaci&oacute;n sobre la bit&aacute;cora de datos</strong></h2><br/>
				<ul class="lista">
					<li class="go"><img src="imagenes2/negro.png" alt="PAEBSA"  class="imagenNegro"/><label data-i18n="bitacora.noConsultado">Archivo no consultado</label></li>
					<li class="go"><img src="imagenes2/azul.png" alt="PAEBSA" class="imagenAzul"/><label data-i18n="bitacora.consultado">Archivo consultado</label></li>
					<li class="go"><img src="imagenes2/rojo.png" alt="PAEBSA" class="imagenRojo"/><label data-i18n="bitacora.depuracion"> Archivo preparado a depuraci&oacute;n</label></li>   
					<li class="go"><label data-i18n="bitacora.sinInformacion.pagina">P&aacute;gina actual: 0</label></li>
					<li class="go"><label data-i18n="bitacora.sinInformacion.registros">Registros por página: 0</label></li>
					<li class="go"><label data-i18n="bitacora.sinInformacion.cantidad">Cantidad de páginas: 0</label></li>
					<li class="go"><label data-i18n="bitacora.sinInformacion.totales">Registros totales: 0</label></li>
					<li class="go"><a href="InfoReceivedSupplier.asp?ln=<%=lg%>" data-i18n="[html]bitacora.informacionEnviada"> Informaci&oacute;n enviada a clientes</a></li>
					<li class="go"><a href="loginPaebsa.asp?ln=<%=lg%>" data-i18n="[html]bitacora.consulta">Consulta general </a></li>
					<li class="go"><a id="btnSalir" href="Cerrar_Ses_Cli.asp" data-i18n="[html]sistema.contenido.enlace">Cerrar sesi&oacute;n</a></li>
				</ul>
			</div>
 
		<!-- Mensajes a clientes -->
		<div class="post_box">
			<div class="slideshow" style="position: relative; z-index: 1;">
				<%
					AvisoGenerico(user)
				%>		
			</div>
			<div class="slideshow" style="position: relative; z-index: 1;">
				<%
					mensajeCliente(user)
				%>		
			</div>
		</div>
		<!-- Mensajes a clientes -->
			
		</div>
		<div class="content_menu">
		<div id="menu">
			<dt id="TituloMenu" class="tituloMenu" data-i18n="menu.titulo">Nuevas funciones del portal</dt>				 
			<ul id="ListaMenu" class="lista">
				<li id="composebtn">
							<a href="#"  class="compose" id="composeicon" data-i18n="menu.factura.titulo"> &raquo;Env&iacuteo de facturas a clientes</a>
								<div class="mainCompose">
									<div class="calloutUp">
										<div class="calloutUp2"></div>
									</div>	
									<div id="msgform" class="msgEnvio" width="700px">
										<form id="sendprivatemsg" class="UsuariosCss" action="EnvioXML/ValidaXML.aspx" method="post" enctype="multipart/form-data">
											<label data-i18n="menu.factura.xml">Factura XML/EDI</label>
											<input type="file" name="archivo[]" accept="text/xml,.edi" size="70" multiple value="Examine"/>
											<br /><br />
											<label style="color:#B40404;" data-i18n="menu.factura.archivos">N&uacutemero m&aacuteximo de archivos por carga: 50</label>
											<br /><br/>
											<label style="color:#0B4C5F;" data-i18n="menu.factura.aviso">IMPORTANTE: Para enviar facturas con addenda resguardo de Walmart vaya a "Captura de Addendas-> Addendas de Wal-Mart-> Addenda Resguardo"</label>
											<br /><br />
											<%
												sqlProveedorMerza = "select rtrim(id_cliente)id_Cliente, Codigo_Cliente,Codigo_Transaccion_Produccion,RFCSpoke,RFCHub from CATSPOKESHUBS where Codigo_Cliente='"&trim(pass)&"' and Id_Cliente='"&trim(user)&"' and Codigo_Transaccion_Produccion='INVOIC' and RFCHub='ADU800131T10'"
												set rsProveedor=server.createobject("ADODB.Recordset") 						
												rsProveedor.Open sqlProveedorMerza,cnn,3,1	
												if rsProveedor.EOF then
												else
												response.write "<label  style='color:#B40404;' data-i18n='menu.factura.avisoMerza'> SI ERES PROVEEDOR DE MERZA, FAVOR DE SUBIR FACTURAS CON LA ADDENDA SOLICITADA</label></a><br/><br/>"
												end if

											%>
											<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
											<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
											<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
											<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>
											<div style="padding-bottom: 25px;">
												<div class="input" style="float:right;">
													<input class="button_opt prtText" onclick="return validarMaximoArchivos()" type="submit" id="Submit1" value="Enviar facturas" data-i18n="[value]menu.factura.boton" />
												</div>
											</div>
											<br /><br />
										</form>
									  </div>
								</div>
								<!-- termina cuadro de dialogo de facturas -->
						</li>
						
						<!-- Modulo_Genera_Addenda_Nube -->
						<li class="has-sub"><a href="#" data-i18n="[html]menu.generarAddendaAutomatica.titulo"> &raquo; Generar Addenda en la nube</a>
							<ul>
								<li>
								<% Call CargaDeAddendaGenerica(pass,user,Nombre,"loginPaebsa.asp?ln="&lg) %><br />
								</li>			
							</ul>
						</li>
						<!-- Modulo_Genera_Addenda_Nube -->
						
						
						<!-- Modulo_Envio_EDI_Clientes -->
						<li id="btnEdi">
						<!-- Inicia cuadro de dialogo de archivos EDI -->
							<a href="#" class="composeEdi" id="composeiconEdi" data-i18n="menu.asn.titulo"> &raquo;Env&iacuteo de archivos DESADV</a>
								<div class="mainEDI">
									<div class="calloutUp">
										<div class="calloutUp2"></div>
									</div>	
									<div id="msgformEDI" class="msgEnvio" width="700px">
										<form id="sendprivatemsgEdi" class="UsuariosCss" action="AplicacionPaebsa/ValidaXML.ashx" method="post" enctype="multipart/form-data">
											<label data-i18n="menu.asn.envio">Archivos ASN(.edi) </label>
											<input type="file" name="archivoEdi[]" id="archivoEdi" accept="text/edi" size="70" multiple />
											<br /><br />
											<label style="color:#B40404;" data-i18n="menu.asn.aviso">N&uacutemero m&aacuteximo de archivos por carga: 10</label>
											<br />
											<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
											<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
											<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
											<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>									  
											<div style="padding-bottom: 25px;">
												<div class="input" style="float:right;">
													<input class="button_opt prtText" onclick="return validarMaximoArchivosEdis()" type="submit" id="btnenviafac" value="Enviar archivos" data-i18n="[value]menu.asn.boton"/>
												</div>
											</div>
											<br /><br />
										</form>
									</div>
								</div>
							<!-- termina cuadro de dialogo de archivos EDI -->	
						</li>
						<!-- Modulo_Envio_EDI_Clientes -->
						
						<li id="link_cargaInfo"><a href="#" onClick="openBrowser('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="menu.cargaInformacion">&raquo;Carga de informaci&oacute;n</a></li>
						<li><a href="#" onClick="openNoSpots('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="menu.cargaNoSpots">&raquo;Carga de NO SPOTS</a></li>
						
						<li class="has-sub">
						<a href="#" data-i18n="menu.administrarCuenta.titulo">&raquo;Administre su cuenta</a>
							<ul>
								<li><a href="RegistroUsuarios.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.usuarios">&raquo;Administre sus usuarios</a></li>
								<!--Sistema viejo de contraseñas-->
								<li><a href="CambioPassword.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.contrasena"> &raquo;Cambiar contraseña</a></li>
								<!--Sistema nuevo de contraseñas>>>Quitar comentario y comentar línea de arriba-->
								<!--<li><a href="AplicacionPaebsa/ReestablecerContrasena.aspx?tipoUsr=M&pagina=loginPaebsa.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.contrasena"> &raquo;Cambiar contraseña</a></li>-->
								<li><a href="#" id="modal"  onClick="ventanaHistorial();" data-i18n="menu.administrarCuenta.historial">&raquo;Historial de usuarios</a></li>
							</ul>
						</li>
						
						
						
						<li><a href="loginPaebsa.asp?ln=<%=lg%>" data-i18n="menu.general">&raquo;Consulta general</a></li>
	 
						<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
						<li id="link_desadv">
							<a href="#" onClick="openTemplate('<%=trim(user)%>','ADMIN')" data-i18n="menu.template">&raquo;Captura de confirmaci&oacute;n para los templates de Walmart/Sahuayo (DESAV)</a>
						</li>
						<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
	 
						<!-- Menu Colgate -->
						<li class="has-sub" style="display:<% if trim(user)="CPA7503043P1" then response.write "block" else response.write "none" end if %>"><a title="" href="#" data-i18n="[html]menu.colgate.titulo">&raquo;Usuarios Colgate</a>
							<ul>
								<!-- Inicia Link SemiEdi-->	
								<%
									Call semiEDI(trim(user),trim(pass),trim(Nombre),"loginPaebsa.asp?ln="&lg)
								%>
								<!-- Termina Link SemiEdi-->						 
							 </ul>
						</li>
						<!-- Menu Colgate -->
						<!-- Link de SemiEdiColgate-->
						<!-- Link de Facturas express -->		
						
						<!-- Fin link -->
						<li><a href="InfoReceivedSupplier.asp?ln=<%=lg%>" data-i18n="menu.enviada">&raquo;Informaci&oacuten enviada a clientes </a></li>
						<!-- Link de Facturas express -->
						<li>
						<%
							Call facturaExpress(pass,user,Nombre)
						%>	
						</li>
						<!-- Link de facturas express -->

						<!-- Modulo_Nestle_Casa_Ley -->
						<li>
						<%
							Call Modulo_Nestle_Casa_Ley(user,pass,Nombre, lg)
						%>	
						</li>	
						<!-- Modulo_Nestle_Casa_Ley -->
						<!-- Modulo_ARS_Nestle -->
						<li>
						<%
							Call Modulo_ARS_Nestle(user, pass, Nombre, lg, "SPOKE")
						%>	
						</li>	
						<!-- Modulo_ARS_Nestle -->
						
					
						<li>
						<%
							Call reporteBitacoras(pass,user,Nombre)%>
						</li>
						<li>
						<% Call reporteExcel(pass,user,Nombre) %>
						</li>
						<!-- Fin link -->
						
						<!-- Inicia Link de generacion de archivo ASN ALMGARCIA -->
								<li  style="display:<% if trim(user)="CIVSA" or trim(user)="MXG1505" or (trim(user) = "MXG1397" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXGU435" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXG2004" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) then response.write "block" else response.write "none" end if %>">
								<%
										Call ASNAlmGarcia(trim(pass),trim(user),trim(Nombre),"loginPaebsa.asp?ln="&lg)
										
								%> 
							</li>
								<!-- Termina Link de generacion de archivo ASN ALMGARCIA -->
								<!-- Inicia Link de administrar brokers  ALMGARCIA -->
								<li  style="display:<% if trim(user)="CIVSA" then response.write "block" else response.write "none" end if %>">
								<%
										Call Brokers(trim(pass), trim(user),trim(Nombre))
										
								%> 
								
							</li>

							<li>
								<%
									Call CargaFragua(pass,user,Nombre)		
								%>
							</li>
							
							<li>
								<%
									Call CargaProductosExcel(rtrim(Nombre),"loginPaebsa.asp?ln="&lg, rtrim(user), rtrim(tipoUser), pass)		
								%>
							</li>

							<li>
								<%
									Call CargaImssExcel(rtrim(user), rtrim(pass), rtrim(Nombre), "", "loginPaebsa.asp?ln="&lg)		
								%>
							</li>
							<li>
								<%
									Call CargaTiendas(rtrim(user), rtrim(Nombre), "ADMIN", "loginPaebsa.asp?ln="&lg)		
								%>
							</li>

						<!-- Termina Link de administrar brokers  ALMGARCIA -->
						<li class="has-sub"><a href="#" data-i18n="menu.addendas">&raquo; Captura de Adendas</a>
							<ul>
								<!-- Link de Facturas Walmart -->
								<li class="has-sub"><a href="#" data-i18n="menu.adendaWalmart">&raquo;Addendas de Wal-Mart</a>
								<ul>
									<li>
									<%
										Call AddendaWalmartEdi(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>
									</li>
									<li>
									<%
										Call addendaWalmartResguardo(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>
									</li>
								</ul>
								</li>
								<!-- Link de facturas Walmart -->								
								<!-- Link de envio de facturas con addenda de amazon-->
								<li>
									<%
										Call addendaAmazon(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
										<!-- Fin link -->							
										<!-- Link de envio de facturas con addenda de BB&B-->
								<li>
									<%							
										Call  addendaEdiBBB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
										<!-- Fin link -->
										<!-- Link de envio de facturas con addenda de almacenes Garcia-->
								<li>
									<%							
										Call addendaAlmacenesGarcia(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
									%>	
								</li>							
								<!-- Fin link -->
								<!-- Inicia Addenda de MERZA -->
								<li>
									<%
										Call addendaMerza(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)	
									%>
								</li>
								<!-- Termina Addenda de Merza -->
								<!-- Inicia Addenda de Corvi -->
								<li>
								<%
									Call addendaCorvi(pass,user,"","loginPaebsa.asp?ln="&lg)		
								%>
								</li>
								<li>
								<%
									Call addendaChedraui(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
								%>
								</li>
								<li>
								<%
									Call AddendaHEB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
								%>
								</li>
								
								<!--Inicia Addenda SORIANA ADMIN -->
								 <!--.functions.asp-->
							    <li class="has-sub"><a href="#" >&raquo;Addendas SORIANA</a>
									<ul>
										<li>
											<%
												Call Addenda_SORIANA_Ext(pass,user,"",Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")
											%>
										</li>
									</ul>
								</li>        
								<!--Termina Addenda SORIANA ADMIN-->
								
							</ul>
						</li>
							<!-- Termina Captura de Addenda -->
							
					</ul>				
		</div>
		<!-- inicia script de acordeon-->		
			<script type="text/javascript">
					$('#ListaMenu').hide();
					$('#ListaMenu').removeClass('activo');
					$('#TituloMenu').click(function()
					{
						var c = $("#ListaMenu");
						var mostrandose = c.css("display");
						if (mostrandose=="block"){
							$("#ListaMenu").slideUp()
						}else{
							$("#ListaMenu").slideDown("slow");
						}
					});
			</script>						
		<!-- termina script de acordeon -->
		</div>
	</div>
		<div class="container_12 divider">
			<div class="info_options">
				<form action="ficheroExcel.php" method="post"  id="FormularioExportacion">
					<div style="">
					 
					<div class="input"><input type="button"  value="Reprocesar archivos"  class="button_opt prtText" id="btnReprocesoEdi" data-i18n="[value]funcionalidad.reproceso"/></div>
					<a class="tooltip" title="" data-i18n="[title]funcionalidad.infoReproceso"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					<div class="input"><input type="button"  value="Reprocesar PDF" class="button_opt prtText" id="btnReprocesoPDF" data-i18n="[value]funcionalidad.reprocesoPDF"/></div>
					<a class="tooltip" title=""  data-i18n="[title]funcionalidad.infoReprocesoPDF"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>									  
					</div>
					<div style="">
					<div class="input"><input  class="button_opt prtText create-user btn_download" type="button" value="Descarga masiva de archivos" id="btnDescargaM"/></div>
					<a class="tooltip" title="" data-i18n="[title]funcionalidad.descargaMasiva"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					<div class="input"><input type="button" value="Enviar informaci&oacute;n por e-mail" class="button_opt prtText" id="btnEmail" data-i18n="[value]funcionalidad.email"/></div>
					<a class="tooltip" title="" data-i18n="[title]funcionalidad.infoEmail"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					<div class="input"><input type="button"  value="Exportar datos a un excel"  class="button_opt prtText" id="btnExcel" data-i18n="[value]funcionalidad.excel"/></div>
					<a class="tooltip" title="" data-i18n="[title]funcionalidad.infoExcel"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					</div>
				</form>
			</div>
			<div class="grid_8 height250">
				<!-- Principia la tabla vacia-->
				<table cellpadding="0" cellspacing="0" border="0" id="myTable02" class="tinytable">
					<thead>
						<tr>
							<th class="nosort"></th>
							<th><input id="cTodos" name="checkbox" type="checkbox"/></th>
							<th><h3 data-i18n="grid.nombre">Nombre cadena</h3></th>
							<th><h3 data-i18n="grid.noProveedor">No. de proveedor cadena</h3></th>
							<th><h3 data-i18n="grid.transaccion">C&oacute;digo de transacci&oacute;n </h3></th>
							<th><h3 data-i18n="grid.noDocumento">No. de documento</h3></th>
							<th><h3 data-i18n="grid.fechaHora">Fecha y hora de consulta </h3></th>
							<!--<th><h3>Fecha de publicaci&oacuten </h3></th>
							<th><h3>Hora de publicaci&oacuten </h3></th>-->
							<th><h3 data-i18n="grid.fechaCancelacion">Fecha cancelaci&oacuten documento </h3></th>
							<th><h3 data-i18n="grid.fechaDocumento">Fecha documento </h3></th>
							<th><h3 data-i18n="grid.claveCliente">Clave cliente </h3></th>
							<th><h3 data-i18n="grid.noControl">No. de control</h3></th>
							<th><h3 data-i18n="grid.estado" class="sizeTittle60">Estado</h3></th>
							<th><h3 data-i18n="grid.codigoTienda" class="sizeTittle60">C&oacute;digo tienda </h3></th>
							<th><h3 data-i18n="grid.descripcion">Descripci&oacute;n del proceso </h3></th>
							<th><h3 data-i18n="grid.descargar" class="sizeTittle40">Descargar</h3></th>
						</tr>
					</thead>
					<tbody>
						<tr>
						<td colspan="8" style="vertical-align: top;height:400px;padding-top: 25px;"><label style="font-size:18px;font-style:italic;" data-i18n="grid.informacion">Sin Informaci&oacute;n </label></td>
						</tr>
					</tbody>
				</table>
				<!-- Termina la tabla vacia-->		
			</div>	
		</div>
	<!-- Termina - Sin informacion -->
	<%
		' Si la cantidad de paginas es igual a cero y el cliente
		' no tiene informacion registrada.
	else
	%>
	
	
		<!-- Inicia - Sin Info -->
		<% If cant_paginas = 0 and condicion=2  Then %>
				<!-- Inicia - Si la cantidad de paginas es igual a cero y fue por una busqueda de informacion que no trajo resultados -->
				<div class="condicion">
					<div class="imagen2">
						<img src="imagenes/Paebsa.png" alt="PAEBSA" width="291" height="226" class="image_fc"/>
					</div>
					<p class="centro" data-i18n="grid.busqueda">No se encontraron resultados de b&uacute;squeda </p>
					<a href="loginPaebsa.asp?ln=<%=lg%>"  data-i18n='[html]grid.enlace'>Buscar nuevamente</a> 		
				</div>
				<!-- Termina - Si la cantidad de paginas es igual a cero y fue por una busqueda de informacion que no trajo resultados -->
		<% else %>
		<!-- Termina - Sin Info -->

		<%
		
		if tamanopagina <> "all" then
			rs.pagesize= cint(tamanopagina)
				rs.absolutepage=cint(paginaabsoluta)
		contador=1

			dim matriz (200)
				PageSize=rs.PageSize
					
					 for i=1 to  PageSize
						 matriz(i)="'fila"&i&"',"
						 
					 if i=pageSize then
				 matriz(i)="'fila"&i&"'"
			 end if
			next
            avisos()
		%>
		
		<script type="text/javascript">
            var once_per_browser=0
            var ns4=document.layers
            var ie4=document.all
            var ns6=document.getElementById&&!document.all
            if (ns4)
                crossobj=document.layers.divAvisoG
            else if (ie4||ns6)
                crossobj=ns6? document.getElementById("divAvisoG") : document.all.divAvisoG
            function closeit()
            {
                if (ie4||ns6)
                    crossobj.style.visibility="hidden"
                else if (ns4)
                    crossobj.visibility = "hide"
            }
            function get_cookie(Name) {
                var search = Name + "="
                var returnvalue = "";
                if (document.cookie.length > 0) {
                    offset = document.cookie.indexOf(search)
                    if (offset != -1) { // if cookie exists
                        offset += search.length
                        end = document.cookie.indexOf(";", offset);
                        if (end == -1)
                            end = document.cookie.length;
                            returnvalue = unescape(document.cookie.substring(offset, end))
                    }
                }
                return returnvalue;
            }
            function showornot(){
                if (get_cookie('postdisplay')==''){
                    showit()
                    document.cookie = "postdisplay=yes"
                }
            }
            function showit(){
	            if (crossobj!=null){
		            if (ie4||ns6)
			            crossobj.style.visibility="visible"
		            else if (ns4)
			            crossobj.visibility = "show"
	            }
            }
            if (once_per_browser)
                showornot()
            else
                showit()



            function drag_drop(e){
                if (ie4&&dragapproved){
                    crossobj.style.left=tempx+event.clientX-offsetx+'px'
                    crossobj.style.top=tempy+event.clientY-offsety+'px'
                    return false
                }
                else if (ns6&&dragapproved){
                    crossobj.style.left=tempx+e.clientX-offsetx+'px'
                    crossobj.style.top=tempy+e.clientY-offsety+'px'
                    return false
                    }
            }
            function initializedrag(e){
                if (ie4&&event.srcElement.id=="divAvisoG"||ns6&&e.target.id=="divAvisoG"){
                    offsetx=ie4? event.clientX : e.clientX
                    offsety=ie4? event.clientY : e.clientY
                    tempx=parseInt(crossobj.style.left)
                    tempy=parseInt(crossobj.style.top)
                    dragapproved=true
                    document.onmousemove=drag_drop
                }
            }
            document.onmousedown=initializedrag
            document.onmouseup=new Function("dragapproved=false")

		</script>
<%
	
%>
 
<div id="templatemo_outer_wrapper_sp" style="margin-top:1px;">
  <div id="templatemo_wrapper_sp"><!-- end of templatemo header -->
		<div class="filtros-busqueda">
			<h2><strong data-i18n="filtros.tituloFiltros"> Filtros de b&uacute;squeda</strong></h2>
			<form name="formulario" action="loginPaebsa.asp?ln=<%=lg%>" id="formInscripcion" method="get">
					         <select  name="seleccione" id="seleccione" class="select-text select-opt" >
					           <option value="" selected="selected" data-i18n="filtros.seleccione.seleccion">Seleccione</option>
					           <option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
					           <option value="Numero_Proveedor_Hub" data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
					           <option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
					           <option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute;n </option>
					           <option value="Status" data-i18n="filtros.seleccione.estado">Estado</option>
                               <option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
					         </select><a class="tooltip" title="[!]Importante[/!]Seleccione una opci&oacute;n" data-i18n="[title]filtros.seleccione.infoSeleccion"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
					         <input  class="captura busqueda"  name="texto" type="text"  id="texto" size="15" placeholder="Valor obligatorio" data-i18n="[placeholder]filtros.seleccione.captura"/>
				             <a class="tooltip" title="[!]Importante[/!]Por favor escriba un texto" data-i18n="[title]filtros.seleccione.infoCaptura"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					        <!-- Campo de busqueda alternativo -->
					        <br/><br/>
					         <select  name="seleccione2" id="seleccione2" class="select-text select-opt"  >
					           <option value="" selected="selected" data-i18n="filtros.seleccione.seleccion">Seleccione (opcional)</option>
					           <option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
					           <option value="Numero_Proveedor_Hub" data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
					           <option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
					           <option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute;n </option>
					           <option value="Status"  data-i18n="filtros.seleccione.estado">Estado</option>
                               <option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
					         </select><a class="tooltip" title="[!]Opcional[/!]Este es un campo opcional para agregar otro valor de b&uacutesqueda" data-i18n="[title]filtros.seleccione.infoSeleccionDos"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
					         <input  class="captura busqueda"  name="texto2" type="text"  id="texto2" size="15" placeholder="Valor opcional" data-i18n="[placeholder]filtros.seleccione.capturaDos"/>
				          <a class="tooltip" title="[!]Opcional[/!]Campo de b&uacutesqueda opcional para agregar un valor de b&uacutesqueda m&aacutes a su consulta" data-i18n="[title]filtros.seleccione.infoCapturaDos"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					        <!-- Fin del campo de busqueda alternativo -->
					        <br/><br/> 
					        <select name="orden"  id="orden" class="select-text select-opt" >
						        <option value="" data-i18n="filtros.ordenar.resultados">Ordenar resultados por</option>
						        <option value="Nombre_Hub" data-i18n="filtros.seleccione.nombre">Nombre cadena</option>
                                <option value="Numero_Proveedor_Hub"  data-i18n="filtros.seleccione.proveedor">No. proveedor</option>
                                <option value="Num_control_dato_docto" data-i18n="filtros.seleccione.documento">No. de documento</option>
						        <option value="Codigo_Transaccion" data-i18n="filtros.seleccione.transaccion">C&oacute;digo de transacci&oacute;n </option>
                                <option value="Codigo_Tienda" data-i18n="filtros.seleccione.tienda">C&oacute;digo tienda</option>
						        <option value="Fecha_Envio_Sistema" data-i18n="filtros.ordenar.fecha">Fecha documento</option>
						        <option value="Fecha_Canc_Documento_Edi"  data-i18n="filtros.ordenar.fechaCancelacion">Fecha cancelaci&oacute;n documento </option>
					           <!--<option value="Consecutivo_Int_Pebsa">Consecutivo int PAEBSA</option>-->
					         </select><a class="tooltip" title="[!]Importante[/!]Seleccione el orden" data-i18n="[title]filtros.ordenar.info"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripción de Nombre" /></a>
					         <select name="alf"  id="alf" class="select-text select-opt" >
						        <option value="desc" data-i18n="filtros.ordenar.descendente">Orden descendente</option>
						        <option value="asc" data-i18n="filtros.ordenar.ascendente">Orden ascendente</option>
					         </select>
					         <br/><br/>
					         <select name="tipofecha"  id="tipofecha" class="select-text select-opt" >
					            <option value="Fecha_Recepcion_Sistema" data-i18n="filtros.ordenar.fecha">Fecha documento</option>
					            <option value="Fecha_Canc_Documento_Edi" data-i18n="filtros.ordenar.fechaCancelacion">Fecha cancelaci&oacuten documento </option>
					            <option value="Fecha_Consulta_Cliente" data-i18n="filtros.ordenar.fechaConsulta">Fecha consulta</option>
					         </select>
					         <input  class="captura fecha" placeholder="Fecha inicial"  type="text" id="datepicker" name="datepicker" data-i18n="[placeholder]filtros.fecha.fechaInicial"/>
					         <input  class="captura fecha" placeholder="Fecha final"  type="text" id="datepickerfinal" name="datepickerfinal" data-i18n="[placeholder]filtros.fecha.fechaFinal"/>
					         <a class="tooltip" title="[!]Rango de Fechas[/!]Las fechas son datos opcionales, en caso de seleccionar solo una entonces la busqueda se hara de forma especifica  de acuerdo a esa fecha." data-i18n="[title]filtros.fecha.info"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
					        <br/><br/> 
					        <!-- inicia nuevo campo de registros por pagina -->
					        <select  name="tamanopagina" id="tamanopagina" class="select-text select-opt" >
					            <option value="25" selected="selected" data-i18n="filtros.pagina.numero">N&uacutemero de registros por p&aacutegina </option>
					            <option value="25">25</option>
					            <option value="50">50</option>
					            <option value="75">75</option>
					            <option value="100">100</option>
					            <option value="200">200</option>
					         </select><a class="tooltip" title="[!]Opcional[/!]Este campo es el n&uacutemero de registros a visualizar por p&aacutegina (25 por default)" data-i18n="[title]filtros.pagina.info"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>	 
					        <!-- termina el nuevo campo de paginas por busqueda-->
					        <br/><br/>
                            <input type="hidden" name="ln" value="<%=lg%>"/>
					        <div class="input"><input class="button_opt prtText" name="Submit"  type="submit" value="Buscar" id="btnBuscar" data-i18n="[value]filtros.botones.buscar"/></div>
					        <div class="input"><input class="button_opt prtText" name="button" onclick="cancelarFormulariodeBusqueda('loginPaebsa.asp?ln=<%=lg%>')" type="button"  value="Restablecer" id="btnRestablecer" data-i18n="[value]filtros.botones.restablecer"/></div>
				        </form>
		</div>
		
		<div class="bitacora-datos">
			<th><h2><strong data-i18n="filtros.tituloBitacoras">Informaci&oacute;n sobre la bit&aacute;cora de datos</strong></h2><br/></th>
			<ul class="lista">
				<li class="go"><img src="imagenes2/negro.png" alt="PAEBSA"  class="imagenNegro"/><label data-i18n="bitacora.noConsultado"> Archivo no consultado</label></li>
				<li class="go"><img src="imagenes2/azul.png" alt="PAEBSA" class="imagenAzul"/><label data-i18n="bitacora.consultado"> Archivo consultado</label></li>
				<li class="go"><img src="imagenes2/rojo.png" alt="PAEBSA" class="imagenRojo"/><label data-i18n="bitacora.depuracion"> Archivo preparado a depuraci&oacute;n </label></li>   
				<li class="go"><label data-i18n="bitacora.conInformacion.pagina"> P&aacute;gina actual:</label> <%= paginaabsoluta %></li>
				<li class="go"><label data-i18n="bitacora.conInformacion.registros">Registros por p&aacute;gina:</label> <%= rs.PageSize %></li>
				<li class="go"><label data-i18n="bitacora.conInformacion.cantidad">Cantidad de p&aacute;ginas:</label> <%= rs.PageCount %></li>
				<li class="go"><label data-i18n="bitacora.conInformacion.totales">Registros totales:</label> <%= rs.RecordCount %></li>
				<li class="go"><a href="InfoReceivedSupplier.asp?ln=<%=lg%>" data-i18n="[html]bitacora.informacionEnviada"> Información enviada a clientes</a></li>
				<li class="go"><a id="btnSalir" href="Cerrar_Ses_Cli.asp" data-i18n="[html]sistema.contenido.enlace">Cerrar sesi&oacute;n</a></li>
			</ul>
		</div>
			

	<!-- Mensajes a clientes -->
	<div class="post_box">
		<div class="slideshow" style="position: relative; z-index: 1;">
			<%
				AvisoGenerico(user)
			%>		
		</div>
		<div class="slideshow" style="position: relative; z-index: 1;">
			<%
				mensajeCliente(user)
			%>		
		</div>
	</div>	
	<!-- Mensajes a clientes -->
  </div>
  
  <div class="content_menu">
	<div id="menu">
		<dt id="TituloMenu" class="tituloMenu" data-i18n="menu.titulo">Nuevas funciones del portal</dt>	
		<ul id="ListaMenu" class="lista">
		
		
					<!-- Cuadro de dialogo. subir facturas -->
					<li id="composebtn">
						<a href="#"  class="compose" id="composeicon" data-i18n="menu.factura.titulo"> &raquo;Env&iacuteo de facturas a clientes</a>
							<div class="mainCompose">
								<div class="calloutUp">
									<div class="calloutUp2"></div>
								</div>	
								<div id="msgform" class="msgEnvio" width="700px">
									<form id="sendprivatemsg" class="UsuariosCss" action="EnvioXML/ValidaXML.aspx" method="post" enctype="multipart/form-data">
										<label data-i18n="menu.factura.xml">Factura XML/EDI</label>
										<input type="file" name="archivo[]" accept="text/xml,.edi" size="70" multiple value="Examine"/>
										<br /><br />
										<label style="color:#B40404;" data-i18n="menu.factura.archivos">N&uacutemero m&aacuteximo de archivos por carga: 50</label>
										<br /><br/>
										<label style="color:#0B4C5F;" data-i18n="menu.factura.aviso">IMPORTANTE: Para enviar facturas con addenda resguardo de Walmart vaya a "Captura de Addendas-> Addendas de Wal-Mart-> Addenda Resguardo"</label>
										<br /><br />
										<%
											sqlProveedorMerza = "select rtrim(id_cliente)id_Cliente, Codigo_Cliente,Codigo_Transaccion_Produccion,RFCSpoke,RFCHub from CATSPOKESHUBS where Codigo_Cliente='"&trim(pass)&"' and Id_Cliente='"&trim(user)&"' and Codigo_Transaccion_Produccion='INVOIC' and RFCHub='ADU800131T10'"
											'response.write sqlProveedorMerza
											set rsProveedor=server.createobject("ADODB.Recordset") 						
											rsProveedor.Open sqlProveedorMerza,cnn,3,1	
											if rsProveedor.EOF then
											else
											response.write "<label  style='color:#B40404;' data-i18n='menu.factura.avisoMerza'> SI ERES PROVEEDOR DE MERZA, FAVOR DE SUBIR FACTURAS CON LA ADDENDA SOLICITADA</label></a><br/><br/>"
											end if
										%>
										<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
										<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
										<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
										<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>
										<div style="padding-bottom: 25px;">
											<div class="input" style="float:right;">
												<input class="button_opt prtText" onclick="return validarMaximoArchivos()" type="submit" id="Submit1" value="Enviar facturas"  data-i18n="[value]menu.factura.boton" />
											</div>
										</div>
										<br /><br />
									</form>
								  </div>
							</div>
					<!-- Cuadro de dialogo. subir facturas -->
					</li>
					
					
					<!-- Modulo_Genera_Addenda_Nube -->
					<li class="has-sub"><a href="#" data-i18n="[html]menu.generarAddendaAutomatica.titulo"> &raquo; Generar Addenda en la nube</a>
						<ul>
							<li>
								<%Call CargaDeAddendaGenerica(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)%><br />
							</li>			
						</ul>
					</li>
					<!-- Modulo_Genera_Addenda_Nube -->
					
					<!-- Modulo_Envio_EDI_Clientes -->
					<li id="btnEdi">
					<!-- Inicia cuadro de dialogo de archivos EDI -->
						<a href="#" class="composeEdi" id="composeiconEdi" data-i18n="menu.asn.titulo"> &raquo;Env&iacuteo de archivos DESADV</a>
							<div class="mainEDI">
								<div class="calloutUp"> 
									<div class="calloutUp2"></div>
								</div>	
								<div id="msgformEDI" class="msgEnvio" width="700px">
									<form id="sendprivatemsgEdi" class="UsuariosCss" action="AplicacionPaebsa/ValidaXML.ashx" method="post" enctype="multipart/form-data">
										<label data-i18n="menu.asn.envio">Archivos ASN(.edi) </label>
										<input type="file" name="archivoEdi[]" id="archivoEdi" accept="text/edi" size="70" multiple />
										<br /><br />
										<label style="color:#B40404;" data-i18n="menu.asn.aviso">N&uacutemero m&aacuteximo de archivos por carga: 10</label>
										<br />
										<input type="hidden" id="pba" name="pba" value="<%=trim(pass)%>"/> 
										<input type="hidden" id="userBuzon" name="userBuzon" value="<%=trim(user)%>"/>
										<input type="hidden" id="paginaRetornoXML" name="paginaRetornoXML" value="loginPaebsa.asp?ln=<%=lg%>"/>
										<input type="hidden" id="SpokeOhub" name="SpokeOhub" value="spoke"/>									  
										<div style="padding-bottom: 25px;">
											<div class="input" style="float:right;">
												<input class="button_opt prtText" onclick="return validarMaximoArchivosEdis()" type="submit" id="btnenviafac" value="Enviar archivos" data-i18n="[value]menu.asn.boton"/>
											</div>
										</div>
										<br /><br />
									</form>
								</div>
							</div>
					<!-- termina cuadro de dialogo de archivos EDI -->	
					</li>
					<!-- Modulo_Envio_EDI_Clientes -->
					
					<li id="link_cargaInfo"><a href="#" onclick="openBrowser('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="[html]menu.cargaInformacion"> &raquo;Carga de informaci&oacute;n</a></li>
					
					<li><a href="#" onClick="openNoSpots('<%=trim(user)%>','<%=trim(Nombre)%>','ADMIN');" data-i18n="menu.cargaNoSpots">&raquo;Carga de NO SPOTS</a></li>
					
					<li class="has-sub"><a href="#" data-i18n="[html]menu.administrarCuenta.titulo"> &raquo;Administre su cuenta</a>
						<ul>
							<li><a href="RegistroUsuarios.asp?ln=<%=lg%>" data-i18n="[html]menu.administrarCuenta.usuarios"> &raquo;Administre sus usuarios </a><br /></li>
							<!-- Sistema contraseñas viejo -->	
							<li><a href="CambioPassword.asp?ln=<%=lg%>" data-i18n="menu.administrarCuenta.contrasena"> &raquo;Cambiar contraseña</a></li>
							<!-- Sistema nuevo de contraseñas>>> Quitar comentario y comentar línea de arriba -->	
							<!--<li><a href="AplicacionPaebsa/ReestablecerContrasena.aspx?tipoUsr=M&pagina=loginPaebsa.asp?ln=<%=lg%>" data-i18n="[html]menu.administrarCuenta.contrasena">  &raquo;Cambiar contraseña</a><br /></li>-->
							<li><a href="#" id="modal"  onclick="ventanaHistorial();" data-i18n="[html]menu.administrarCuenta.historial">  &raquo;Historial de usuarios</a></li>
							<!-- Historial de usuarios -->	
						</ul>
					</li>
					<li><a href="loginPaebsa.asp?ln=<%=lg%>" data-i18n="[html]menu.general"> &raquo;Consulta general </a><br /></li>
					<li><a onclick="most()" href="InfoReceivedSupplier.asp?ln=<%=lg%>" data-i18n="[html]menu.enviada"> &raquo;Informaci&oacuten enviada a clientes</a></li>
					
					<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
					<li id="link_desadv">
						<a href="#" onclick="openTemplate('<%=trim(user)%>','ADMIN')" data-i18n="[html]menu.template">&raquo;Captura de confirmaci&oacute;n para los templates de Walmart/Sahuayo (DESAV)</a>
					</li>		
					 	<!-- Captura de confirmación para los templates de Walmart(DESAV) -->	
						
                    <!-- Menu Colgate -->
					<li class="has-sub" style="display:<% if trim(user)="CPA7503043P1" then response.write "block" else response.write "none" end if %>"><a title="" href="#" data-i18n="[html]menu.colgate.titulo">&raquo;Usuarios Colgate</a>
						<ul>
							<!-- Inicia Link SemiEdi-->	
							<%
											
                                Call semiEDI(trim(user),trim(pass),trim(Nombre),"loginPaebsa.asp?ln="&lg)
							%>
							<!-- Termina Link SemiEdi-->						 
						 </ul>
					</li>
		            <!-- Menu Colgate -->
                    <!-- Link de SemiEdiColgate-->
					<!-- Link de Facturas express -->		
					<li>
					<%
						Call facturaExpress(pass,user,Nombre)
					%>	
					</li>
                    <!-- Link de facturas express -->	
							
					<!-- Modulo_Nestle_Casa_Ley -->
					<li>
					<%
						Call Modulo_Nestle_Casa_Ley(user, pass, Nombre, lg)
					%>	
					</li>	
					<!-- Modulo_Nestle_Casa_Ley -->
	
					<!-- Modulo_ARS_Nestle -->
					<li>
					<%
						Call Modulo_ARS_Nestle(user, pass, Nombre, lg,"SPOKE")
					%>	
					</li>	
					<!-- Modulo_ARS_Nestle -->
	
	
	
			   
					<li><!-- Link de Reportes log -->
					<%
						Call reporteBitacoras(pass,user,Nombre)
					%>	
					</li>
							<!-- Link de Reportes log -->
						<!-- Link de Reportes Excel -->
						<li>
						<%
							Call reporteExcel(pass,user,Nombre) 
						%>	
						</li>
						<!-- Link de Reportes Excel -->
						<!-- Inicia Link de generacion de archivo ASN ALMGARCIA --> 
						<li  style="display:<% if trim(user)="CIVSA" or trim(user)="MXG1505" or (trim(user) = "MXG1397" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXGU435" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) or (trim(user) = "MXG2004" and CDate("2023-01-01 00:00:00") >= CDate("2022-01-24 00:00:00")) then response.write "block" else response.write "none" end if %>">
							<%
									Call ASNAlmGarcia(trim(pass),trim(user),trim(Nombre),"loginPaebsa.asp?ln="&lg)
							%>
						</li>
						<!-- Termina Link de generacion de archivo ASN ALMGARCIA -->
						<!-- Inicia Link de administrar brokers  ALMGARCIA -->
						<li  style="display:<% if trim(user)="CIVSA" then response.write "block" else response.write "none" end if %>">
							<%
									Call Brokers(trim(pass), trim(user),trim(Nombre))
							%>
						</li>
						<li>
							<%
								Call CargaFragua(pass,user,Nombre)		
							%>
						</li>
						<li>
							<%
								Call CargaProductosExcel(rtrim(Nombre),"loginPaebsa.asp?ln="&lg, rtrim(user), rtrim(tipoUser), pass)		
							%>
						</li>

						<li>
							<%
								Call CargaImssExcel(rtrim(user), rtrim(pass), rtrim(Nombre), "", "loginPaebsa.asp?ln="&lg)		
							%>
						</li>
						<li>
							<%
								Call CargaTiendas(rtrim(user), rtrim(Nombre), "ADMIN", "loginPaebsa.asp?ln="&lg)		
							%>
						</li>

							<!-- Termina Link de administrar brokers  ALMGARCIA -->	
						<li class="has-sub">
						<a href="#" data-i18n="menu.addendas">&raquo; Captura de Adendas</a>
							<ul>
							
							<!-- Link de Facturas Walmart -->
							<li class="has-sub"><a href="#" data-i18n="menu.adendaWalmart">&raquo;Addendas de Wal-Mart</a>
							    <ul>
								    <li>
								    <%
								        Call AddendaWalmartEdi(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
								    %>
								    </li>
								    <li>
								    <%
								        Call addendaWalmartResguardo(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
								    %>
								    </li>
							    </ul>
							</li>
							<!-- Link de facturas Walmart -->							
							<!-- Link de envio de facturas con addenda de amazon-->
							<li>
							<%							
								Call addendaAmazon(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
							%>	
							</li>
							<!-- Fin link -->
							<!-- Link de envio de facturas con addenda de BB&B-->
							<li>
							<%							
								Call  addendaEdiBBB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg) 
							%>	
							</li>							
							<!-- Fin link -->
							<!-- Link de envio de facturas con addenda de almacenes Garcia-->
							<li>
							<%							
								Call addendaAlmacenesGarcia(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)
							%>	
							</li>							
							<!-- Fin link -->
							
							<!-- Inicia Addenda de MERZA -->
							<li>
							<%
								Call addendaMerza(pass,user,Nombre,"loginPaebsa.asp?ln="&lg)		
							%>
							</li>
                                <!-- Termina Addenda de Merza -->
                            <!-- Inicia Addenda de Corvi -->
							<li>
							<%
								Call addendaCorvi(pass,user,"", "loginPaebsa.asp?ln="&lg)		
							%>
							</li>
							<li>
							<%
								Call addendaChedraui(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
							%>
							</li>
							<li>
							<%
								Call AddendaHEB(pass,user,Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")		
							%>
							</li>
						    <!--Inicia Addenda SORIANA-->
						    <!--.functions.asp-->
							<li class="has-sub"><a href="#" >&raquo;Addendas SORIANA</a>
							    <ul>
								    <li>
								    <%
								       
								    %>
								    </li>
								    <li>
								    <%
								        Call Addenda_SORIANA_Ext(pass,user,"",Nombre,"loginPaebsa.asp?ln="&lg,"ADMIN")
								    %>
								    </li>
							    </ul>
							</li>        
		                    <!--Termina Addenda SORIANA-->




						</ul>
				</li>
		</ul>		
	</div>
	<!-- inicia script de acordeon-->		
	<script type="text/javascript">
	$('#ListaMenu').hide();
	$('#ListaMenu').removeClass('activo');
	$('#TituloMenu').click(function()
	{
		var c = $("#ListaMenu");
		var mostrandose = c.css("display");
		if (mostrandose=="block"){
			$("#ListaMenu").slideUp()
		}else{
			$("#ListaMenu").slideDown("slow");
		}
	});
	</script>						
	<!-- termina script de acordeon -->	
							

	</div>
</div>
 
	<div class="container_12 divider">
		<div class="info_options">
		<h2 class="titulo-grid" data-i18n="funcionalidad.titulo"> Informaci&oacute;n  Recibida </h2>
		<!--<strong><label style="font-size: 10pt;color:#000; "><< Informaci&oacuten Enviada >> </label></strong>-->
		<form action="ficheroExcel.php" method="post"  id="FormularioExportacion">
			<div class="izq" style="padding-left:30px">
				<div class="input"><input type="button"  class="button_opt prtText" value="Reprocesar archivos" onclick="reprocesoarchivos(this,<%For i = 0 to ubound(matriz) 
									Response.Write matriz(i) 
									next%>)" id="btnReprocesoEdi" data-i18n="[value]funcionalidad.reproceso"/></div>
									
				<a class="tooltip" title="[!]Importante[/!]Para el reproceso de archivos solo se tomaran los primeros 20 registros seleccionados ademas que deberan de estar en formato EDI." data-i18n="[title]funcionalidad.infoReproceso"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
				<div class="input"><input type="button" class="button_opt prtText" value="Reprocesar PDF" onclick="generarPDFs(this,<%For i = 0 to ubound(matriz) 
									Response.Write matriz(i) 
									next%>)" id="btnReprocesoPDF" data-i18n="[value]funcionalidad.reprocesoPDF" /></div>
				<a class="tooltip" title="[!]Importante[/!]Para la generación de PDF solo se tomaran los primeros 20 registros seleccionados ademas que deberan de estar en formato EDI." data-i18n="[title]funcionalidad.infoReprocesoPDF"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
			</div>
			<div class="der" style="padding-right: 30px;">
			<div class="input"><input  class="button_opt prtText create-user btn_download" type="button" value="Descarga masiva de archivos" id="btnDescargaM"/></div>
			<a class="tooltip" title="" data-i18n="[title]funcionalidad.descargaMasiva"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
			<div class="input"><input  class="button_opt prtText" type="button" value="Enviar informaci&oacute;n por e-mail" onclick="marcarb('S')" id="btnEmail" data-i18n="[value]funcionalidad.email"/></div>
			<a class="tooltip" title="[!]Importante[/!]Para el Envio de mail solo se adjuntaran los primeros 20 registros seleccionados " data-i18n="[title]funcionalidad.infoEmail"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a>
			<div class="input"><input  class="button_opt prtText" type="button"  value="Exportar datos a un excel" onclick="descargaExcel()" id="btnExcel" data-i18n="[value]funcionalidad.excel"/></div>
			<a class="tooltip" title="[!]Importante[/!]Se exporta todo el resultado de la consulta" data-i18n="[title]funcionalidad.infoExcel"><img src="imagenes2/infoAd.jpg" width="15" height="15" alt="info" longdesc="Descripcion de Nombre" /></a> 
			</div>
		</form>
	</div>
	
		<div class="grid_8 height250" id="gridData">
			<table cellpadding="0" cellspacing="0" border="0" id="myTable02" class="tinytable">
				<thead>
					<tr>
						<th class="nosort"></th>
						<th class="sizeTittle"><input id="cTodos" name="checkbox" type="checkbox" onClick="marcar(this,<%For i = 0 to ubound(matriz) 
						Response.Write matriz(i) 
						next%>)"/></th>
						<th><h3 data-i18n="grid.nombre">Nombre cadena</h3></th>
						<th><h3 data-i18n="grid.noProveedor">No. de proveedor cadena</h3></th>
						<th><h3 data-i18n="grid.transaccion">C&oacute;digo de transacci&oacute;n </h3></th>
						<th><h3 data-i18n="grid.noDocumento">No. de documento</h3></th>
						<th><h3 data-i18n="grid.fechaHora">Fecha y hora de consulta </h3></th>
                        <th><h3 data-i18n="grid.fechaCancelacion">Fecha cancelaci&oacute;n documento </h3></th>
						<th><h3 data-i18n="grid.fechaDocumento">Fecha documento </h3></th>
						<th><h3 data-i18n="grid.claveCliente">Clave cliente </h3></th>
						<th><h3 data-i18n="grid.noControl">No. de control</h3></th>
						<th><h3 data-i18n="grid.estado">Estado</h3></th>
                        <th><h3 data-i18n="grid.codigoTienda">C&oacute;digo tienda </h3></th>
                        <th><h3 data-i18n="grid.descripcion">Descripci&oacute;n del proceso </h3></th>
                        <th><h3 data-i18n="grid.descargar" class="sizeTittle40">Descargar</h3></th>
					</tr>
				</thead>
				<% 
					PeriodoDepMeses=Info_Dias
					
					if isNull (PeriodoDepMeses) or PeriodoDepMeses=0  or  PeriodoDepMeses="" or PeriodoDepMeses<=3 then
							PeriodoDepMeses=90
 
					else 
							PeriodoDepMeses=PeriodoDepMeses*30
					end if
					
					fechaactual = date()
					'se convierte la fecha actual a dd/mm/yyyy por que el servidor regresa el formato ingles mm/dd/yyyy
					fechaactual=Month(date())&"/"&Day(date())&"/"&Year(date())
					diasRestantes=PeriodoDepMeses-5
					
					do while not rs.eof and contador <= cint(tamanopagina) 
						fila ="fila"&contador
						id="c"&contador
						valorFcc=trim(rs("Fecha_Consulta_Cliente")&"")
						valorHcc=trim(rs("Hora_Consulta_Cliente")&"")
						valorFEnvio=trim(rs("Fecha_Envio_Proveedor")&"")
						envio=""&valorFEnvio&""
						'fecha validacion
						ayer = ""&valorFcc&""
						if ayer="" then
						ayer ="20220101"
						end if

						ayeranio = mid(ayer,1,4)
						ayermes = mid(ayer,5,2)
						ayerdia = mid(ayer,7,2)
						
						fechaAnio=mid(envio,1,4)
						fechaMes=mid(envio,5,2)
						fechaDia=mid(envio,7,2)
						
						fechaEnvio=fechaMes&"/"&fechaDia&"/"&fechaAnio
						ayerFecha=ayermes&"/"&ayerdia&"/"&ayeranio
						
						diasonline = DateDiff("d", ayerFecha, fechaactual)
						diasNoConsultado=DateDiff("d",fechaEnvio,fechaactual)
 	
						'fecha validacion
						
						if valorFcc<>"" and  valorHcc<>"" then
							color="si"
							if valorFcc<>"" and  valorHcc<>"" and diasonline =>diasRestantes then
								color="limite"
							end if
						else
							'color="no"
							if ((isNull (valorFcc) and isNull (valorHcc)) or(valorFcc="" and valorHcc="")) and diasNoConsultado=>diasRestantes then
								color="limite"
							else
								color="no"
							end if
						end if
				%>
				<tbody>
					<tr class="<%=color%>" id="<%=fila%>" >
					<td><%= contador%></td>
					<td class="columnas-caja"><input id="<%=id%>" type="checkbox" value="<%= "ndd"&contador&"="&trim(rs("Num_control_dato_docto"))&"&"&"idf"&contador&"="&trim(rs("Identificador_Formato_1"))&"&ctr"&contador&"="&trim(rs("Codigo_Transaccion")) &"&na"&contador&"="&trim(rs("Nombre_Archivo")) %>" onClick="marcar(this,'<%=fila%>')"/></td>
					<td><%= rs("Nombre_Hub")%></td>
					<td><%= rs("Numero_Proveedor_Hub")%></td>
					<td><%= rs("Codigo_Transaccion")%></td>
					<td class="size100 wrap"><%= rs("Num_control_dato_docto")%></td>
					<td><% 
						if  (Trim(rs("Fecha_Consulta_Cliente")) = "" or isNull (rs("Fecha_Consulta_Cliente"))) AND (Trim(rs("Hora_Consulta_Cliente")) = "" or isNull (rs("Hora_Consulta_Cliente"))) then
							response.Write("-")
						else
							consultaCliente=trim(rs("Fecha_Consulta_Cliente"))
							consultaClienteFinal=formatoFechas(consultaCliente)
							horaConsultaCliente=trim(rs("Hora_Consulta_Cliente"))
							horaFinalConsultaCliente=formatoHora(horaConsultaCliente)
							response.Write(""&consultaClienteFinal&"-"&horaFinalConsultaCliente)
						end if
					
					%></td>
					<td><%response.write formatoFechas(trim(rs("Fecha_Canc_Documento_Edi")))%></td>
                    <td><%response.write formatoFechas(trim(rs("Fecha_Recepcion_Sistema")))%></td>
		            <td><%= rs("Id_Hub")%></td>
					<td class="size100 wrap"><%= rs("Num_Intercambio_Recibido")%></td>
					<td><%
							estadoArchivo=Trim(rs("Status"))
								if estadoArchivo="ERROR07" then 
									response.Write("No es proveedor") 
									else 					if estadoArchivo="ERROR11" then 
									response.Write("No es cliente PAEBSA") 
									else 					if estadoArchivo="ERROR13" then 
									response.Write("Desconectado") 
									else 					if estadoArchivo="ERROR14" then 
									response.Write("Duplicado en transmisión") 
									else 					if estadoArchivo="ERROR15" then 
									response.Write("Enviado anteriormente") 						
									else
									response.Write(""&estadoArchivo&"")
								end if
							end if 
							end if
						end if
						end if
						%>
					</td>
					<td class="size80 wrap">
                        <%= rs("Codigo_Tienda")%>
					</td>
                    <td class="size overflow">
                        <%=Trim(rs("Descripcion_Error")) %>
                    </td>
					<td>
					<%
					' Creacion de la lista de archivos para su descarga 
                    transaccion=trim(rs("Codigo_Transaccion"))
					idHub=trim(rs("Id_Hub"))
					
					If isNull (rs("Nombre_Archivo")) or Trim(rs("Nombre_Archivo")) = "" Then
						response.Write(" <img src=imagenes2/error3.png alt=PAEBSA/>")			
					else
						 
						NombreArchivo= (rs.fields ("Nombre_Archivo")  & " ")
						
						Id_Cliente= (rs.fields ("Id_Cliente")  & " ")
						cliente = trim(Id_Cliente)
						archivo = rtrim(NombreArchivo)
 
						NombreArchivoPdf=rtrim(rs.fields("Nombre_Archivo_PDF"))
						NombreArchivoExcel=rtrim(rs.fields("Nombre_Archivo_Excel"))
						NombreArchivoCsv=rtrim(rs.fields("Nombre_Archivo_CSV"))
						NombreArchivoTxt=rtrim(rs.fields("Nombre_Archivo_Txt"))
						NombreArchivoEtq=rtrim(rs.fields("Nombre_Archivo_Etiquetas"))
						NombreArchivoXml=rtrim(rs.fields("Nombre_Archivo_XML"))
                        NombreArchivoLog=rtrim(rs.fields("Nombre_Archivo_Log"))
 
						NombreArchivoPdf=iif(NombreArchivoPdf,NombreArchivoPdf,"N-PDF")
						NombreArchivoExcel=iif(NombreArchivoExcel,NombreArchivoExcel,"N-XLS")
						NombreArchivoCsv=iif(NombreArchivoCsv,NombreArchivoCsv,"N-CSV")
						NombreArchivoTxt=iif(NombreArchivoTxt,NombreArchivoTxt,"N-TXT")
						NombreArchivoEtq=iif(NombreArchivoEtq,NombreArchivoEtq,"N-Etq")
						NombreArchivoXml=iif(NombreArchivoXml,NombreArchivoXml,"N-XML")
                        NombreArchivoLog=iif(NombreArchivoLog,NombreArchivoLog,"N-LOG")
 
					   Set dataFiles = Server.CreateObject("Scripting.Dictionary")
						   dataFiles.Add "EDI",archivo
						   dataFiles.Add "PDF",NombreArchivoPdf
						   dataFiles.Add "XLS",NombreArchivoExcel
						   dataFiles.Add "TXT",NombreArchivoTxt
						   dataFiles.Add "XML",NombreArchivoXml
							
                        If transaccion="CONTRL" or transaccion="APERAK" or transaccion="APECOM" or transaccion="APEFIS" or transaccion="864" then
                            dataFiles.Add "LOG",NombreArchivoLog
                            Call dictionaryArchive(cliente,idHub,user,dataFiles,contador)
                        Else
                            dataFiles.Add "CSV",NombreArchivoCsv
                            dataFiles.Add "ETQ",NombreArchivoEtq
                            Call dictionaryArchive(cliente,idHub,user,dataFiles,contador)
						End If
					End If 
					' Creacion de la lista de archivos para su descarga 
					%>
					</td>
					</tr>
					<%
						rs.movenext
						contador = contador +1
						loop
					%> 
				</tbody>
			</table>	
		</div>
 
        <div id="tablefooter">
				<div id="tablenav">
						<div>
                            
							<%
                            texto= value(texto)
                            texto2= value(texto2)
                            j=0
							if cint(paginaabsoluta) <> 1 then
								response.write "<td align=center><a href=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta=" & atras & "><img src=imagenes2/first.png width=16 height=16 style=border:0 alt=First Page /></a></td>"
							    j=j+1
							end if
							%>
							<%j=0
							if cint(paginaabsoluta) <> 1 then
								atras=cint(paginaabsoluta)-1	
								response.write "<td align=center><a href=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta=" & atras & "><img src=imagenes2/previous.png width=16 height=16 style=border:0 alt=Previous Page  /></a></td>"
							    j=j+1
							end if
							%>  
							<%if cint(paginaabsoluta) <> rs.pagecount then
								atras=cint(paginaabsoluta)+1
                                response.write "<td align=center><a href=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta="&atras&"><img src=imagenes2/next.png width=16 height=16 style=border:0 alt=Next Page  /></a></td>"
							end if%>
							<%j=0
							if cint(paginaabsoluta) <> rs.pagecount then
								atras=rs.pagecount
								response.write"<td align=center><a href=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta="&atras&"><img src=imagenes2/last.png width=16 height=16 style=border:0 alt=Last Page /></a></td>"
							    j=j+1
							end if
							%>       
							<label><span class="page" data-i18n="grid.pagina">P&aacute;gina -</span></label>
							<%
							response.write "<form name=frmDireccionesASP1 id=frmDireccionesASP1 action=loginPaebsa.asp>"
							response.write "<select class='select-text select-opt' style='width: 100%' size=1 name=listaDireccionesASP1 onchange=window.top.location.href=frmDireccionesASP1.listaDireccionesASP1.options[frmDireccionesASP1.listaDireccionesASP1.selectedIndex].value >"	
							for i = 1 to rs.pagecount
								j=j+1
								if cint(i) = cint(paginaabsoluta) then
									response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta="& i &" selected="&paginaabsoluta&">"&i&"</option>"
									
								else
									response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina="&tamanopagina&"&paginaabsoluta="& i &">"&i&"</option>"
									
								end if
							next
							end if
							response.write "</select>"
							response.write "</form>"
							%>
						</div>
				</div>
			<div id="tablelocation">
            	<div>
					<%
						response.write "<form name=frmDireccionesASP id=frmDireccionesASP action=loginPaebsa.asp>"
						response.write "<select class='select-text select-opt' style='width: 100%;' size=1 name=listaDireccionesASP onchange=window.top.location.href=frmDireccionesASP.listaDireccionesASP.options[frmDireccionesASP.listaDireccionesASP.selectedIndex].value >"
						response.write "<option  selected=selected data-i18n='grid.seleccionar'> Seleccione</option>"
						response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina=25&paginaabsoluta=1  >25</option>"
						response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina=50&paginaabsoluta=1    >50</option>"
						response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina=75&paginaabsoluta=1    >75</option>"
						response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina=100&paginaabsoluta=1    >100</option>"
						response.write "<option value=loginPaebsa.asp?ln="&lg&"&seleccione="&seleccione&"&texto="&texto&"&seleccione2="&seleccione2&"&texto2="&texto2&"&alf="&alf&"&orden="&orden&"&tipofecha="&tipofecha&"&datepicker="&fechaini&"&datepickerfinal="&fechafin&"&tamanopagina=200&paginaabsoluta=1   >200</option>"		 
						response.write "</select>"
						response.write "</form>"
					%>   
                </div>
                <div class="page"> <span class="page" data-i18n ="grid.entradas">Entradas por p&aacute;gina </span><span data-i18n="grid.pag">P&aacute;gina</span> <span id="currentpage"><%= paginaabsoluta %></span><span data-i18n="grid.de"> de</span><span id="totalpages"><%= rs.PageCount %></span></div>
            </div>
        </div>
		<div class="clear"></div>
	</div>
	<%
		Response.write("<div id=dialog-form title=Seleccione su archivo data-i18n='[title]dialogo.archivo'>")
		Response.write("<div id=links></div>")
		Response.write("</div>")
	%>
	<%
        Response.Write("<div id='contentPDF'></div>")
	%>
    <div id="dialog-confirm" title="Aviso"><div id="content_msg"></div></div>  
</div>
<!--  Termina general -->
<%
	end if
%>        
<%
	end if
%>
<%
	rs.Close
	Set rs = Nothing
	cnn.Close
	Set cnn = Nothing
%>
<%
	end if
		On Error Goto 0	
%>
<script src="js/app.js"></script>
<script>
	function descargaMasiva(fActual, hActual, tipoUsr, idC, idH, arr, tokenUsr)
    {
		var tokenNuevo="";
        //console.log('Exito');
		//console.log('Parámetros\nFecha actual: ' + fActual + '\nHoraActual: ' + hActual + '\nTipo usuario: ' + tipoUsr + '\nId de Cliente: ' + idC + '\nArreglo: ' + arr + '\nToken: ' + tokenUsr);
		if(tipoUsr==undefined||tipoUsr==""||tipoUsr=="undefined")
			tipoUsr = "Spokes";
		$.ajax
		(
			{
				type: "POST",
                url: "AplicacionPaebsa/CompressFiles.asmx/EjecutarProcesoZip",
                data: "{folder: '" + tipoUsr + "', inbox: '" + idC + "', listaArchivos: '" + arr + "', idHub: '', banderaCliente: 'S'}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (data, status) 
				{
					var respuesta = data.d.split('|');
                    alert(respuesta[0]);
					//var ruta = respuesta[respuesta.length-2];
					var archivo = respuesta[respuesta.length-1];
					if(respuesta[0]=="GENERATED_ZIP_FILE")					
						window.location.href = "AplicacionPaebsa/descargarArchivo.aspx?archivo=" + archivo + "&idC=" + idC + "&tipoUsr=" + tipoUsr;
					
                },
                failure: function (xhr, status, error) 
				{
            		console.log("Error");
					console.log(xhr);
				}
			}
		);
		
		
    }

</script>
<script>
	$(document).ready(function(e)
	{
		$(document).on('click', '#btnCerrarSesion', function(e)
		{
			e.preventDefault();
			$.ajax
			(
				{
					type: "POST",
					url: "AplicacionPaebsa/Procesos.asmx/cerrarSesionMaestro",
					data: "{idCliente: '" + '<%=rtrim(session("usuario"))%>' + "', tipoUsr: 'S'}",
					contentType: "application/json; charset=utf-8",
					async: true,
					dataType: "json",
					success: function (data, status) 
					{
						var respuesta = data.d;
						console.log(respuesta);
						location.href = "Cerrar_Ses_Cli.asp"
						
					},
					failure: function (xhr, status, error) 
					{
						console.log("Error");
						console.log(xhr);
					}
				}
			);
		});

		$(document).on('click', '#btnSalir', function(e)
		{
			e.preventDefault();
			$.ajax
			(
				{
					type: "POST",
					url: "AplicacionPaebsa/Procesos.asmx/cerrarSesionMaestro",
					data: "{idCliente: '" + '<%=rtrim(session("usuario"))%>' + "', tipoUsr: 'S'}",
					contentType: "application/json; charset=utf-8",
					async: true,
					dataType: "json",
					success: function (data, status) 
					{
						var respuesta = data.d;
						console.log(respuesta);
						location.href = "Cerrar_Ses_Cli.asp"
						
					},
					failure: function (xhr, status, error) 
					{
						console.log("Error");
						console.log(xhr);
					}
				}
			);
		});
	});
</script>
</body><!--TERMINA BODY -->
</html>