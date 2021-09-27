<%@ page language="java" contentType="text/html; charset=ISO-8859-1" pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html style="width: 100%; height:100%; margin: 0; padding: 0; border: 0; background-color:gold;">
<head>
	<meta charset="ISO-8859-1">
	<title>Compilación de exceles</title>
	<base href="${pageContext.request.contextPath}/">
	
	<script>
	<% String mensaje = (String)request.getAttribute("mensaje"); %>
	<% if (mensaje != null) { %>
		alert("<%=mensaje%>")
	<% }%> 
	</script>
</head>

<body style="width: 100%; height:100%; margin: 0px; padding: 0; border: 0">
<div style="width: 100%; height:100%; background-color: #BBBBBB;">
	<form action="servletExcel" style="margin: 0; padding: 0; border: 0;">
		<h1 style="margin: 0; padding:0;">Herramienta de compilación de exceles de se&ntilde;alizaci&oacute;n</h1>
		<br>		
		<label for="directorioEntrada">Directorio de entrada</label>
		<input type="text" name="directorioEntrada" value="C:/senalizacion_exceles/" id="directorioEntrada" required="required">
		<br><br>
		<label for="directorioSalida">Directorio de salida</label>
		<input type="text" name="directorioSalida" value="C:/senalizacion_exceles/compilados/" required="required" style="width:230px;">
		<br><br>
		<button>Compilar</button>
	</form>
	</div>
</body>
</html>