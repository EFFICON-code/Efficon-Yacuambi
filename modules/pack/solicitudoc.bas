Attribute VB_Name = "solicitudoc"
Sub Solicitud_de_Orden_de_Compra()
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object

    Dim objHTTP As Object
    Dim objStream As Object
    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant
    
    Dim ws As Worksheet
    Dim wsBase As Worksheet

    Dim CLAVE As String
    Dim siglas As String
    Dim periodo As String
    Dim lugar As String
    Dim fecha As String
    Dim administrativo As String
    Dim cargoAdministrativo As String
    Dim objetoContratacion As String
    Dim disposicionPublicacion As String
    Dim fechaDisposicionPublicacion As String
    Dim fechaPublicacion As String
    Dim codigoNecesidad As String
    Dim nroCertificacionPresupuestaria As String
    Dim fechaCertificacionPresupuestaria As String
    Dim entidad As String
    Dim presupuesto As String
    Dim valorLetras As String
    Dim nroInforme As String
    Dim proveedor As String
    Dim ruc As String
    Dim comprasPublicas As String
    Dim cargoComprasPublicas As String

    ' Claves para desproteger
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B146
    plantillaID = wsBase.range("B146").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B146 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (p.e., Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "Solicitud_Orden_Compra_Terminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    ' Asignar la hoja "SECUENCIAS" y desprotegerla
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=claveSecuencias

    ' Leer datos de Excel
    siglas = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    administrativo = CStr(ws.range("K2").Value)
    cargoAdministrativo = CStr(ws.range("L2").Value)
    objetoContratacion = CStr(ws.range("Q2").Value)
    disposicionPublicacion = CStr(ws.range("DM2").Value)
    fechaDisposicionPublicacion = CStr(ws.range("DN2").Value)
    fechaPublicacion = CStr(ws.range("DQ2").Value)
    codigoNecesidad = CStr(ws.range("DP2").Value)
    nroCertificacionPresupuestaria = CStr(ws.range("DV2").Value)
    fechaCertificacionPresupuestaria = CStr(ws.range("DW2").Value)
    entidad = CStr(ws.range("A2").Value)
    presupuesto = CStr(ws.range("DC2").Value)
    valorLetras = CStr(ws.range("DD2").Value)
    nroInforme = CStr(ws.range("EJ2").Value)
    proveedor = CStr(ws.range("DE2").Value)
    ruc = CStr(ws.range("DF2").Value)
    comprasPublicas = CStr(ws.range("I2").Value)
    cargoComprasPublicas = CStr(ws.range("J2").Value)

    ' Proteger y ocultar la hoja nuevamente permitiendo modificar escenarios
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_SolicitudOrdenCompra_Temp.docx"
    
    ' Descargar la plantilla usando MSXML2.ServerXMLHTTP
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' binario
        objStream.Open
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile rutaDescargaTemporal, 2 ' Sobrescribe si existe
        objStream.Close
    Else
        MsgBox "Error al descargar la plantilla. Verifique la conexión o el enlace." & vbCrLf & _
               "Código de estado: " & objHTTP.status & " - " & objHTTP.statusText, vbExclamation
        Exit Sub
    End If

    ' Iniciar Word y abrir la plantilla descargada
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject(Class:="Word.Application")
    End If
    On Error GoTo 0

    If wdApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word.", vbCritical
        Exit Sub
    End If

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(rutaDescargaTemporal)

    If wdDoc Is Nothing Then
        MsgBox "No se pudo abrir el documento de Word.", vbCritical
        wdApp.Quit
        Exit Sub
    End If

    ' Insertar datos en los marcadores de la plantilla
    With wdDoc
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Administrativo") Then .Bookmarks("Administrativo").range.Text = administrativo
        If .Bookmarks.Exists("Cargo_Administrativo") Then .Bookmarks("Cargo_Administrativo").range.Text = cargoAdministrativo
        If .Bookmarks.Exists("Objeto_Contratacion") Then .Bookmarks("Objeto_Contratacion").range.Text = objetoContratacion
        If .Bookmarks.Exists("Disposicion_publicacion") Then .Bookmarks("Disposicion_publicacion").range.Text = disposicionPublicacion
        If .Bookmarks.Exists("Fecha_disposicion_publicacion") Then .Bookmarks("Fecha_disposicion_publicacion").range.Text = fechaDisposicionPublicacion
        If .Bookmarks.Exists("Fecha_Publicacion") Then .Bookmarks("Fecha_Publicacion").range.Text = fechaPublicacion
        If .Bookmarks.Exists("Codigo_Necesidad") Then .Bookmarks("Codigo_Necesidad").range.Text = codigoNecesidad
        If .Bookmarks.Exists("Nro_Certificacion_presupuestaria") Then .Bookmarks("Nro_Certificacion_presupuestaria").range.Text = nroCertificacionPresupuestaria
        If .Bookmarks.Exists("Fecha_Certificacion_presupuestaria") Then .Bookmarks("Fecha_Certificacion_presupuestaria").range.Text = fechaCertificacionPresupuestaria
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Presupuesto") Then .Bookmarks("Presupuesto").range.Text = presupuesto
        If .Bookmarks.Exists("Valor_letras") Then .Bookmarks("Valor_letras").range.Text = valorLetras
        If .Bookmarks.Exists("Nro_Informe") Then .Bookmarks("Nro_Informe").range.Text = nroInforme
        If .Bookmarks.Exists("Proveedor") Then .Bookmarks("Proveedor").range.Text = proveedor
        If .Bookmarks.Exists("Ruc") Then .Bookmarks("Ruc").range.Text = ruc
        If .Bookmarks.Exists("Compras_Publicas") Then .Bookmarks("Compras_Publicas").range.Text = comprasPublicas
        If .Bookmarks.Exists("Cargo_Compras_Publicas") Then .Bookmarks("Cargo_Compras_Publicas").range.Text = cargoComprasPublicas
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

    ' Eliminar archivo temporal descargado
    On Error Resume Next
    Kill rutaDescargaTemporal
    On Error GoTo 0

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsBase = Nothing

    MsgBox "El documento se ha generado correctamente.", vbInformation
End Sub


