Attribute VB_Name = "informereformapac"
Sub Informe_de_Reforma_PAC()
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
    
    ' Variables para datos de la hoja "SECUENCIAS"
    Dim unidadRequirente As String, unidadRequirente1 As String
    Dim antecedenteReforma As String, necesidadReforma As String
    Dim justificacionTecnica As String
    Dim nroIJN As String, fechaIJN As String
    Dim nombreTecnicoUnidad As String, cargoTecnicoUnidad As String
    Dim objetoDeContratacion As String, entidad As String
    Dim certificacionPresupuestaria As String, fechaCertificacion As String
    Dim partida1 As String, denominacion1 As String
    Dim partida2 As String, denominacion2 As String
    Dim objetoDeContratacion1 As String, partidaPAC As String
    Dim presupuestoReferencial As String, valorLetras As String
    Dim objetoDeContratacion2 As String, cpcDelProceso As String
    Dim partidaPAC1 As String, codigoCPCProceso As String
    Dim tipoDeCompra As String, TipoRegimen As String
    Dim TipoProducto As String, tipoDeProceso As String
    Dim objetoDeContratacion3 As String, presupuestoReferencial1 As String
    Dim cuatrimestre As String
    Dim nombreTitularUnidad As String, cargoTitularUnidad As String
    Dim directorAdministrativo As String, cargoAdministrativo As String
    Dim fechaElaboracion As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave para la estructura del libro y otras hojas
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D146
    plantillaID = wsBase.range("D146").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D146 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "DocumentoTerminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
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
    unidadRequirente = ws.range("D2").Value
    unidadRequirente1 = ws.range("D2").Value
    antecedenteReforma = ws.range("BZ2").Value
    necesidadReforma = ws.range("CA2").Value
    justificacionTecnica = ws.range("CB2").Value
    nroIJN = ws.range("X2").Value
    fechaIJN = ws.range("Y2").Value
    nombreTecnicoUnidad = ws.range("G2").Value
    cargoTecnicoUnidad = ws.range("H2").Value
    objetoDeContratacion = ws.range("Q2").Value
    entidad = ws.range("A2").Value
    certificacionPresupuestaria = ws.range("BN2").Value
    fechaCertificacion = ws.range("BO2").Value
    partida1 = ws.range("BP2").Value
    denominacion1 = ws.range("BQ2").Value
    partida2 = ws.range("BR2").Value
    denominacion2 = ws.range("BS2").Value
    objetoDeContratacion1 = ws.range("Q2").Value
    partidaPAC = ws.range("BU2").Value
    presupuestoReferencial = ws.range("BV2").Value
    valorLetras = ws.range("BW2").Value
    objetoDeContratacion2 = ws.range("Q2").Value
    cpcDelProceso = ws.range("BA2").Value
    partidaPAC1 = ws.range("BU2").Value
    codigoCPCProceso = ws.range("BT2").Value
    tipoDeCompra = ws.range("O2").Value
    TipoRegimen = ws.range("BX2").Value
    TipoProducto = ws.range("P2").Value
    tipoDeProceso = ws.range("S2").Value
    objetoDeContratacion3 = ws.range("Q2").Value
    presupuestoReferencial1 = ws.range("BV2").Value
    cuatrimestre = ws.range("BY2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    directorAdministrativo = ws.range("K2").Value
    cargoAdministrativo = ws.range("L2").Value
    fechaElaboracion = ws.range("GZ2").Value

    ' Proteger y ocultar la hoja "SECUENCIAS" nuevamente
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla en la ruta temporal
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_ReformaPAC_Temp.docx"

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send
    If objHTTP.status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binario
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
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Unidad_Requirente1") Then .Bookmarks("Unidad_Requirente1").range.Text = unidadRequirente1
        If .Bookmarks.Exists("Antecedente_Reforma") Then .Bookmarks("Antecedente_Reforma").range.Text = antecedenteReforma
        If .Bookmarks.Exists("Necesidad_Reforma") Then .Bookmarks("Necesidad_Reforma").range.Text = necesidadReforma
        If .Bookmarks.Exists("Justificacion_Tecnica") Then .Bookmarks("Justificacion_Tecnica").range.Text = justificacionTecnica
        If .Bookmarks.Exists("Nro_IJN") Then .Bookmarks("Nro_IJN").range.Text = nroIJN
        If .Bookmarks.Exists("Fecha_IJN") Then .Bookmarks("Fecha_IJN").range.Text = fechaIJN
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad") Then .Bookmarks("Nombre_Tecnico_Unidad").range.Text = nombreTecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad") Then .Bookmarks("Cargo_Tecnico_Unidad").range.Text = cargoTecnicoUnidad
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Certificacion_Presupuestaria") Then .Bookmarks("Certificacion_Presupuestaria").range.Text = certificacionPresupuestaria
        If .Bookmarks.Exists("Fecha_Certificacion") Then .Bookmarks("Fecha_Certificacion").range.Text = fechaCertificacion
        If .Bookmarks.Exists("Partida1") Then .Bookmarks("Partida1").range.Text = partida1
        If .Bookmarks.Exists("Denominacion1") Then .Bookmarks("Denominacion1").range.Text = denominacion1
        If .Bookmarks.Exists("Partida2") Then .Bookmarks("Partida2").range.Text = partida2
        If .Bookmarks.Exists("Denominacion2") Then .Bookmarks("Denominacion2").range.Text = denominacion2
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Partida_PAC") Then .Bookmarks("Partida_PAC").range.Text = partidaPAC
        If .Bookmarks.Exists("Presupuesto_Referencial") Then .Bookmarks("Presupuesto_Referencial").range.Text = presupuestoReferencial
        If .Bookmarks.Exists("Valor_Letras") Then .Bookmarks("Valor_Letras").range.Text = valorLetras
        If .Bookmarks.Exists("Objeto_de_Contratacion2") Then .Bookmarks("Objeto_de_Contratacion2").range.Text = objetoDeContratacion2
        If .Bookmarks.Exists("CPC_del_proceso") Then .Bookmarks("CPC_del_proceso").range.Text = cpcDelProceso
        If .Bookmarks.Exists("Partida_PAC1") Then .Bookmarks("Partida_PAC1").range.Text = partidaPAC1
        If .Bookmarks.Exists("Codigo_CPC_Proceso") Then .Bookmarks("Codigo_CPC_Proceso").range.Text = codigoCPCProceso
        If .Bookmarks.Exists("Tipo_de_Compra") Then .Bookmarks("Tipo_de_Compra").range.Text = tipoDeCompra
        If .Bookmarks.Exists("Tipo_Regimen") Then .Bookmarks("Tipo_Regimen").range.Text = TipoRegimen
        If .Bookmarks.Exists("Tipo_Producto") Then .Bookmarks("Tipo_Producto").range.Text = TipoProducto
        If .Bookmarks.Exists("Tipo_de_Proceso") Then .Bookmarks("Tipo_de_Proceso").range.Text = tipoDeProceso
        If .Bookmarks.Exists("Objeto_de_Contratacion3") Then .Bookmarks("Objeto_de_Contratacion3").range.Text = objetoDeContratacion3
        If .Bookmarks.Exists("Presupuesto_Referencial1") Then .Bookmarks("Presupuesto_Referencial1").range.Text = presupuestoReferencial1
        If .Bookmarks.Exists("Cuatrimestre") Then .Bookmarks("Cuatrimestre").range.Text = cuatrimestre
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Director_Administrativo") Then .Bookmarks("Director_Administrativo").range.Text = directorAdministrativo
        If .Bookmarks.Exists("Cargo_Administrativo") Then .Bookmarks("Cargo_Administrativo").range.Text = cargoAdministrativo
        If .Bookmarks.Exists("Fecha_elaboracion") Then .Bookmarks("Fecha_elaboracion").range.Text = fechaElaboracion
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

    ' Eliminar el archivo temporal descargado (por si se requiriera)
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


