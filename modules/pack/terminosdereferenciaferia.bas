Attribute VB_Name = "terminosdereferenciaferia"
Sub Terminos_de_Referencia_Ferias_Inclusivas()
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object

    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant

    Dim ws As Worksheet
    Dim wsProductos As Worksheet
    Dim wsPresupuesto As Worksheet
    Dim wsBase As Worksheet

    ' Claves
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Variables para datos en la plantilla
    Dim entidad As String
    Dim titulo As String
    Dim objetoDeContratacion As String
    Dim unidadRequirente As String
    Dim antecedente1 As String, antecedente2 As String, antecedente3 As String, antecedente4 As String
    Dim justificacion As String
    Dim objetivoGeneral As String
    Dim objetivosEspecificos As String
    Dim objetoDeContratacion1 As String
    Dim alcance As String
    Dim metodologiaDeTrabajo As String
    Dim informacionEntidad As String
    Dim vigenciaOferta As String
    Dim plazo As String
    Dim formaDePago As String
    Dim entidad1 As String
    Dim experienciaGeneral As String
    Dim montoGeneral As String
    Dim porContratoG As String
    Dim experienciaEspecifica As String
    Dim montoEspecifica As String
    Dim porContratoE As String
    Dim tipoEntrega As String
    Dim lugarDeEntrega As String
    Dim garantia As String
    Dim entidad2 As String
    Dim obligacionesContratista As String
    Dim marcoLegalProceso As String
    Dim nombreTecnicoUnidad As String
    Dim cargoTecnicoUnidad As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String
    Dim fechaElaboracion As String

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D135
    plantillaID = wsBase.range("D135").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D135 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", _
        "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
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

    ' Leer datos de Excel (hoja SECUENCIAS)
    entidad = ws.range("A2").Value
    titulo = ws.range("AO2").Value
    objetoDeContratacion = ws.range("Q2").Value
    unidadRequirente = ws.range("D2").Value
    antecedente1 = ws.range("Z2").Value
    antecedente2 = ws.range("AA2").Value
    antecedente3 = ws.range("AB2").Value
    antecedente4 = ws.range("AC2").Value
    justificacion = ws.range("AF2").Value
    objetivoGeneral = ws.range("AD2").Value
    objetivosEspecificos = ws.range("AE2").Value
    objetoDeContratacion1 = ws.range("Q2").Value
    alcance = ws.range("AQ2").Value
    metodologiaDeTrabajo = ws.range("AP2").Value
    informacionEntidad = ws.range("AR2").Value
    vigenciaOferta = ws.range("AU2").Value
    plazo = ws.range("T2").Value
    formaDePago = ws.range("AS2").Value
    entidad1 = ws.range("A2").Value
    experienciaGeneral = ws.range("BC2").Value
    montoGeneral = ws.range("BD2").Value
    porContratoG = ws.range("BE2").Value
    experienciaEspecifica = ws.range("BF2").Value
    montoEspecifica = ws.range("BG2").Value
    porContratoE = ws.range("BH2").Value
    tipoEntrega = ws.range("CL2").Value
    lugarDeEntrega = ws.range("AT2").Value
    garantia = ws.range("U2").Value
    entidad2 = ws.range("A2").Value
    obligacionesContratista = ws.range("BI2").Value
    marcoLegalProceso = ws.range("AL2").Value
    nombreTecnicoUnidad = ws.range("G2").Value
    cargoTecnicoUnidad = ws.range("H2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    fechaElaboracion = ws.range("GZ2").Value

    ' Proteger y ocultar la hoja SECUENCIAS
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla a una ruta temporal
    Dim objHTTP As Object, objStream As Object
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_TDRFeria_Temp.docx"

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
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Titulo") Then .Bookmarks("Titulo").range.Text = titulo
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Alcance") Then .Bookmarks("Alcance").range.Text = alcance
        If .Bookmarks.Exists("Metodologia_de_Trabajo") Then .Bookmarks("Metodologia_de_Trabajo").range.Text = metodologiaDeTrabajo
        If .Bookmarks.Exists("Informacion_Entidad") Then .Bookmarks("Informacion_Entidad").range.Text = informacionEntidad
        If .Bookmarks.Exists("Vigencia_Oferta") Then .Bookmarks("Vigencia_Oferta").range.Text = vigenciaOferta
        If .Bookmarks.Exists("Plazo") Then .Bookmarks("Plazo").range.Text = plazo
        If .Bookmarks.Exists("Forma_de_Pago") Then .Bookmarks("Forma_de_Pago").range.Text = formaDePago
        If .Bookmarks.Exists("Entidad1") Then .Bookmarks("Entidad1").range.Text = entidad1
        If .Bookmarks.Exists("Experiencia_General") Then .Bookmarks("Experiencia_General").range.Text = experienciaGeneral
        If .Bookmarks.Exists("Monto_General") Then .Bookmarks("Monto_General").range.Text = montoGeneral
        If .Bookmarks.Exists("Por_contrato_G") Then .Bookmarks("Por_contrato_G").range.Text = porContratoG
        If .Bookmarks.Exists("Experiencia_Especifica") Then .Bookmarks("Experiencia_Especifica").range.Text = experienciaEspecifica
        If .Bookmarks.Exists("Monto_Especifica") Then .Bookmarks("Monto_Especifica").range.Text = montoEspecifica
        If .Bookmarks.Exists("Por_contrato_E") Then .Bookmarks("Por_contrato_E").range.Text = porContratoE
        If .Bookmarks.Exists("Tipo_Entrega") Then .Bookmarks("Tipo_Entrega").range.Text = tipoEntrega
        If .Bookmarks.Exists("Lugar_de_Entrega") Then .Bookmarks("Lugar_de_Entrega").range.Text = lugarDeEntrega
        If .Bookmarks.Exists("Garantia") Then .Bookmarks("Garantia").range.Text = garantia
        If .Bookmarks.Exists("Entidad2") Then .Bookmarks("Entidad2").range.Text = entidad2
        If .Bookmarks.Exists("Obligaciones_Contratista") Then .Bookmarks("Obligaciones_Contratista").range.Text = obligacionesContratista
        If .Bookmarks.Exists("Marco_Legal_Proceso") Then .Bookmarks("Marco_Legal_Proceso").range.Text = marcoLegalProceso
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad") Then .Bookmarks("Nombre_Tecnico_Unidad").range.Text = nombreTecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad") Then .Bookmarks("Cargo_Tecnico_Unidad").range.Text = cargoTecnicoUnidad
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Fecha_elaboracion") Then .Bookmarks("Fecha_elaboracion").range.Text = fechaElaboracion

        ' Transferir datos de la hoja "PRODUCTOS"
        Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
        wsProductos.Unprotect password:=claveGeneral

        Dim rangoProductos As range
        On Error Resume Next
        Set rangoProductos = wsProductos.UsedRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoProductos Is Nothing Then
            rangoProductos.Copy
            If .Bookmarks.Exists("Productos") Then
                With .Bookmarks("Productos").range
                    .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                    If .Tables.Count > 0 Then .Tables(1).AutoFitBehavior wdAutoFitWindow
                End With
            Else
                MsgBox "El marcador 'Productos' no existe en la plantilla.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRODUCTOS.", vbExclamation
        End If

        wsProductos.Protect password:=claveGeneral, AllowFormattingRows:=True

        ' Transferir datos de la hoja "PRESUPUESTO"
        Set wsPresupuesto = ThisWorkbook.Sheets("PRESUPUESTO")
        wsPresupuesto.Unprotect password:=claveGeneral

        Dim rangoPresupuesto As range
        On Error Resume Next
        Set rangoPresupuesto = wsPresupuesto.UsedRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoPresupuesto Is Nothing Then
            rangoPresupuesto.Copy
            If .Bookmarks.Exists("Presupuesto_detalle") Then
                With .Bookmarks("Presupuesto_detalle").range
                    .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                    If .Tables.Count > 0 Then .Tables(1).AutoFitBehavior wdAutoFitWindow
                End With
            Else
                MsgBox "El marcador 'Presupuesto_detalle' no existe en la plantilla.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRESUPUESTO.", vbExclamation
        End If

        wsPresupuesto.Protect password:=claveGeneral, AllowFormattingRows:=True
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET'S-TDR"
    ThisWorkbook.Sheets("ET'S-TDR").Activate

    ' Eliminar archivo temporal (plantilla descargada)
    On Error Resume Next
    Kill rutaDescargaTemporal
    On Error GoTo 0

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsProductos = Nothing
    Set wsPresupuesto = Nothing
    Set wsBase = Nothing

    MsgBox "El documento se ha generado correctamente.", vbInformation
End Sub



