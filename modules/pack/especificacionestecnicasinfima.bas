Attribute VB_Name = "especificacionestecnicasinfima"
Sub Especificaciones_Tecnicas_IC()
    ' Declaración de variables
    Dim objHTTP As Object
    Dim objStream As Object
    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant

    Dim ws As Worksheet
    Dim wsProductos As Worksheet
    Dim wsBase As Worksheet

    Dim unidadRequirente As String
    Dim objetoDeContratacion As String
    Dim antecedente1 As String
    Dim antecedente2 As String
    Dim antecedente3 As String
    Dim antecedente4 As String
    Dim objetivoGeneral As String
    Dim objetivosEspecificos As String
    Dim alcance As String
    Dim metodologiaDeTrabajo As String
    Dim informacionEntidad As String
    Dim formaDePago As String
    Dim plazo As String
    Dim obligacionesContratista As String
    Dim vigenciaOferta As String
    Dim datosProforma As String
    Dim proforma As String
    Dim lugarDeEntrega As String
    Dim garantia As String
    Dim fechaElaborado As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String
    Dim marcoLegal1 As String
    Dim marcoLegal2 As String
    Dim justificacion As String
    Dim vigenciatecnologica As String
    ' Clave para desproteger la hoja
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B137
    plantillaID = wsBase.range("B137").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B137 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de descarga de la plantilla
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

    ' Leer datos de Excel
    unidadRequirente = ws.range("D2").Value
    objetoDeContratacion = ws.range("Q2").Value
    antecedente1 = ws.range("Z2").Value
    antecedente2 = ws.range("AA2").Value
    antecedente3 = ws.range("AB2").Value
    antecedente4 = ws.range("AC2").Value
    objetivoGeneral = ws.range("AD2").Value
    objetivosEspecificos = ws.range("AE2").Value
    alcance = ws.range("AQ2").Value
    metodologiaDeTrabajo = ws.range("AP2").Value
    informacionEntidad = ws.range("AR2").Value
    formaDePago = ws.range("AS2").Value
    plazo = ws.range("T2").Value
    obligacionesContratista = ws.range("BI2").Value
    vigenciaOferta = ws.range("AU2").Value
    datosProforma = ws.range("AV2").Value
    proforma = ws.range("AW2").Value
    lugarDeEntrega = ws.range("AT2").Value
    garantia = ws.range("U2").Value
    fechaElaborado = ws.range("FM2").Value
    firmaTecnico = ws.range("G2").Value
    cargoTecnico = ws.range("H2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    marcoLegal1 = ws.range("Z2").Value
    marcoLegal2 = ws.range("AL2").Value
    justificacion = ws.range("AF2").Value
    vigenciatecnologica = ws.range("V2").Value
    ' Proteger y ocultar la hoja nuevamente
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_EspecTecnicasIC_Temp.docx"
    Debug.Print "Ruta temporal: " & rutaDescargaTemporal

    ' Descargar la plantilla
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objHTTP.Open "GET", plantillaRuta, False
        objHTTP.send

    If objHTTP.status = 200 Then
        ' Guardar el archivo en la ubicación temporal
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
        Set wdApp = CreateObject("Word.Application")
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
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Alcance") Then .Bookmarks("Alcance").range.Text = alcance
        If .Bookmarks.Exists("Metodologia_de_Trabajo") Then .Bookmarks("Metodologia_de_Trabajo").range.Text = metodologiaDeTrabajo
        If .Bookmarks.Exists("Informacion_Entidad") Then .Bookmarks("Informacion_Entidad").range.Text = informacionEntidad
        If .Bookmarks.Exists("Forma_de_Pago") Then .Bookmarks("Forma_de_Pago").range.Text = formaDePago
        If .Bookmarks.Exists("Plazo") Then .Bookmarks("Plazo").range.Text = plazo
        If .Bookmarks.Exists("Obligaciones_Contratista") Then .Bookmarks("Obligaciones_Contratista").range.Text = obligacionesContratista
        If .Bookmarks.Exists("Vigencia_Oferta") Then .Bookmarks("Vigencia_Oferta").range.Text = vigenciaOferta
        If .Bookmarks.Exists("Datos_Proforma") Then .Bookmarks("Datos_Proforma").range.Text = datosProforma
        If .Bookmarks.Exists("Proforma") Then .Bookmarks("Proforma").range.Text = proforma
        If .Bookmarks.Exists("Lugar_de_Entrega") Then .Bookmarks("Lugar_de_Entrega").range.Text = lugarDeEntrega
        If .Bookmarks.Exists("Garantia") Then .Bookmarks("Garantia").range.Text = garantia
        If .Bookmarks.Exists("Fecha_Elaborado") Then .Bookmarks("Fecha_Elaborado").range.Text = fechaElaborado
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Marco_Legal1") Then .Bookmarks("Marco_Legal1").range.Text = marcoLegal1
        If .Bookmarks.Exists("Marco_Legal2") Then .Bookmarks("Marco_Legal2").range.Text = marcoLegal2
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Vigencia_Tecnologica") Then .Bookmarks("Vigencia_Tecnologica").range.Text = vigenciatecnologica
        ' Añadir datos de productos desde el rango visible
        Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
        wsProductos.Unprotect password:=claveGeneral

        Dim rangoVisible As range
        On Error Resume Next
        Set rangoVisible = wsProductos.range("Productosdt").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            rangoVisible.Copy
            If .Bookmarks.Exists("Productos") Then
                With .Bookmarks("Productos").range
                    .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                    .Tables(1).AutoFitBehavior wdAutoFitWindow
                End With
            Else
                MsgBox "El marcador 'Productos' no existe en la plantilla.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRODUCTOS.", vbExclamation
        End If
        
        wsProductos.Protect password:=claveGeneral, Scenarios:=True, AllowFormattingRows:=True
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET'S-TDR"
    ThisWorkbook.Sheets("ET'S-TDR").Activate

    ' Eliminar el archivo temporal
    On Error Resume Next
    Kill rutaDescargaTemporal
    On Error GoTo 0

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsProductos = Nothing
    Set wsBase = Nothing

    MsgBox "El documento se ha generado correctamente.", vbInformation
End Sub

