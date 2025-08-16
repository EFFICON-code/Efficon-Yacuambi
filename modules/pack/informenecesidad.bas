Attribute VB_Name = "informenecesidad"
Sub Informe_de_Necesidad()
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object

    ' Objetos y variables para la descarga de la plantilla
    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant
    
    Dim ws As Worksheet, wsProductos As Worksheet, wsBase As Worksheet

    ' Variables para almacenar datos
    Dim unidadRequirente As String
    Dim nombreTecnicoUnidad As String
    Dim cargoTecnicoUnidad As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String
    Dim numeroDeInforme As String
    Dim antecedente1 As String
    Dim antecedente2 As String
    Dim antecedente3 As String
    Dim antecedente4 As String
    Dim problematica As String
    Dim justificacion As String
    Dim objetivoGeneral As String
    Dim objetivosEspecificos As String
    Dim objetoDeContratacion As String
    Dim metodoDeCompra As String
    Dim competenciasEntidadMunicipal As String
    Dim capacidadInstalada As String
    Dim instalacionesEjecusion As String
    Dim instalacionesEjecusion1 As String
    Dim analisisEconomico As String
    Dim analisisBeneficio As String
    Dim analisisEficiencia As String
    Dim analisisEfectividad As String
    Dim conclusiones As String
    Dim recomendaciones As String
    Dim fechaInforme As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim partida As String
    Dim denominacion As String
    Dim cpc As String
    Dim unidadRequirente1 As String
    Dim TipoCompra As String
    Dim TipoProducto As String
    Dim tipoProcedimiento As String
    Dim estandarizacion As String

    ' Claves para desproteger/proteger
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B136
    plantillaID = wsBase.range("B136").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B136 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo de Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    ' Asignar la hoja "SECUENCIAS" y desprotegerla
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    ws.Unprotect password:=claveSecuencias
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If

    ' Leer datos de Excel
    unidadRequirente = ws.range("D2").Value
    nombreTecnicoUnidad = ws.range("G2").Value
    cargoTecnicoUnidad = ws.range("H2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    numeroDeInforme = ws.range("M2").Value
    antecedente1 = ws.range("Z2").Value
    antecedente2 = ws.range("AA2").Value
    antecedente3 = ws.range("AB2").Value
    antecedente4 = ws.range("AC2").Value
    problematica = ws.range("R2").Value
    justificacion = ws.range("AF2").Value
    objetivoGeneral = ws.range("AD2").Value
    objetivosEspecificos = ws.range("AE2").Value
    objetoDeContratacion = ws.range("Q2").Value
    metodoDeCompra = ws.range("AL2").Value
    competenciasEntidadMunicipal = ws.range("AG2").Value
    capacidadInstalada = ws.range("AK2").Value
    instalacionesEjecusion = ws.range("CX2").Value
    instalacionesEjecusion1 = ws.range("CY2").Value
    analisisEconomico = ws.range("CZ2").Value
    analisisBeneficio = ws.range("AH2").Value
    analisisEficiencia = ws.range("AI2").Value
    analisisEfectividad = ws.range("AJ2").Value
    conclusiones = ws.range("AM2").Value
    recomendaciones = ws.range("AN2").Value
    fechaInforme = ws.range("Y2").Value
    firmaTecnico = ws.range("G2").Value
    cargoTecnico = ws.range("H2").Value
    partida = ws.range("GW2").Value
    denominacion = ws.range("GX2").Value
    cpc = ws.range("GY2").Value
    unidadRequirente1 = ws.range("DA2").Value
    TipoCompra = ws.range("O2").Value
    TipoProducto = ws.range("P2").Value
    tipoProcedimiento = ws.range("S2").Value
    estandarizacion = ws.range("HD2").Value

    ' Proteger y ocultar la hoja "SECUENCIAS"
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Ruta temporal para descargar la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_InformeNecesidad_Temp.docx"

   ' Descargar la plantilla con MSXML2.ServerXMLHTTP (sin volver a declarar objHTTP y objStream)
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.status = 200 Then
        ' Guardar el archivo descargado en la ubicación temporal
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
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad") Then .Bookmarks("Nombre_Tecnico_Unidad").range.Text = nombreTecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad") Then .Bookmarks("Cargo_Tecnico_Unidad").range.Text = cargoTecnicoUnidad
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Numero_de_Informe") Then .Bookmarks("Numero_de_Informe").range.Text = numeroDeInforme
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Problematica") Then .Bookmarks("Problematica").range.Text = problematica
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Metodo_de_Compra") Then .Bookmarks("Metodo_de_Compra").range.Text = metodoDeCompra
        If .Bookmarks.Exists("Competencias_Entidad_Municipal") Then .Bookmarks("Competencias_Entidad_Municipal").range.Text = competenciasEntidadMunicipal
        If .Bookmarks.Exists("Capacidad_Instalada") Then .Bookmarks("Capacidad_Instalada").range.Text = capacidadInstalada
        If .Bookmarks.Exists("Instalaciones_Ejecusion") Then .Bookmarks("Instalaciones_Ejecusion").range.Text = instalacionesEjecusion
        If .Bookmarks.Exists("Instalaciones_Ejecusion1") Then .Bookmarks("Instalaciones_Ejecusion1").range.Text = instalacionesEjecusion1
        If .Bookmarks.Exists("Analisis_Economico") Then .Bookmarks("Analisis_Economico").range.Text = analisisEconomico
        If .Bookmarks.Exists("Analisis_Beneficio") Then .Bookmarks("Analisis_Beneficio").range.Text = analisisBeneficio
        If .Bookmarks.Exists("Analisis_Eficiencia") Then .Bookmarks("Analisis_Eficiencia").range.Text = analisisEficiencia
        If .Bookmarks.Exists("Analisis_Efectividad") Then .Bookmarks("Analisis_Efectividad").range.Text = analisisEfectividad
        If .Bookmarks.Exists("Conclusiones") Then .Bookmarks("Conclusiones").range.Text = conclusiones
        If .Bookmarks.Exists("Recomendaciones") Then .Bookmarks("Recomendaciones").range.Text = recomendaciones
        If .Bookmarks.Exists("Fecha_Informe") Then .Bookmarks("Fecha_Informe").range.Text = fechaInforme
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Partida") Then .Bookmarks("Partida").range.Text = partida
        If .Bookmarks.Exists("Denominacion") Then .Bookmarks("Denominacion").range.Text = denominacion
        If .Bookmarks.Exists("CPC") Then .Bookmarks("CPC").range.Text = cpc
        
        ' Nuevos marcadores
        If .Bookmarks.Exists("Unidad_requirente1") Then .Bookmarks("Unidad_requirente1").range.Text = unidadRequirente1
        If .Bookmarks.Exists("Tipo_compra") Then .Bookmarks("Tipo_compra").range.Text = TipoCompra
        If .Bookmarks.Exists("Tipo_producto") Then .Bookmarks("Tipo_producto").range.Text = TipoProducto
        If .Bookmarks.Exists("Tipo_procedimiento") Then .Bookmarks("Tipo_procedimiento").range.Text = tipoProcedimiento
        If .Bookmarks.Exists("Estandarizacion") Then .Bookmarks("Estandarizacion").range.Text = estandarizacion
        ' Añadir datos de productos desde la hoja PRODUCTOS
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

        ' Proteger la hoja PRODUCTOS nuevamente
        wsProductos.Protect password:=claveGeneral, Scenarios:=True, AllowFormattingRows:=True
    End With

    ' Guardar y cerrar el documento de Word
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Ubicarse en la hoja "INFORME NECESIDAD"
    ThisWorkbook.Sheets("INFORME NECESIDAD").Activate

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

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



