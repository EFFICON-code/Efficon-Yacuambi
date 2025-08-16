Attribute VB_Name = "tdretcatalogoelectronico"
Sub TDR_ET_Catalogo_Electronico()
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
    Dim wsProductos As Worksheet
    Dim wsBase As Worksheet

    ' Claves
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Variables para marcadores
    Dim titulo As String
    Dim objetoDeContratacion As String
    Dim unidadRequirente As String
    Dim antecedente1 As String
    Dim antecedente2 As String
    Dim antecedente3 As String
    Dim antecedente4 As String
    Dim objetivoGeneral As String
    Dim objetivosEspecificos As String
    Dim justificacion As String
    Dim objetoDeContratacion1 As String
    Dim tipoDeCompra As String
    Dim tipoDeProceso As String
    Dim tipoRecepcion As String
    Dim fechaElaborado As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D133
    plantillaID = wsBase.range("D133").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D133 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para la ubicación donde guardar el documento terminado
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

    ' Leer datos de Excel desde la hoja "SECUENCIAS"
    titulo = ws.range("AO2").Value
    objetoDeContratacion = ws.range("Q2").Value
    unidadRequirente = ws.range("D2").Value
    antecedente1 = ws.range("Z2").Value
    antecedente2 = ws.range("AA2").Value
    antecedente3 = ws.range("AB2").Value
    antecedente4 = ws.range("AC2").Value
    objetivoGeneral = ws.range("AD2").Value
    objetivosEspecificos = ws.range("AE2").Value
    justificacion = ws.range("AF2").Value
    objetoDeContratacion1 = ws.range("Q2").Value
    tipoDeCompra = ws.range("O2").Value
    tipoDeProceso = ws.range("S2").Value
    tipoRecepcion = ws.range("AX2").Value
    fechaElaborado = ws.range("FM2").Value
    firmaTecnico = ws.range("G2").Value
    cargoTecnico = ws.range("H2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value

    ' Proteger y ocultar nuevamente la hoja "SECUENCIAS"
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla a una ruta temporal
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_TDR_ET_Catalogo_Temp.docx"
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Tipo binario
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
        If .Bookmarks.Exists("Titulo") Then .Bookmarks("Titulo").range.Text = titulo
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Tipo_de_Compra") Then .Bookmarks("Tipo_de_Compra").range.Text = tipoDeCompra
        If .Bookmarks.Exists("Tipo_de_Proceso") Then .Bookmarks("Tipo_de_Proceso").range.Text = tipoDeProceso
        If .Bookmarks.Exists("Tipo_Recepcion") Then .Bookmarks("Tipo_Recepcion").range.Text = tipoRecepcion
        If .Bookmarks.Exists("Fecha_Elaborado") Then .Bookmarks("Fecha_Elaborado").range.Text = fechaElaborado
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad

        ' Añadir datos de productos desde la hoja "PRODUCTOS"
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
                    .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                End With
            Else
                MsgBox "El marcador 'Productos' no existe en la plantilla.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRODUCTOS.", vbExclamation
        End If

        ' Proteger la hoja "PRODUCTOS" nuevamente
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

    ' Eliminar el archivo temporal descargado
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

