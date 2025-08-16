Attribute VB_Name = "especificacionessubastainversa"
Sub Especificiones_Terminos_SIE()
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object

    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant
    
    Dim ws As Worksheet
    Dim wsProductos As Worksheet
    Dim wsBase As Worksheet
    
    Dim rangoVisible As range

    ' Clave para la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave general
    Const claveGeneral As String = "PROEST2023"

    ' Variables para marcadores
    Dim objetoDeContratacion As String
    Dim unidadRequirente As String
    Dim antecedente1 As String
    Dim antecedente2 As String
    Dim antecedente3 As String
    Dim antecedente4 As String
    Dim justificacion As String
    Dim objetivoGeneral As String
    Dim objetivosEspecificos As String
    Dim alcance As String
    Dim metodologiaDeTrabajo As String
    Dim informacionEntidad As String
    Dim objetoDeContratacion1 As String
    Dim presupuestoReferencial As String
    Dim valorLetras As String
    Dim plazo As String
    Dim vigenciaOferta As String
    Dim formaDePago As String
    Dim lugarDeEntrega As String
    Dim variacionPuja As String
    Dim tiempoPuja As String
    Dim transferenciaTecnologica As String
    Dim requisitosTT As String
    Dim garantia As String
    Dim canje As String
    Dim fielCumplimiento As String
    Dim buenUsoAnticipo As String
    Dim obligacionesContratista As String
    Dim tipoDeRecepcion As String
    Dim fechaElaborado As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String
    Dim vigenciatecnologica As String
    Dim valordinero As String
    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral
    
    ' Leer el ID de la plantilla desde la celda D139
    plantillaID = wsBase.range("D139").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D139 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID
    
    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar dónde se guardará el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "DocumentoTerminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    ' Asignar la hoja "SECUENCIAS" y desprotegerla
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    ws.Visible = xlSheetVisible
    ws.Unprotect password:=claveSecuencias

    ' Leer datos desde la hoja "SECUENCIAS"
    objetoDeContratacion = ws.range("Q2").Value
    unidadRequirente = ws.range("D2").Value
    antecedente1 = ws.range("Z2").Value
    antecedente2 = ws.range("AA2").Value
    antecedente3 = ws.range("AB2").Value
    antecedente4 = ws.range("AC2").Value
    justificacion = ws.range("AF2").Value
    objetivoGeneral = ws.range("AD2").Value
    objetivosEspecificos = ws.range("AE2").Value
    alcance = ws.range("AQ2").Value
    metodologiaDeTrabajo = ws.range("AP2").Value
    informacionEntidad = ws.range("AR2").Value
    objetoDeContratacion1 = ws.range("Q2").Value
    presupuestoReferencial = ws.range("BV2").Value
    valorLetras = ws.range("BW2").Value
    plazo = ws.range("T2").Value
    vigenciaOferta = ws.range("AU2").Value
    formaDePago = ws.range("AS2").Value
    lugarDeEntrega = ws.range("AT2").Value
    variacionPuja = ws.range("FH2").Value
    tiempoPuja = ws.range("FI2").Value
    transferenciaTecnologica = ws.range("CN2").Value
    requisitosTT = ws.range("CO2").Value
    garantia = ws.range("U2").Value
    canje = ws.range("FJ2").Value
    fielCumplimiento = ws.range("FF2").Value
    buenUsoAnticipo = ws.range("FG2").Value
    obligacionesContratista = ws.range("BI2").Value
    tipoDeRecepcion = ws.range("AX2").Value
    fechaElaborado = ws.range("FM2").Value
    firmaTecnico = ws.range("G2").Value
    cargoTecnico = ws.range("H2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    vigenciatecnologica = ws.range("V2").Value
    valordinero = ws.range("HF2").Value
    ' Proteger y ocultar la hoja "SECUENCIAS"
    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla a una ruta temporal
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_Especificaciones_SIE_Temp.docx"
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

    ' Insertar datos en los marcadores
    With wdDoc
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Alcance") Then .Bookmarks("Alcance").range.Text = alcance
        If .Bookmarks.Exists("Metodologia_de_Trabajo") Then .Bookmarks("Metodologia_de_Trabajo").range.Text = metodologiaDeTrabajo
        If .Bookmarks.Exists("Informacion_Entidad") Then .Bookmarks("Informacion_Entidad").range.Text = informacionEntidad
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Presupuesto_Referencial") Then .Bookmarks("Presupuesto_Referencial").range.Text = presupuestoReferencial
        If .Bookmarks.Exists("Valor_Letras") Then .Bookmarks("Valor_Letras").range.Text = valorLetras
        If .Bookmarks.Exists("Plazo") Then .Bookmarks("Plazo").range.Text = plazo
        If .Bookmarks.Exists("Vigencia_Oferta") Then .Bookmarks("Vigencia_Oferta").range.Text = vigenciaOferta
        If .Bookmarks.Exists("Forma_de_Pago") Then .Bookmarks("Forma_de_Pago").range.Text = formaDePago
        If .Bookmarks.Exists("Lugar_de_Entrega") Then .Bookmarks("Lugar_de_Entrega").range.Text = lugarDeEntrega
        If .Bookmarks.Exists("Variacion_Puja") Then .Bookmarks("Variacion_Puja").range.Text = variacionPuja
        If .Bookmarks.Exists("Tiempo_Puja") Then .Bookmarks("Tiempo_Puja").range.Text = tiempoPuja
        If .Bookmarks.Exists("Transferencia_Tecnologica") Then .Bookmarks("Transferencia_Tecnologica").range.Text = transferenciaTecnologica
        If .Bookmarks.Exists("Requisitos_TT") Then .Bookmarks("Requisitos_TT").range.Text = requisitosTT
        If .Bookmarks.Exists("Garantia") Then .Bookmarks("Garantia").range.Text = garantia
        If .Bookmarks.Exists("Canje") Then .Bookmarks("Canje").range.Text = canje
        If .Bookmarks.Exists("Fiel_Cumplimiento") Then .Bookmarks("Fiel_Cumplimiento").range.Text = fielCumplimiento
        If .Bookmarks.Exists("Buen_Uso_Anticipo") Then .Bookmarks("Buen_Uso_Anticipo").range.Text = buenUsoAnticipo
        If .Bookmarks.Exists("Obligaciones_Contratista") Then .Bookmarks("Obligaciones_Contratista").range.Text = obligacionesContratista
        If .Bookmarks.Exists("Tipo_de_Recepcion") Then .Bookmarks("Tipo_de_Recepcion").range.Text = tipoDeRecepcion
        If .Bookmarks.Exists("Fecha_Elaborado") Then .Bookmarks("Fecha_Elaborado").range.Text = fechaElaborado
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Vigencia_Tecnologica") Then .Bookmarks("Vigencia_Tecnologica").range.Text = vigenciatecnologica
        If .Bookmarks.Exists("Valor_Dinero") Then .Bookmarks("Valor_Dinero").range.Text = valordinero
        ' Trabajar con la hoja "PRODUCTOS"
        Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
        wsProductos.Unprotect password:=claveGeneral

        On Error Resume Next
        Set rangoVisible = wsProductos.range("Productosdt").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            rangoVisible.Copy
            If .Bookmarks.Exists("Productos") Then
                With .Bookmarks("Productos").range
                    .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                    If .Tables.Count > 0 Then .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                End With
            Else
                MsgBox "El marcador 'Productos' no existe en la plantilla.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja 'PRODUCTOS'.", vbExclamation
        End If

        wsProductos.Protect password:=claveGeneral, AllowFormattingRows:=True
    End With

    ' Guardar y cerrar documento
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


