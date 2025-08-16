Attribute VB_Name = "terminosregimenespecial"
Sub TDR_ET_Regimen_Especial()
    ' --- DECLARACIONES ---
    Dim wdApp As Object, wdDoc As Object
    Dim objHTTP As Object, objStream As Object
    Dim rutaDescargaTemporal As String, plantillaID As String, plantillaRuta As String, guardarRuta As Variant

    ' Declaración de todas las hojas necesarias
    Dim ws As Worksheet, wsBase As Worksheet, wsProductos As Worksheet, wsPresupuesto As Worksheet
    Dim wsPersonalT As Worksheet, wsExpPT As Worksheet, wsEquipo As Worksheet
    
    ' Claves para desproteger
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Variables para marcadores
    Dim entidad As String, titulo As String, tipoDeProcedimiento As String, objetoDeContratacion As String
    Dim unidadRequirente As String, antecedente1 As String, antecedente2 As String, antecedente3 As String
    Dim antecedente4 As String, objetoDeContratacion1 As String, codigoCPC As String, modalidadContratacion As String
    Dim justificacion As String, objetivoGeneral As String, objetivosEspecificos As String, alcance As String
    Dim metodologiaDeTrabajo As String, informacionEntidad As String, experienciaGeneral As String, montoGeneral As String
    Dim porContratoG As String, experienciaEspecifica As String, montoEspecifica As String, porContratoE As String
    Dim presupuestoReferencial As String, valorLetras As String, plazo As String, formaDePago As String
    Dim tipoEntrega As String, lugarDeEntrega As String, garantia As String, vigenciaOferta As String
    Dim formaPresentacion As String, obligacionesContratista As String, recepcion As String, nombreTecnicoUnidad As String
    Dim cargoTecnicoUnidad As String, nombreTitularUnidad As String, cargoTitularUnidad As String, fechaElaboracion As String

    ' --- INICIO DE LA EJECUCIÓN ---
    On Error GoTo GestorErrores

    ThisWorkbook.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral
    plantillaID = wsBase.range("D138").Value
    wsBase.Protect password:=claveGeneral
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D138 de la hoja BBDD.", vbCritical
        GoTo SalidaLimpia
    End If
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Diálogo para guardar el archivo final
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        GoTo SalidaLimpia
    End If

    ' --- LECTURA DE DATOS DE LA HOJA "SECUENCIAS" ---
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Unprotect password:=claveSecuencias

    ' Usamos la función LeerCeldaComoString para una lectura segura
    entidad = LeerCeldaComoString(ws.range("A2"))
    titulo = LeerCeldaComoString(ws.range("AO2"))
    ' ... (resto de la lectura de datos se mantiene igual) ...
    tipoDeProcedimiento = LeerCeldaComoString(ws.range("S2"))
    objetoDeContratacion = LeerCeldaComoString(ws.range("Q2"))
    unidadRequirente = LeerCeldaComoString(ws.range("D2"))
    antecedente1 = LeerCeldaComoString(ws.range("Z2"))
    antecedente2 = LeerCeldaComoString(ws.range("AA2"))
    antecedente3 = LeerCeldaComoString(ws.range("AB2"))
    antecedente4 = LeerCeldaComoString(ws.range("AC2"))
    objetoDeContratacion1 = LeerCeldaComoString(ws.range("Q2"))
    codigoCPC = LeerCeldaComoString(ws.range("BA2"))
    modalidadContratacion = LeerCeldaComoString(ws.range("CM2"))
    justificacion = LeerCeldaComoString(ws.range("AF2"))
    objetivoGeneral = LeerCeldaComoString(ws.range("AD2"))
    objetivosEspecificos = LeerCeldaComoString(ws.range("AE2"))
    alcance = LeerCeldaComoString(ws.range("AQ2"))
    metodologiaDeTrabajo = LeerCeldaComoString(ws.range("AP2"))
    informacionEntidad = LeerCeldaComoString(ws.range("AR2"))
    experienciaGeneral = LeerCeldaComoString(ws.range("BC2"))
    montoGeneral = LeerCeldaComoString(ws.range("BD2"))
    porContratoG = LeerCeldaComoString(ws.range("BE2"))
    experienciaEspecifica = LeerCeldaComoString(ws.range("BF2"))
    montoEspecifica = LeerCeldaComoString(ws.range("BG2"))
    porContratoE = LeerCeldaComoString(ws.range("BH2"))
    presupuestoReferencial = LeerCeldaComoString(ws.range("BV2"))
    valorLetras = LeerCeldaComoString(ws.range("BW2"))
    plazo = LeerCeldaComoString(ws.range("T2"))
    formaDePago = LeerCeldaComoString(ws.range("AS2"))
    tipoEntrega = LeerCeldaComoString(ws.range("CL2"))
    lugarDeEntrega = LeerCeldaComoString(ws.range("AT2"))
    garantia = LeerCeldaComoString(ws.range("U2"))
    vigenciaOferta = LeerCeldaComoString(ws.range("AU2"))
    formaPresentacion = LeerCeldaComoString(ws.range("AW2"))
    obligacionesContratista = LeerCeldaComoString(ws.range("BI2"))
    recepcion = LeerCeldaComoString(ws.range("AX2"))
    nombreTecnicoUnidad = LeerCeldaComoString(ws.range("G2"))
    cargoTecnicoUnidad = LeerCeldaComoString(ws.range("H2"))
    nombreTitularUnidad = LeerCeldaComoString(ws.range("E2"))
    cargoTitularUnidad = LeerCeldaComoString(ws.range("F2"))
    fechaElaboracion = LeerCeldaComoString(ws.range("GZ2"))

    ws.Protect password:=claveSecuencias
    ws.Visible = xlSheetHidden

    ' --- DESCARGA DE PLANTILLA ---
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_TDRRegimenEspecial_Temp.docx"
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send
    If objHTTP.status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1
        objStream.Open
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile rutaDescargaTemporal, 2
        objStream.Close
    Else
        MsgBox "Error al descargar la plantilla. Código de estado: " & objHTTP.status & " - " & objHTTP.statusText, vbCritical
        GoTo SalidaLimpia
    End If

    ' --- PROCESAMIENTO CON WORD ---
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject(Class:="Word.Application")
    On Error GoTo GestorErrores
    If wdApp Is Nothing Then
        MsgBox "No se pudo iniciar Microsoft Word.", vbCritical
        GoTo SalidaLimpia
    End If

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(rutaDescargaTemporal)
    If wdDoc Is Nothing Then
        MsgBox "No se pudo abrir el documento de Word.", vbCritical
        GoTo SalidaLimpia
    End If

    ' --- LLENADO DE MARCADORES Y TABLAS EN WORD ---
    With wdDoc
        ' Llenado de marcadores de texto (se mantiene igual)
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Titulo") Then .Bookmarks("Titulo").range.Text = titulo
        ' ... (todos los demás marcadores de texto) ...
        If .Bookmarks.Exists("Tipo_de_Procedimiento") Then .Bookmarks("Tipo_de_Procedimiento").range.Text = tipoDeProcedimiento
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Antecedente1") Then .Bookmarks("Antecedente1").range.Text = antecedente1
        If .Bookmarks.Exists("Antecedente2") Then .Bookmarks("Antecedente2").range.Text = antecedente2
        If .Bookmarks.Exists("Antecedente3") Then .Bookmarks("Antecedente3").range.Text = antecedente3
        If .Bookmarks.Exists("Antecedente4") Then .Bookmarks("Antecedente4").range.Text = antecedente4
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Codigo_CPC") Then .Bookmarks("Codigo_CPC").range.Text = codigoCPC
        If .Bookmarks.Exists("Modalidad_Contratacion") Then .Bookmarks("Modalidad_Contratacion").range.Text = modalidadContratacion
        If .Bookmarks.Exists("Justificacion") Then .Bookmarks("Justificacion").range.Text = justificacion
        If .Bookmarks.Exists("Objetivo_General") Then .Bookmarks("Objetivo_General").range.Text = objetivoGeneral
        If .Bookmarks.Exists("Objetivos_Especificos") Then .Bookmarks("Objetivos_Especificos").range.Text = objetivosEspecificos
        If .Bookmarks.Exists("Alcance") Then .Bookmarks("Alcance").range.Text = alcance
        If .Bookmarks.Exists("Metodologia_de_Trabajo") Then .Bookmarks("Metodologia_de_Trabajo").range.Text = metodologiaDeTrabajo
        If .Bookmarks.Exists("Informacion_Entidad") Then .Bookmarks("Informacion_Entidad").range.Text = informacionEntidad
        If .Bookmarks.Exists("Experiencia_General") Then .Bookmarks("Experiencia_General").range.Text = experienciaGeneral
        If .Bookmarks.Exists("Monto_General") Then .Bookmarks("Monto_General").range.Text = montoGeneral
        If .Bookmarks.Exists("Por_contrato_G") Then .Bookmarks("Por_contrato_G").range.Text = porContratoG
        If .Bookmarks.Exists("Experiencia_Especifica") Then .Bookmarks("Experiencia_Especifica").range.Text = experienciaEspecifica
        If .Bookmarks.Exists("Monto_Especifica") Then .Bookmarks("Monto_Especifica").range.Text = montoEspecifica
        If .Bookmarks.Exists("Por_contrato_E") Then .Bookmarks("Por_contrato_E").range.Text = porContratoE
        If .Bookmarks.Exists("Presupuesto_Referencial") Then .Bookmarks("Presupuesto_Referencial").range.Text = presupuestoReferencial
        If .Bookmarks.Exists("Valor_Letras") Then .Bookmarks("Valor_Letras").range.Text = valorLetras
        If .Bookmarks.Exists("Plazo") Then .Bookmarks("Plazo").range.Text = plazo
        If .Bookmarks.Exists("Forma_de_Pago") Then .Bookmarks("Forma_de_Pago").range.Text = formaDePago
        If .Bookmarks.Exists("Tipo_Entrega") Then .Bookmarks("Tipo_Entrega").range.Text = tipoEntrega
        If .Bookmarks.Exists("Lugar_de_Entrega") Then .Bookmarks("Lugar_de_Entrega").range.Text = lugarDeEntrega
        If .Bookmarks.Exists("Garantia") Then .Bookmarks("Garantia").range.Text = garantia
        If .Bookmarks.Exists("Vigencia_Oferta") Then .Bookmarks("Vigencia_Oferta").range.Text = vigenciaOferta
        If .Bookmarks.Exists("Forma_Presentacion") Then .Bookmarks("Forma_Presentacion").range.Text = formaPresentacion
        If .Bookmarks.Exists("Obligaciones_Contratista") Then .Bookmarks("Obligaciones_Contratista").range.Text = obligacionesContratista
        If .Bookmarks.Exists("Recepcion") Then .Bookmarks("Recepcion").range.Text = recepcion
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad") Then .Bookmarks("Nombre_Tecnico_Unidad").range.Text = nombreTecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad") Then .Bookmarks("Cargo_Tecnico_Unidad").range.Text = cargoTecnicoUnidad
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Fecha_elaboracion") Then .Bookmarks("Fecha_elaboracion").range.Text = fechaElaboracion

        ' --- Copiado de todas las tablas ---
        CopiarRangoVisibleAWord "PersonalT", "A1:F11", "Personal_Tecnico", wdDoc, claveGeneral
        CopiarRangoVisibleAWord "ExperienciaPT", "A1:F11", "Exp_Personal_Tecnico", wdDoc, claveGeneral
        CopiarRangoVisibleAWord "EquipoMinimo", "A1:C11", "Equipo_Minimo", wdDoc, claveGeneral
        CopiarRangoVisibleAWord "PRODUCTOS", "Productosdt", "Productos", wdDoc, claveGeneral
        CopiarRangoVisibleAWord "PRESUPUESTO", "A1:Z100", "Presupuesto_detalle", wdDoc, claveGeneral
        
        ' --- INICIO DE LA NUEVA ADICIÓN: Copiado de tablas en los nuevos marcadores ---
        CopiarRangoVisibleAWord "PRODUCTOS", "Productosdt", "Productos2", wdDoc, claveGeneral
        CopiarRangoVisibleAWord "PersonalT", "A1:F11", "Persomal_Tecnico2", wdDoc, claveGeneral
        ' --- FIN DE LA NUEVA ADICIÓN ---

    End With

    ' --- FINALIZACIÓN Y GUARDADO ---
    wdDoc.SaveAs2 fileName:=guardarRuta
    MsgBox "El documento se ha generado correctamente en: " & vbCrLf & guardarRuta, vbInformation

SalidaLimpia:
    ' Rutina de limpieza para salir de forma ordenada
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    If Len(rutaDescargaTemporal) > 0 And Dir(rutaDescargaTemporal) <> "" Then Kill rutaDescargaTemporal
    
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True
    Application.CutCopyMode = False
    
    ThisWorkbook.Sheets("ET'S-TDR").Activate
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsBase = Nothing
    Set wsProductos = Nothing
    Set wsPresupuesto = Nothing
    Set wsPersonalT = Nothing
    Set wsExpPT = Nothing
    Set wsEquipo = Nothing
    Set objHTTP = Nothing
    Set objStream = Nothing
    Exit Sub

GestorErrores:
    MsgBox "Ha ocurrido un error inesperado:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Error en la Ejecución"
    GoTo SalidaLimpia
End Sub


' =============================================================================================
' ===================          SUBRUTINAS Y FUNCIONES AUXILIARES          ===================
' =============================================================================================

Private Sub CopiarRangoVisibleAWord(ByVal wsName As String, ByVal rangeAddress As String, ByVal bookmarkName As String, ByRef wdDoc As Object, ByVal password As String)
    ' Objetivo: Encapsula la lógica para copiar un rango de celdas visibles a un marcador de Word.
    Dim ws As Worksheet, rngToCopy As range
    Dim sheetWasHidden As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    If ws Is Nothing Then
        MsgBox "Advertencia: La hoja '" & wsName & "' no fue encontrada. Se omitirá este paso.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ws.Unprotect password:=password
    
    ' Manejar visibilidad de la hoja
    If ws.Visible <> xlSheetVisible Then
        sheetWasHidden = True
        ws.Visible = xlSheetVisible
    Else
        sheetWasHidden = False
    End If
    
    ' Copiar rango visible
    On Error Resume Next
    ' Usamos .Range(rangeAddress) si es una dirección fija, o un rango con nombre.
    If rangeAddress = "UsedRange" Then
        Set rngToCopy = ws.UsedRange.SpecialCells(xlCellTypeVisible)
    Else
        Set rngToCopy = ws.range(rangeAddress).SpecialCells(xlCellTypeVisible)
    End If
    On Error GoTo 0

    If Not rngToCopy Is Nothing Then
        If wdDoc.Bookmarks.Exists(bookmarkName) Then
            rngToCopy.Copy
            With wdDoc.Bookmarks(bookmarkName).range
                .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                If .Tables.Count > 0 Then .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
            End With
        Else
            MsgBox "Advertencia: El marcador '" & bookmarkName & "' no existe en la plantilla. Se omitirá este paso.", vbExclamation
        End If
    End If
    
    ' Restaurar estado de la hoja
    Application.CutCopyMode = False
    If sheetWasHidden Then ws.Visible = xlSheetHidden
    ws.Protect password:=password, AllowFormattingRows:=True
End Sub


Private Function LeerCeldaComoString(ByVal Rango As range) As String
    ' Objetivo: Lee el valor de una celda de forma segura, devolviendo "" si está vacía o contiene un error.
    On Error Resume Next
    If IsError(Rango.Value) Or IsEmpty(Rango.Value) Then
        LeerCeldaComoString = ""
    Else
        LeerCeldaComoString = CStr(Rango.Value)
    End If
    On Error GoTo 0
End Function

