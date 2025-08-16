Attribute VB_Name = "estudiodemercado"
Sub Estudio_de_Mercado()
    ' Declaración de variables
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim guardarRuta As Variant

    ' Objetos para descarga
    Dim objHTTP As Object
    Dim objStream As Object
    Dim rutaDescargaTemporal As String
    Dim plantillaID As String
    Dim plantillaRuta As String

    ' Hojas
    Dim ws As Worksheet
    Dim wsProductos As Worksheet
    Dim wsAplicabilidad As Worksheet
    Dim wsPreciosAdjudicados As Worksheet
    Dim wsPreciosActualizados As Worksheet
    Dim wsPreciosProformas As Worksheet
    Dim wsPresupuesto As Worksheet
    Dim wsBase As Worksheet

    ' Variables de datos
    Dim unidadRequirente As String
    Dim objetoDeContratacion As String
    Dim analisisMercado As String
    Dim presupuestoReferencial As String
    Dim valorLetras As String
    Dim fechaElaborado As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim tipoDeCompra As String
    Dim tipoDeProceso As String
    Dim canton As String
    Dim valordinero As String
    Dim rangoVisible As range

    ' Claves y constantes
    Dim CLAVE As String, claveOtrasHojas As String
    Const xlSheetVisible As Long = -1
    Const xlSheetHidden As Long = 0
    Const xlSheetVeryHidden As Long = 2

    ' Inicializar claves
    CLAVE = "Admin1991"
    claveOtrasHojas = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveOtrasHojas

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveOtrasHojas

    ' Leer el ID de la plantilla desde la celda D141
    plantillaID = wsBase.range("D141").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D141 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger la hoja "BBDD" nuevamente
    wsBase.Protect password:=claveOtrasHojas

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        GoTo Fin
    End If

    ' Asignar la hoja de trabajo "SECUENCIAS"
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=CLAVE

    ' Leer datos desde la hoja "SECUENCIAS"
    unidadRequirente = CStr(ws.range("D2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    analisisMercado = CStr(ws.range("BM2").Value)
    presupuestoReferencial = CStr(ws.range("BV2").Value)
    valorLetras = CStr(ws.range("BW2").Value)
    fechaElaborado = CStr(ws.range("BL2").Value)
    firmaTecnico = CStr(ws.range("G2").Value)
    cargoTecnico = CStr(ws.range("H2").Value)
    tipoDeCompra = CStr(ws.range("O2").Value)
    tipoDeProceso = CStr(ws.range("S2").Value)
    canton = CStr(ws.range("FQ2").Value)
    valordinero = CStr(ws.range("HE2").Value)
    ' Proteger y ocultar la hoja "SECUENCIAS"
    ws.Protect password:=CLAVE
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla en la carpeta temporal
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_Estudio_de_Mercado_Temp.docx"
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
        MsgBox "Error al descargar la plantilla. Revise la conexión o el enlace." & vbCrLf & _
               "Código de estado: " & objHTTP.status & " - " & objHTTP.statusText, vbExclamation
        GoTo Fin
    End If

    ' Iniciar Word y abrir la plantilla descargada
    On Error GoTo ErrorHandler
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(rutaDescargaTemporal)

    ' Insertar datos en los marcadores de la plantilla
    With wdDoc
        If .Bookmarks.Exists("Unidad_Requirente") Then
            .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        End If
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then
            .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        End If
        If .Bookmarks.Exists("Tipo_de_Compra") Then
            .Bookmarks("Tipo_de_Compra").range.Text = tipoDeCompra
        End If
        If .Bookmarks.Exists("Tipo_de_Proceso") Then
            .Bookmarks("Tipo_de_Proceso").range.Text = tipoDeProceso
        End If
        If .Bookmarks.Exists("Canton") Then
            .Bookmarks("Canton").range.Text = canton
        End If
        If .Bookmarks.Exists("Analisis_Mercado") Then
            .Bookmarks("Analisis_Mercado").range.Text = analisisMercado
        End If
        If .Bookmarks.Exists("Presupuesto_Referencial") Then
            .Bookmarks("Presupuesto_Referencial").range.Text = presupuestoReferencial
        End If
        If .Bookmarks.Exists("Valor_Letras") Then
            .Bookmarks("Valor_Letras").range.Text = valorLetras
        End If
        If .Bookmarks.Exists("Fecha_Elaborado") Then
            .Bookmarks("Fecha_Elaborado").range.Text = fechaElaborado
        End If
        If .Bookmarks.Exists("Firma_Tecnico") Then
            .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        End If
        If .Bookmarks.Exists("Cargo_Tecnico") Then
            .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        End If

        If .Bookmarks.Exists("Valor_Dinero") Then
            .Bookmarks("Valor_Dinero").range.Text = valordinero
        End If

        ' Transferir datos desde la hoja "PRODUCTOS"
        Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
        wsProductos.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsProductos.range("Productosdt").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Productos") Then
                    With .Bookmarks("Productos").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                End If
            Else
                MsgBox "El rango 'Productosdt' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRODUCTOS.", vbExclamation
        End If

        ' Proteger la hoja "PRODUCTOS" con opciones adicionales
        wsProductos.Protect password:=claveOtrasHojas, Scenarios:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True

        ' --- Transferir datos de "APLICABILIDAD" ---
        Set wsAplicabilidad = ThisWorkbook.Sheets("APLICABILIDAD")
        wsAplicabilidad.Visible = xlSheetVisible
        wsAplicabilidad.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsAplicabilidad.range("A1:C45").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Aplicabilidad") Then
                    With .Bookmarks("Aplicabilidad").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                Else
                    MsgBox "El marcador 'Aplicabilidad' no se encontró en la plantilla."
                End If
            Else
                MsgBox "El rango en 'APLICABILIDAD' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja APLICABILIDAD.", vbExclamation
        End If

        wsAplicabilidad.Protect password:=claveOtrasHojas, AllowFormattingRows:=True, AllowFormattingColumns:=True
        wsAplicabilidad.Visible = xlSheetHidden

        ' --- Transferir datos de "PRECIOS_ADJUDICADOS" ---
        Set wsPreciosAdjudicados = ThisWorkbook.Sheets("PRECIOS_ADJUDICADOS")
        wsPreciosAdjudicados.Visible = xlSheetVisible
        wsPreciosAdjudicados.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsPreciosAdjudicados.range("A1:H800").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Precios_Adjudicados") Then
                    With .Bookmarks("Precios_Adjudicados").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                Else
                    MsgBox "El marcador 'Precios_Adjudicados' no se encontró en la plantilla."
                End If
            Else
                MsgBox "El rango en 'PRECIOS_ADJUDICADOS' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRECIOS_ADJUDICADOS.", vbExclamation
        End If

        wsPreciosAdjudicados.Protect password:=claveOtrasHojas, AllowFormattingRows:=True, AllowFormattingColumns:=True
        wsPreciosAdjudicados.Visible = xlSheetHidden

        ' --- Transferir datos de "PRECIOS_ACTUALIZADOS" ---
        Set wsPreciosActualizados = ThisWorkbook.Sheets("PRECIOS_ACTUALIZADOS")
        wsPreciosActualizados.Visible = xlSheetVisible
        wsPreciosActualizados.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsPreciosActualizados.range("A1:H800").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Precios_Actualizados") Then
                    With .Bookmarks("Precios_Actualizados").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                Else
                    MsgBox "El marcador 'Precios_Actualizados' no se encontró en la plantilla."
                End If
            Else
                MsgBox "El rango en 'PRECIOS_ACTUALIZADOS' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRECIOS_ACTUALIZADOS.", vbExclamation
        End If

        wsPreciosActualizados.Protect password:=claveOtrasHojas, AllowFormattingRows:=True, AllowFormattingColumns:=True
        wsPreciosActualizados.Visible = xlSheetHidden

        ' --- Transferir datos de "PRECIOS_PROFORMAS" ---
        Set wsPreciosProformas = ThisWorkbook.Sheets("PRECIOS_PROFORMAS")
        wsPreciosProformas.Visible = xlSheetVisible
        wsPreciosProformas.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsPreciosProformas.range("A1:H800").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Precios_Proformas") Then
                    With .Bookmarks("Precios_Proformas").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                Else
                    MsgBox "El marcador 'Precios_Proformas' no se encontró en la plantilla."
                End If
            Else
                MsgBox "El rango en 'PRECIOS_PROFORMAS' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRECIOS_PROFORMAS.", vbExclamation
        End If

        wsPreciosProformas.Protect password:=claveOtrasHojas, AllowFormattingRows:=True, AllowFormattingColumns:=True
        wsPreciosProformas.Visible = xlSheetHidden

        ' --- Transferir datos de "PRESUPUESTO" ---
        Set wsPresupuesto = ThisWorkbook.Sheets("PRESUPUESTO")
        wsPresupuesto.Visible = xlSheetVisible
        wsPresupuesto.Unprotect password:=claveOtrasHojas

        On Error Resume Next
        Set rangoVisible = wsPresupuesto.range("A1:G801").SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            If Application.WorksheetFunction.CountA(rangoVisible) > 0 Then
                rangoVisible.Copy
                If .Bookmarks.Exists("Detalle_Presupuesto") Then
                    With .Bookmarks("Detalle_Presupuesto").range
                        .Paste
                        If .Tables.Count > 0 Then
                            .Tables(1).AutoFitBehavior 1 ' wdAutoFitWindow
                        End If
                    End With
                Else
                    MsgBox "El marcador 'Detalle_Presupuesto' no se encontró en la plantilla."
                End If
            Else
                MsgBox "El rango en 'PRESUPUESTO' está vacío.", vbExclamation
            End If
        Else
            MsgBox "No hay datos visibles para copiar en la hoja PRESUPUESTO.", vbExclamation
        End If

        wsPresupuesto.Protect password:=claveOtrasHojas, AllowFormattingRows:=True, AllowFormattingColumns:=True
        wsPresupuesto.Visible = xlSheetHidden
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveOtrasHojas, Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

Fin:
    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wsProductos = Nothing
    Set wsAplicabilidad = Nothing
    Set wsPreciosAdjudicados = Nothing
    Set wsPreciosActualizados = Nothing
    Set wsPreciosProformas = Nothing
    Set wsPresupuesto = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    ThisWorkbook.Protect password:=claveOtrasHojas, Structure:=True
    GoTo Fin
End Sub


