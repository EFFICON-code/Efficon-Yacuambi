Attribute VB_Name = "requerimiento"
Sub Requerimiento()
    Dim wdApp As Object
    Dim wdDoc As Object
    
    Dim objHTTP As Object
    Dim objStream As Object
    Dim rutaDescargaTemporal As String
    
    Dim plantillaID As String
    Dim plantillaRuta As String
    Dim guardarRuta As Variant
    
    Dim ws As Worksheet, wsProductos As Worksheet, wsBase As Worksheet
    
    Dim CLAVE As String
    Dim nombreTecnicoUnidad As String
    Dim cargoTecnicoUnidad As String
    Dim nombreTecnicoUnidad1 As String
    Dim cargoTecnicoUnidad1 As String
    Dim nroRequerimiento As String
    Dim nombreTitularUnidad As String
    Dim cargoTitularUnidad As String
    Dim fechaRequerimiento As String
    Dim objetoDeContratacion As String
    Dim formaDePago As String
    Dim garantia As String
    Dim justificacionNecesidad As String
    Dim plazoDeEntrega As String
    Dim tipoDeCompra As String
    Dim unidadRequirente As String
    Dim nombreUnidadRequirente As String
    
    ' Clave para la hoja "SECUENCIAS"
    CLAVE = "Admin1991"

    ' Desproteger la estructura del libro (clave general)
    ThisWorkbook.Unprotect password:="PROEST2023"
    
    ' Asignar la hoja "BBDD" y desprotegerla
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:="PROEST2023"
    
    ' Leer el ID de la plantilla desde la celda B133 de la hoja "BBDD"
    plantillaID = wsBase.range("B133").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B133 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de descarga
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID
    
    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:="PROEST2023"

    ' Mostrar cuadro de diálogo para la ubicación donde se guardará el documento terminado
    guardarRuta = Application.GetSaveAsFilename("DocumentoTerminado.docx", _
        "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    ' Asignar la hoja de trabajo "SECUENCIAS" y desprotegerla
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    ws.Visible = xlSheetVisible
    ws.Unprotect password:=CLAVE

    ' Leer datos de la hoja "SECUENCIAS"
    nombreTecnicoUnidad = ws.range("G2").Value
    cargoTecnicoUnidad = ws.range("H2").Value
    nombreTecnicoUnidad1 = ws.range("G2").Value
    cargoTecnicoUnidad1 = ws.range("H2").Value
    nroRequerimiento = ws.range("M2").Value
    nombreTitularUnidad = ws.range("E2").Value
    cargoTitularUnidad = ws.range("F2").Value
    fechaRequerimiento = ws.range("N2").Value
    objetoDeContratacion = ws.range("Q2").Value
    formaDePago = ws.range("AS2").Value
    garantia = ws.range("U2").Value
    justificacionNecesidad = ws.range("AF2").Value
    plazoDeEntrega = ws.range("T2").Value
    tipoDeCompra = ws.range("O2").Value
    unidadRequirente = ws.range("D2").Value
    nombreUnidadRequirente = ws.range("DA2").Value

    ' Proteger y ocultar nuevamente la hoja "SECUENCIAS"
    ws.Protect password:=CLAVE
    ws.Visible = xlSheetHidden
    
    ' Construir la ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_Requerimiento_Temp.docx"
    
    ' Descargar la plantilla con MSXML2.ServerXMLHTTP
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.Status = 200 Then
        ' Guardar el archivo descargado en la ubicación temporal
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Tipo binario
        objStream.Open
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile rutaDescargaTemporal, 2 ' Sobrescribe si existe
        objStream.Close
    Else
        MsgBox "Error al descargar la plantilla. Verifique la conexión o el enlace." & vbCrLf & _
               "Código de estado: " & objHTTP.Status & " - " & objHTTP.statusText, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
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

    ' Insertar datos en el documento de Word
    With wdDoc
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad") Then .Bookmarks("Nombre_Tecnico_Unidad").range.Text = nombreTecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad") Then .Bookmarks("Cargo_Tecnico_Unidad").range.Text = cargoTecnicoUnidad
        If .Bookmarks.Exists("Nombre_Tecnico_Unidad1") Then .Bookmarks("Nombre_Tecnico_Unidad1").range.Text = nombreTecnicoUnidad1
        If .Bookmarks.Exists("Cargo_Tecnico_Unidad1") Then .Bookmarks("Cargo_Tecnico_Unidad1").range.Text = cargoTecnicoUnidad1
        If .Bookmarks.Exists("Nro_Requerimiento") Then .Bookmarks("Nro_Requerimiento").range.Text = nroRequerimiento
        If .Bookmarks.Exists("Nombre_Titular_Unidad") Then .Bookmarks("Nombre_Titular_Unidad").range.Text = nombreTitularUnidad
        If .Bookmarks.Exists("Cargo_Titular_Unidad") Then .Bookmarks("Cargo_Titular_Unidad").range.Text = cargoTitularUnidad
        If .Bookmarks.Exists("Fecha_Requerimiento") Then .Bookmarks("Fecha_Requerimiento").range.Text = fechaRequerimiento
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Forma_de_Pago") Then .Bookmarks("Forma_de_Pago").range.Text = formaDePago
        If .Bookmarks.Exists("Garantia") Then .Bookmarks("Garantia").range.Text = garantia
        If .Bookmarks.Exists("Justificacion_Necesidad") Then .Bookmarks("Justificacion_Necesidad").range.Text = justificacionNecesidad
        If .Bookmarks.Exists("Plazo_de_Entrega") Then .Bookmarks("Plazo_de_Entrega").range.Text = plazoDeEntrega
        If .Bookmarks.Exists("Tipo_de_Compra") Then .Bookmarks("Tipo_de_Compra").range.Text = tipoDeCompra
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Nombre_Unidad_Requirente") Then .Bookmarks("Nombre_Unidad_Requirente").range.Text = nombreUnidadRequirente

        ' Transferir datos desde la hoja "PRODUCTOS"
        Set wsProductos = ThisWorkbook.Sheets("PRODUCTOS")
        wsProductos.Unprotect password:="PROEST2023"
        
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
        
        wsProductos.Protect password:="PROEST2023", Scenarios:=True, AllowFormattingRows:=True
    End With

    ' Guardar y cerrar el documento de Word
    wdDoc.SaveAs2 Filename:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Ubicarse en la hoja "REQUERIMIENTO"
    ThisWorkbook.Sheets("REQUERIMIENTO").Activate

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:="PROEST2023", Structure:=True

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




