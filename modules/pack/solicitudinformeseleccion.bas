Attribute VB_Name = "solicitudinformeseleccion"
Sub Solicitud_de_Informe_Seleccion()
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
    
    ' Variables para leer datos de Excel
    Dim siglas As String
    Dim periodo As String
    Dim lugar As String
    Dim fecha As String
    Dim titularUnidadRequirente As String
    Dim cargoTitularUnidadRequirente As String
    Dim comprasPublicas As String
    Dim cargoComprasPublicas As String
    Dim objetoContratacion As String
    Dim memorandoDisposicionPublicacion As String
    Dim fechaDisposicionPublicacion As String
    Dim administrativo As String
    Dim publicacion As String
    Dim codigoNecesidad As String
    Dim fechaRecepcionProformas As String
    Dim nroProformas As String

    ' Clave para la hoja SECUENCIAS
    Const claveSecuencias As String = "Admin1991"
    ' Clave general para la estructura y otras hojas
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B142
    plantillaID = wsBase.range("B142").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B142 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "Entrega_Proformas_Terminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
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
    titularUnidadRequirente = CStr(ws.range("E2").Value)
    cargoTitularUnidadRequirente = CStr(ws.range("F2").Value)
    comprasPublicas = CStr(ws.range("I2").Value)
    cargoComprasPublicas = CStr(ws.range("J2").Value)
    objetoContratacion = CStr(ws.range("Q2").Value)
    memorandoDisposicionPublicacion = CStr(ws.range("DM2").Value)
    fechaDisposicionPublicacion = CStr(ws.range("DN2").Value)
    administrativo = CStr(ws.range("K2").Value)
    publicacion = CStr(ws.range("DQ2").Value)
    codigoNecesidad = CStr(ws.range("DP2").Value)
    fechaRecepcionProformas = CStr(ws.range("EH2").Value)
    nroProformas = CStr(ws.range("EI2").Value)

    ' Proteger y ocultar la hoja nuevamente
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal para la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_EntregaProformas_Temp.docx"

    ' Descargar la plantilla
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
        If .Bookmarks.Exists("Titular_Unidad_Requirente") Then .Bookmarks("Titular_Unidad_Requirente").range.Text = titularUnidadRequirente
        If .Bookmarks.Exists("Cargo_Titular_Unidad_Requirente") Then .Bookmarks("Cargo_Titular_Unidad_Requirente").range.Text = cargoTitularUnidadRequirente
        If .Bookmarks.Exists("Compras_Publicas") Then .Bookmarks("Compras_Publicas").range.Text = comprasPublicas
        If .Bookmarks.Exists("Cargo_Compras_Publicas") Then .Bookmarks("Cargo_Compras_Publicas").range.Text = cargoComprasPublicas
        If .Bookmarks.Exists("Objeto_Contratacion") Then .Bookmarks("Objeto_Contratacion").range.Text = objetoContratacion
        If .Bookmarks.Exists("Memorando_Disposicion_Publicacion") Then .Bookmarks("Memorando_Disposicion_Publicacion").range.Text = memorandoDisposicionPublicacion
        If .Bookmarks.Exists("Fecha_Disposicion_Publicacion") Then .Bookmarks("Fecha_Disposicion_Publicacion").range.Text = fechaDisposicionPublicacion
        If .Bookmarks.Exists("Administrativo") Then .Bookmarks("Administrativo").range.Text = administrativo
        If .Bookmarks.Exists("Publicacion") Then .Bookmarks("Publicacion").range.Text = publicacion
        If .Bookmarks.Exists("Codigo_Necesidad") Then .Bookmarks("Codigo_Necesidad").range.Text = codigoNecesidad
        If .Bookmarks.Exists("Fecha_Recepcion_proformas") Then .Bookmarks("Fecha_Recepcion_proformas").range.Text = fechaRecepcionProformas
        If .Bookmarks.Exists("Nro_proformas") Then .Bookmarks("Nro_proformas").range.Text = nroProformas
        If .Bookmarks.Exists("Compras_Publicas1") Then .Bookmarks("Compras_Publicas1").range.Text = comprasPublicas
        If .Bookmarks.Exists("Cargo_Compras_Publicas1") Then .Bookmarks("Cargo_Compras_Publicas1").range.Text = cargoComprasPublicas
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET'S-TDR"
    ThisWorkbook.Sheets("ET'S-TDR").Activate

    ' Eliminar archivo temporal
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


