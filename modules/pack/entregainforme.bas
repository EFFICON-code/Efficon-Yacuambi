Attribute VB_Name = "entregainforme"
Sub Entrega_de_Informe()
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

    ' Variables para datos a insertar
    Dim siglasUnidad As String
    Dim siglas As String
    Dim periodo As String
    Dim lugar As String
    Dim fecha As String
    Dim comprasPublicas As String
    Dim cargoComprasPublicas As String
    Dim titularUnidadRequirente As String
    Dim cargoTitularUnidadRequirente As String
    Dim memorandoSolicitudInforme As String
    Dim fechaSolicitudInforme As String
    Dim objetoContratacion As String
    Dim nroInforme As String

    ' Clave para la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave general para la estructura del libro y otras hojas
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B143
    plantillaID = wsBase.range("B143").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B143 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("Informe_Terminado.docx", _
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
    siglasUnidad = CStr(ws.range("DB2").Value)
    siglas = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    comprasPublicas = CStr(ws.range("I2").Value)
    cargoComprasPublicas = CStr(ws.range("J2").Value)
    titularUnidadRequirente = CStr(ws.range("E2").Value)
    cargoTitularUnidadRequirente = CStr(ws.range("F2").Value)
    memorandoSolicitudInforme = CStr(ws.range("DR2").Value)
    fechaSolicitudInforme = CStr(ws.range("DS2").Value)
    objetoContratacion = CStr(ws.range("Q2").Value)
    nroInforme = CStr(ws.range("EJ2").Value)

    ' Proteger y ocultar la hoja nuevamente
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Ruta temporal para descargar la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_EntregaInforme_Temp.docx"

    ' Descargar la plantilla con MSXML2.ServerXMLHTTP
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
        If .Bookmarks.Exists("Siglas_Unidad") Then .Bookmarks("Siglas_Unidad").range.Text = siglasUnidad
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Compras_Publicas") Then .Bookmarks("Compras_Publicas").range.Text = comprasPublicas
        If .Bookmarks.Exists("Cargo_Compras_Publicas") Then .Bookmarks("Cargo_Compras_Publicas").range.Text = cargoComprasPublicas
        If .Bookmarks.Exists("Titular_Unidad_Requirente") Then .Bookmarks("Titular_Unidad_Requirente").range.Text = titularUnidadRequirente
        If .Bookmarks.Exists("Cargo_Titular_Unidad_Requirente") Then .Bookmarks("Cargo_Titular_Unidad_Requirente").range.Text = cargoTitularUnidadRequirente
        If .Bookmarks.Exists("Memorando_Solicitud_Informe") Then .Bookmarks("Memorando_Solicitud_Informe").range.Text = memorandoSolicitudInforme
        If .Bookmarks.Exists("Fecha_Solicitud_Informe") Then .Bookmarks("Fecha_Solicitud_Informe").range.Text = fechaSolicitudInforme
        If .Bookmarks.Exists("Objeto_Contratacion") Then .Bookmarks("Objeto_Contratacion").range.Text = objetoContratacion
        If .Bookmarks.Exists("Nro_Informe") Then .Bookmarks("Nro_Informe").range.Text = nroInforme
        If .Bookmarks.Exists("Titular_Unidad_Requirente1") Then .Bookmarks("Titular_Unidad_Requirente1").range.Text = titularUnidadRequirente
        If .Bookmarks.Exists("Cargo_Titular_Unidad_Requirente1") Then .Bookmarks("Cargo_Titular_Unidad_Requirente1").range.Text = cargoTitularUnidadRequirente
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "CUADRO-INF"
    ThisWorkbook.Sheets("CUADRO-INF").Activate

    ' Eliminar el archivo temporal
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


