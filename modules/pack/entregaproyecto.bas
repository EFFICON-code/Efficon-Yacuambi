Attribute VB_Name = "entregaproyecto"
Sub Entrega_de_Proyecto()
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
    
    ' Claves para desproteger/proteger
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Variables para la inserción en marcadores
    Dim fecha As String
    Dim lugar As String
    Dim siglas As String
    Dim siglasUnidad As String
    Dim periodo As String
    Dim titularUnidadRequirente As String
    Dim cargoTitularUnidadRequirente As String
    Dim objetoContratacion As String
    Dim memorandoSolicitudEstudios As String
    Dim nombreTecnico As String
    Dim cargoTecnico As String
    Dim tituloTDR As String

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D142
    plantillaID = wsBase.range("D142").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D142 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para la ubicación donde se guardará el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "Entrega_Proyecto_Terminado.docx", _
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

    ' Leer datos de Excel desde "SECUENCIAS"
    fecha = CStr(ws.range("GZ2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    siglas = CStr(ws.range("HA2").Value)
    siglasUnidad = CStr(ws.range("DB2").Value)
    periodo = CStr(ws.range("HB2").Value)
    titularUnidadRequirente = CStr(ws.range("E2").Value)
    cargoTitularUnidadRequirente = CStr(ws.range("F2").Value)
    objetoContratacion = CStr(ws.range("Q2").Value)
    memorandoSolicitudEstudios = CStr(ws.range("EF2").Value)
    nombreTecnico = CStr(ws.range("G2").Value)
    cargoTecnico = CStr(ws.range("H2").Value)
    tituloTDR = CStr(ws.range("AO2").Value)

    ' Proteger y ocultar la hoja "SECUENCIAS"
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal para la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_EntregaProyecto_Temp.docx"

    ' Descargar la plantilla usando MSXML2.ServerXMLHTTP
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
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Siglas_Unidad") Then .Bookmarks("Siglas_Unidad").range.Text = siglasUnidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Titular_Unidad_Requirente") Then .Bookmarks("Titular_Unidad_Requirente").range.Text = titularUnidadRequirente
        If .Bookmarks.Exists("Cargo_Titular_Unidad_Requirente") Then .Bookmarks("Cargo_Titular_Unidad_Requirente").range.Text = cargoTitularUnidadRequirente
        If .Bookmarks.Exists("Objeto_Contratacion") Then .Bookmarks("Objeto_Contratacion").range.Text = objetoContratacion
        If .Bookmarks.Exists("Memorando_Solicitud_Estudios") Then .Bookmarks("Memorando_Solicitud_Estudios").range.Text = memorandoSolicitudEstudios
        If .Bookmarks.Exists("Objeto_Contratacion1") Then .Bookmarks("Objeto_Contratacion1").range.Text = objetoContratacion
        If .Bookmarks.Exists("Nombre_Tecnico") Then .Bookmarks("Nombre_Tecnico").range.Text = nombreTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Titulo_TDR") Then .Bookmarks("Titulo_TDR").range.Text = tituloTDR
    End With

    ' Guardar y cerrar documento
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
    Set wsBase = Nothing

    MsgBox "El documento se ha generado correctamente.", vbInformation
End Sub


