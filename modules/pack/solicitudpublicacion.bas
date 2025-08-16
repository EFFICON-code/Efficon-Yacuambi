Attribute VB_Name = "solicitudpublicacion"
Sub Solicitud_de_Publicacion()
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

    Dim siglas As String
    Dim lugar As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim objetoDeContratacion As String
    Dim firmaTecnico As String
    Dim cargoTecnico As String
    Dim fecha As String
    Dim siglaEntidad As String
    Dim periodo As String

    ' Claves para desproteger/proteger
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar la hoja "BBDD" y desprotegerla
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B140
    plantillaID = wsBase.range("B140").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B140 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para la ubicación donde se guardará el documento terminado
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
    siglas = CStr(ws.range("DB2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    presidente = CStr(ws.range("K2").Value)
    cargoPresidente = CStr(ws.range("L2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    firmaTecnico = CStr(ws.range("E2").Value)
    cargoTecnico = CStr(ws.range("F2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    siglaEntidad = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)

    ' Proteger y ocultar nuevamente la hoja SECUENCIAS
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_SolicituddePublicacion_Temp.docx"

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
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Firma_Tecnico") Then .Bookmarks("Firma_Tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglaEntidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
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


