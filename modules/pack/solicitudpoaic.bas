Attribute VB_Name = "solicitudpoaic"
Sub Solicitud_POA_IC()
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

    ' Variables para almacenar datos a insertar en Word
    Dim siglas As String
    Dim periodo As String
    Dim lugar As String
    Dim fecha As String
    Dim comprasPublicas As String
    Dim cargoComprasPublicas As String
    Dim responsablePOA As String
    Dim cargoResponsablePOA As String
    Dim entidad As String
    Dim comprasPublicas1 As String
    Dim cargoComprasPublicas1 As String
    Dim objetoDeContratacion As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave general
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda B145
    plantillaID = wsBase.range("B145").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B145 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("SolicitudPOA_IC_Terminado.docx", _
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
    siglas = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    comprasPublicas = CStr(ws.range("I2").Value)
    cargoComprasPublicas = CStr(ws.range("J2").Value)
    responsablePOA = CStr(ws.range("CF2").Value)
    cargoResponsablePOA = CStr(ws.range("CG2").Value)
    entidad = CStr(ws.range("A2").Value)
    comprasPublicas1 = CStr(ws.range("I2").Value)
    cargoComprasPublicas1 = CStr(ws.range("J2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)

    ' Proteger y ocultar nuevamente la hoja "SECUENCIAS"
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Descargar la plantilla a una ruta temporal
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_SolicitudPOAIC_Temp.docx"
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
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglas
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Compras_Publicas") Then .Bookmarks("Compras_Publicas").range.Text = comprasPublicas
        If .Bookmarks.Exists("Cargo_Compras_Publicas") Then .Bookmarks("Cargo_Compras_Publicas").range.Text = cargoComprasPublicas
        If .Bookmarks.Exists("Responsable_POA") Then .Bookmarks("Responsable_POA").range.Text = responsablePOA
        If .Bookmarks.Exists("Cargo_Responsable_POA") Then .Bookmarks("Cargo_Responsable_POA").range.Text = cargoResponsablePOA
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Compras_Publicas1") Then .Bookmarks("Compras_Publicas1").range.Text = comprasPublicas1
        If .Bookmarks.Exists("Cargo_Compras_Publicas1") Then .Bookmarks("Cargo_Compras_Publicas1").range.Text = cargoComprasPublicas1
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

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


