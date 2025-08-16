Attribute VB_Name = "autorizacionpublicacion"
Sub Autorizacion_de_Publicacion()
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
    
    Dim lugar As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim tecnicoRequirente As String
    Dim cargoTecnico As String
    Dim objetoDeContratacion As String
    Dim firmaTecnico As String
    Dim cargoTecnico1 As String
    Dim fecha As String
    Dim siglaEntidad As String
    Dim periodo As String
    Dim administrativo As String
    Dim cargoAdministrativo As String
    
    ' Claves
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral
    
    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral
    
    ' Leer el ID de la plantilla desde la celda B141
    plantillaID = wsBase.range("B141").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda B141 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If
    
    ' Construir la URL de la plantilla (p.e. Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID
    
    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral
    
    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
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
    lugar = CStr(ws.range("FQ2").Value)
    presidente = CStr(ws.range("I2").Value)
    cargoPresidente = CStr(ws.range("J2").Value)
    tecnicoRequirente = CStr(ws.range("K2").Value)
    cargoTecnico = CStr(ws.range("L2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    firmaTecnico = CStr(ws.range("E2").Value)
    cargoTecnico1 = CStr(ws.range("F2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    siglaEntidad = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    administrativo = CStr(ws.range("K2").Value)
    cargoAdministrativo = CStr(ws.range("L2").Value)

    ' Proteger y ocultar la hoja nuevamente
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal para la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_AutPublicacion_Temp.docx"

    ' Descargar la plantilla usando MSXML2.ServerXMLHTTP
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
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Tecnico_requirente") Then .Bookmarks("Tecnico_requirente").range.Text = tecnicoRequirente
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Firma_tecnico") Then .Bookmarks("Firma_tecnico").range.Text = firmaTecnico
        If .Bookmarks.Exists("Cargo_Tecnico1") Then .Bookmarks("Cargo_Tecnico1").range.Text = cargoTecnico1
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglaEntidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Administrativo") Then .Bookmarks("Administrativo").range.Text = administrativo
        If .Bookmarks.Exists("Cargo_administrativo") Then .Bookmarks("Cargo_administrativo").range.Text = cargoAdministrativo
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
