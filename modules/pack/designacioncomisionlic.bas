Attribute VB_Name = "designacioncomisionlic"
Sub Designacion_Comision_Licitacion()
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

    ' Claves de protección
    Const claveSecuencias As String = "Admin1991"
    Const claveGeneral As String = "PROEST2023"

    ' Variables para marcadores
    Dim lugar As String
    Dim delegado1 As String, cargoDelegado1 As String
    Dim delegado2 As String, cargoDelegado2 As String
    Dim delegado3 As String, cargoDelegado3 As String
    Dim delegado4 As String, cargoDelegado4 As String
    Dim delegado5 As String, cargoDelegado5 As String
    Dim delegado11 As String, cargoDelegado11 As String
    Dim cedula1 As String, funcion1 As String
    Dim delegado22 As String, cargoDelegado22 As String
    Dim cedula2 As String, funcion2 As String
    Dim delegado33 As String, cargoDelegado33 As String
    Dim cedula3 As String, funcion3 As String
    Dim delegado44 As String, cargoDelegado44 As String
    Dim cedula4 As String, funcion4 As String
    Dim delegado55 As String, cargoDelegado55 As String
    Dim cedula5 As String, funcion5 As String
    Dim tipoDeProcedimiento As String
    Dim objetoDeContratacion As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim fecha As String

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D150
    plantillaID = wsBase.range("D150").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D150 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Solicitar la ubicación donde se guardará el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "Designacion_Comision_Licitacion_Terminado.docx", _
        "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = False Or guardarRuta = "" Then
        MsgBox "Operación cancelada por el usuario.", vbInformation
        Exit Sub
    End If

    ' Asignar la hoja de trabajo "SECUENCIAS" y desprotegerla
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=claveSecuencias

    ' Leer datos de Excel
    lugar = CStr(ws.range("FQ2").Value)
    delegado1 = CStr(ws.range("FR2").Value)
    cargoDelegado1 = CStr(ws.range("GB2").Value)
    delegado2 = CStr(ws.range("FS2").Value)
    cargoDelegado2 = CStr(ws.range("GC2").Value)
    delegado3 = CStr(ws.range("FT2").Value)
    cargoDelegado3 = CStr(ws.range("GD2").Value)
    delegado4 = CStr(ws.range("FU2").Value)
    cargoDelegado4 = CStr(ws.range("GE2").Value)
    delegado5 = CStr(ws.range("FV2").Value)
    cargoDelegado5 = CStr(ws.range("GF2").Value)
    delegado11 = CStr(ws.range("FR2").Value)
    cargoDelegado11 = CStr(ws.range("GB2").Value)
    cedula1 = CStr(ws.range("FW2").Value)
    funcion1 = CStr(ws.range("GL2").Value)
    delegado22 = CStr(ws.range("FS2").Value)
    cargoDelegado22 = CStr(ws.range("GC2").Value)
    cedula2 = CStr(ws.range("FX2").Value)
    funcion2 = CStr(ws.range("GM2").Value)
    delegado33 = CStr(ws.range("FT2").Value)
    cargoDelegado33 = CStr(ws.range("GD2").Value)
    cedula3 = CStr(ws.range("FY2").Value)
    funcion3 = CStr(ws.range("GN2").Value)
    delegado44 = CStr(ws.range("FU2").Value)
    cargoDelegado44 = CStr(ws.range("GE2").Value)
    cedula4 = CStr(ws.range("FZ2").Value)
    funcion4 = CStr(ws.range("GP2").Value)
    delegado55 = CStr(ws.range("FV2").Value)
    cargoDelegado55 = CStr(ws.range("GF2").Value)
    cedula5 = CStr(ws.range("GA2").Value)
    funcion5 = CStr(ws.range("GQ2").Value)
    tipoDeProcedimiento = CStr(ws.range("S2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    presidente = CStr(ws.range("B2").Value)
    cargoPresidente = CStr(ws.range("C2").Value)
    fecha = CStr(ws.range("GZ2").Value)

    ' Proteger y ocultar la hoja nuevamente permitiendo modificar escenarios
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_DesignacionComisionLicitacion_Temp.docx"

    ' Descargar la plantilla
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "GET", plantillaRuta, False
    objHTTP.send

    If objHTTP.status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binario
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
        If .Bookmarks.Exists("Delegado1") Then .Bookmarks("Delegado1").range.Text = delegado1
        If .Bookmarks.Exists("Cargo_delegado1") Then .Bookmarks("Cargo_delegado1").range.Text = cargoDelegado1
        If .Bookmarks.Exists("Delegado2") Then .Bookmarks("Delegado2").range.Text = delegado2
        If .Bookmarks.Exists("Cargo_delegado2") Then .Bookmarks("Cargo_delegado2").range.Text = cargoDelegado2
        If .Bookmarks.Exists("Delegado3") Then .Bookmarks("Delegado3").range.Text = delegado3
        If .Bookmarks.Exists("Cargo_delegado3") Then .Bookmarks("Cargo_delegado3").range.Text = cargoDelegado3
        If .Bookmarks.Exists("Delegado4") Then .Bookmarks("Delegado4").range.Text = delegado4
        If .Bookmarks.Exists("Cargo_delegado4") Then .Bookmarks("Cargo_delegado4").range.Text = cargoDelegado4
        If .Bookmarks.Exists("Delegado5") Then .Bookmarks("Delegado5").range.Text = delegado5
        If .Bookmarks.Exists("Cargo_delegado5") Then .Bookmarks("Cargo_delegado5").range.Text = cargoDelegado5
        If .Bookmarks.Exists("Delegado11") Then .Bookmarks("Delegado11").range.Text = delegado11
        If .Bookmarks.Exists("Cargo_delegado11") Then .Bookmarks("Cargo_delegado11").range.Text = cargoDelegado11
        If .Bookmarks.Exists("Cedula1") Then .Bookmarks("Cedula1").range.Text = cedula1
        If .Bookmarks.Exists("Funcion1") Then .Bookmarks("Funcion1").range.Text = funcion1
        If .Bookmarks.Exists("Delegado22") Then .Bookmarks("Delegado22").range.Text = delegado22
        If .Bookmarks.Exists("Cargo_delegado22") Then .Bookmarks("Cargo_delegado22").range.Text = cargoDelegado22
        If .Bookmarks.Exists("Cedula2") Then .Bookmarks("Cedula2").range.Text = cedula2
        If .Bookmarks.Exists("Funcion2") Then .Bookmarks("Funcion2").range.Text = funcion2
        If .Bookmarks.Exists("Delegado33") Then .Bookmarks("Delegado33").range.Text = delegado33
        If .Bookmarks.Exists("Cargo_delegado33") Then .Bookmarks("Cargo_delegado33").range.Text = cargoDelegado33
        If .Bookmarks.Exists("Cedula3") Then .Bookmarks("Cedula3").range.Text = cedula3
        If .Bookmarks.Exists("Funcion3") Then .Bookmarks("Funcion3").range.Text = funcion3
        If .Bookmarks.Exists("Delegado44") Then .Bookmarks("Delegado44").range.Text = delegado44
        If .Bookmarks.Exists("Cargo_delegado44") Then .Bookmarks("Cargo_delegado44").range.Text = cargoDelegado44
        If .Bookmarks.Exists("Cedula4") Then .Bookmarks("Cedula4").range.Text = cedula4
        If .Bookmarks.Exists("Funcion4") Then .Bookmarks("Funcion4").range.Text = funcion4
        If .Bookmarks.Exists("Delegado55") Then .Bookmarks("Delegado55").range.Text = delegado55
        If .Bookmarks.Exists("Cargo_delegado55") Then .Bookmarks("Cargo_delegado55").range.Text = cargoDelegado55
        If .Bookmarks.Exists("Cedula5") Then .Bookmarks("Cedula5").range.Text = cedula5
        If .Bookmarks.Exists("Funcion5") Then .Bookmarks("Funcion5").range.Text = funcion5
        If .Bookmarks.Exists("Tipo_de_procedimiento") Then .Bookmarks("Tipo_de_procedimiento").range.Text = tipoDeProcedimiento
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
    End With

    ' Guardar y cerrar el documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:=claveGeneral, Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

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

