Attribute VB_Name = "solicitudinicio"
Sub Solicitud_de_Autorizacion_Inicio()
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
    
    ' Variables para datos en marcadores
    Dim siglas As String
    Dim lugar As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim objetoDeContratacion As String
    Dim unidadRequirente As String
    Dim tipoDeProcedimiento As String
    Dim unidadRequirente1 As String
    Dim objetoDeContratacion1 As String
    Dim presupuesto As String
    Dim valorLetras As String
    Dim tecnicoRequirente As String
    Dim cargoTecnico As String
    Dim fecha As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    Const claveSecuencias As String = "Admin1991"
    ' Clave para la estructura y otras hojas
    Const claveGeneral As String = "PROEST2023"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:=claveGeneral

    ' Asignar y desproteger la hoja "BBDD"
    Set wsBase = ThisWorkbook.Sheets("BBDD")
    wsBase.Unprotect password:=claveGeneral

    ' Leer el ID de la plantilla desde la celda D147
    plantillaID = wsBase.range("D147").Value
    If plantillaID = "" Then
        MsgBox "No se encontró el ID de la plantilla en la celda D147 de la hoja BBDD.", vbExclamation
        Exit Sub
    End If

    ' Construir la URL de la plantilla (por ejemplo, Google Drive)
    plantillaRuta = "https://drive.google.com/uc?export=download&id=" & plantillaID

    ' Proteger nuevamente la hoja "BBDD"
    wsBase.Protect password:=claveGeneral

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename( _
        "Solicitud_Autorizacion_Inicio_Terminado.docx", _
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
    presidente = CStr(ws.range("B2").Value)
    cargoPresidente = CStr(ws.range("C2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    unidadRequirente = CStr(ws.range("DA2").Value)
    tipoDeProcedimiento = CStr(ws.range("S2").Value)
    unidadRequirente1 = CStr(ws.range("DA2").Value)
    objetoDeContratacion1 = CStr(ws.range("Q2").Value)
    presupuesto = CStr(ws.range("BV2").Value)
    valorLetras = CStr(ws.range("BW2").Value)
    tecnicoRequirente = CStr(ws.range("E2").Value)
    cargoTecnico = CStr(ws.range("F2").Value)
    fecha = CStr(ws.range("GZ2").Value)

    ' Proteger y ocultar la hoja "SECUENCIAS" nuevamente
    ws.Protect password:=claveSecuencias, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Construir la ruta temporal donde se descargará la plantilla
    rutaDescargaTemporal = Environ("TEMP") & "\Plantilla_SolicitudAutorizacionInicio_Temp.docx"

    ' Descargar la plantilla usando MSXML2.ServerXMLHTTP
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
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Unidad_Requirente") Then .Bookmarks("Unidad_Requirente").range.Text = unidadRequirente
        If .Bookmarks.Exists("Tipo_de_procedimiento") Then .Bookmarks("Tipo_de_procedimiento").range.Text = tipoDeProcedimiento
        If .Bookmarks.Exists("Unidad_Requirente1") Then .Bookmarks("Unidad_Requirente1").range.Text = unidadRequirente1
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Presupuesto") Then .Bookmarks("Presupuesto").range.Text = presupuesto
        If .Bookmarks.Exists("Valor_letras") Then .Bookmarks("Valor_letras").range.Text = valorLetras
        If .Bookmarks.Exists("Tecnico_requirente") Then .Bookmarks("Tecnico_requirente").range.Text = tecnicoRequirente
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
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


