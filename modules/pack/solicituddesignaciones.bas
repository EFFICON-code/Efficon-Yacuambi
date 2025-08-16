Attribute VB_Name = "solicituddesignaciones"
Sub Solicitud_de_Designaciones()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim plantillaRuta As Variant
    Dim guardarRuta As Variant
    Dim ws As Worksheet
    Dim siglas As String
    Dim lugar As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim tipoDeProcedimiento As String
    Dim objetoDeContratacion As String
    Dim designacion As String
    Dim tecnicoRequirente As String
    Dim cargoTecnico As String
    Dim fecha As String
    Dim CLAVE As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    CLAVE = "Admin1991"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:="PROEST2023"

    ' Mostrar cuadro de diálogo para seleccionar la plantilla de Word
    plantillaRuta = Application.GetOpenFilename("Archivos de Word (*.docx), *.docx", , "Seleccionar plantilla de Word")
    If plantillaRuta = "False" Then Exit Sub ' Si el usuario cancela la selección, salir de la macro

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("Solicitud_Designaciones_Terminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = "False" Then Exit Sub ' Si el usuario cancela la selección, salir de la macro

    ' Asignar la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")

    ' Mostrar y desproteger la hoja si está oculta y protegida
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=CLAVE

    ' Leer datos de Excel
    siglas = CStr(ws.range("DB2").Value)
    lugar = CStr(ws.range("FQ2").Value)
    presidente = CStr(ws.range("B2").Value)
    cargoPresidente = CStr(ws.range("C2").Value)
    tipoDeProcedimiento = CStr(ws.range("S2").Value)
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    designacion = CStr(ws.range("CC2").Value)
    tecnicoRequirente = CStr(ws.range("I2").Value)
    cargoTecnico = CStr(ws.range("J2").Value)
    fecha = CStr(ws.range("GZ2").Value)

    ' Proteger y ocultar la hoja nuevamente permitiendo modificar escenarios
    ws.Protect password:=CLAVE, Scenarios:=True
    ws.Visible = xlSheetHidden

    ' Iniciar Word y abrir la plantilla
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject(Class:="Word.Application")
    End If
    On Error GoTo 0

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(plantillaRuta)

    ' Insertar datos en los marcadores de la plantilla
    With wdDoc
        If .Bookmarks.Exists("Siglas") Then .Bookmarks("Siglas").range.Text = siglas
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Tipo_de_procedimiento") Then .Bookmarks("Tipo_de_procedimiento").range.Text = tipoDeProcedimiento
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Designación") Then .Bookmarks("Designación").range.Text = designacion
        If .Bookmarks.Exists("Tecnico_requirente") Then .Bookmarks("Tecnico_requirente").range.Text = tecnicoRequirente
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:="PROEST2023", Structure:=True

    ' Ubicarse en la hoja "ET-REFPAC-INF-CONSULT"
    ThisWorkbook.Sheets("ET-REFPAC-INF-CONSULT").Activate

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing

End Sub


