Attribute VB_Name = "solicitudadjudicacion"
Sub Solicitud_Adjudicacion()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim plantillaRuta As Variant
    Dim guardarRuta As Variant
    Dim ws As Worksheet
    Dim siglas As String
    Dim lugar As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim objetoDeContratacion As String
    Dim nroCertificacionPresupuesto As String
    Dim fechaCertificacion As String
    Dim objetoDeContratacion1 As String
    Dim presupuesto As String
    Dim valorLetras As String
    Dim cuadroComparativo As String
    Dim proveedor As String
    Dim ruc As String
    Dim objetoDeContratacion2 As String
    Dim tecnicoRequirente As String
    Dim cargoTecnico As String
    Dim fecha As String
    Dim siglaEntidad As String
    Dim periodo As String
    Dim entidad As String
    Dim CLAVE As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    CLAVE = "Admin1991"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:="PROEST2023"

    ' Mostrar cuadro de diálogo para seleccionar la plantilla de Word
    plantillaRuta = Application.GetOpenFilename("Archivos de Word (*.docx), *.docx", , "Seleccionar plantilla de Word")
    If plantillaRuta = "False" Then Exit Sub ' Si el usuario cancela la selección, salir de la macro

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("SolicitudAdjudicacion_Terminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
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
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    nroCertificacionPresupuesto = CStr(ws.range("DR2").Value)
    fechaCertificacion = CStr(ws.range("DS2").Value)
    objetoDeContratacion1 = CStr(ws.range("Q2").Value)
    presupuesto = CStr(ws.range("DC2").Value)
    valorLetras = CStr(ws.range("DD2").Value)
    cuadroComparativo = CStr(ws.range("DM2").Value)
    proveedor = CStr(ws.range("DE2").Value)
    ruc = CStr(ws.range("DF2").Value)
    objetoDeContratacion2 = CStr(ws.range("Q2").Value)
    tecnicoRequirente = CStr(ws.range("H2").Value)
    cargoTecnico = CStr(ws.range("G2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    siglaEntidad = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    entidad = CStr(ws.range("A2").Value)

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
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Nro_Certificacion_Presupuesto") Then .Bookmarks("Nro_Certificacion_Presupuesto").range.Text = nroCertificacionPresupuesto
        If .Bookmarks.Exists("Fecha_Certificacion") Then .Bookmarks("Fecha_Certificacion").range.Text = fechaCertificacion
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Presupuesto") Then .Bookmarks("Presupuesto").range.Text = presupuesto
        If .Bookmarks.Exists("Valor_letras") Then .Bookmarks("Valor_letras").range.Text = valorLetras
        If .Bookmarks.Exists("Cuadro_Comparativo") Then .Bookmarks("Cuadro_Comparativo").range.Text = cuadroComparativo
        If .Bookmarks.Exists("Proveedor") Then .Bookmarks("Proveedor").range.Text = proveedor
        If .Bookmarks.Exists("Ruc") Then .Bookmarks("Ruc").range.Text = ruc
        If .Bookmarks.Exists("Objeto_de_Contratacion2") Then .Bookmarks("Objeto_de_Contratacion2").range.Text = objetoDeContratacion2
        If .Bookmarks.Exists("Tecnico_requirente") Then .Bookmarks("Tecnico_requirente").range.Text = tecnicoRequirente
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglaEntidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:="PROEST2023", Structure:=True

    ' Ubicarse en la hoja "CUADRO-INF"
    ThisWorkbook.Sheets("CUADRO-INF").Activate

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing

End Sub



