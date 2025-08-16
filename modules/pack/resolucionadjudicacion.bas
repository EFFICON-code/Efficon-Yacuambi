Attribute VB_Name = "resolucionadjudicacion"
Sub Resolucion_Adjudicacion()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim plantillaRuta As Variant
    Dim guardarRuta As Variant
    Dim ws As Worksheet
    Dim objetoDeContratacion As String
    Dim Requerimiento As String
    Dim fechaRequerimiento As String
    Dim presidente As String
    Dim cargoPresidente As String
    Dim objetoDeContratacion1 As String
    Dim certificacionCATE As String
    Dim fechaCertificacion As String
    Dim tecnicoUnidad As String
    Dim cargoTecnico As String
    Dim codigoCPC As String
    Dim autorizacion As String
    Dim fechaAutorizacion As String
    Dim presidente1 As String
    Dim cargoPresidente1 As String
    Dim objetoDeContratacion2 As String
    Dim fechaPublicacion As String
    Dim nroNIC As String
    Dim titulo As String
    Dim objetoDeContratacion3 As String
    Dim nroCuadro As String
    Dim financiero As String
    Dim cargoFinanciero As String
    Dim certificacioPresupuestaria As String
    Dim fechaCertificacionPresupuesto As String
    Dim partida As String
    Dim denominacion As String
    Dim presupuesto As String
    Dim valorLetras As String
    Dim objetoDeContratacion4 As String
    Dim nroCuadro1 As String
    Dim nroNIC1 As String
    Dim objetoDeContratacion5 As String
    Dim nroNIC2 As String
    Dim objetoDeContratacion6 As String
    Dim proveedor As String
    Dim ruc As String
    Dim presupuesto1 As String
    Dim valorLetras1 As String
    Dim plazo As String
    Dim administrador As String
    Dim cargoAdministrador As String
    Dim objetoDeContratacion7 As String
    Dim proveedor1 As String
    Dim ruc1 As String
    Dim entidad As String
    Dim presidente2 As String
    Dim cargoPresidente2 As String
    Dim compras As String
    Dim siglaEntidad As String
    Dim periodo As String
    Dim fecha As String
    Dim lugar As String
    Dim CLAVE As String

    ' Clave para desproteger la hoja "SECUENCIAS"
    CLAVE = "Admin1991"

    ' Desproteger la estructura del libro
    ThisWorkbook.Unprotect password:="PROEST2023"

    ' Mostrar cuadro de diálogo para seleccionar la plantilla de Word
    plantillaRuta = Application.GetOpenFilename("Archivos de Word (*.docx), *.docx", , "Seleccionar plantilla de Word")
    If plantillaRuta = "False" Then Exit Sub ' Si el usuario cancela la selección, salir de la macro

    ' Mostrar cuadro de diálogo para seleccionar la ubicación donde guardar el documento terminado
    guardarRuta = Application.GetSaveAsFilename("Resolucion_Adjudicacion_Terminado.docx", "Documentos de Word (*.docx), *.docx", , "Guardar documento terminado")
    If guardarRuta = "False" Then Exit Sub ' Si el usuario cancela la selección, salir de la macro

    ' Asignar la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("SECUENCIAS")

    ' Mostrar y desproteger la hoja si está oculta y protegida
    If ws.Visible = xlSheetVeryHidden Or ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
    ws.Unprotect password:=CLAVE

    ' Leer datos de Excel
    objetoDeContratacion = CStr(ws.range("Q2").Value)
    Requerimiento = CStr(ws.range("M2").Value)
    fechaRequerimiento = CStr(ws.range("N2").Value)
    presidente = CStr(ws.range("B2").Value)
    cargoPresidente = CStr(ws.range("C2").Value)
    objetoDeContratacion1 = CStr(ws.range("Q2").Value)
    certificacionCATE = CStr(ws.range("DG2").Value)
    fechaCertificacion = CStr(ws.range("DH2").Value)
    tecnicoUnidad = CStr(ws.range("G2").Value)
    cargoTecnico = CStr(ws.range("H2").Value)
    codigoCPC = CStr(ws.range("BA2").Value)
    autorizacion = CStr(ws.range("DK2").Value)
    fechaAutorizacion = CStr(ws.range("DL2").Value)
    presidente1 = CStr(ws.range("B2").Value)
    cargoPresidente1 = CStr(ws.range("C2").Value)
    objetoDeContratacion2 = CStr(ws.range("Q2").Value)
    fechaPublicacion = CStr(ws.range("DQ2").Value)
    nroNIC = CStr(ws.range("DP2").Value)
    titulo = CStr(ws.range("AO2").Value)
    objetoDeContratacion3 = CStr(ws.range("Q2").Value)
    nroCuadro = CStr(ws.range("DM2").Value)
    financiero = CStr(ws.range("CH2").Value)
    cargoFinanciero = CStr(ws.range("CI2").Value)
    certificacioPresupuestaria = CStr(ws.range("DR2").Value)
    fechaCertificacionPresupuesto = CStr(ws.range("DS2").Value)
    partida = CStr(ws.range("BP2").Value)
    denominacion = CStr(ws.range("BQ2").Value)
    presupuesto = CStr(ws.range("DC2").Value)
    valorLetras = CStr(ws.range("DD2").Value)
    objetoDeContratacion4 = CStr(ws.range("Q2").Value)
    nroCuadro1 = CStr(ws.range("DM2").Value)
    nroNIC1 = CStr(ws.range("DP2").Value)
    objetoDeContratacion5 = CStr(ws.range("Q2").Value)
    nroNIC2 = CStr(ws.range("DP2").Value)
    objetoDeContratacion6 = CStr(ws.range("Q2").Value)
    proveedor = CStr(ws.range("DE2").Value)
    ruc = CStr(ws.range("DF2").Value)
    presupuesto1 = CStr(ws.range("DC2").Value)
    valorLetras1 = CStr(ws.range("DD2").Value)
    plazo = CStr(ws.range("T2").Value)
    administrador = CStr(ws.range("DJ2").Value)
    cargoAdministrador = CStr(ws.range("GS2").Value)
    objetoDeContratacion7 = CStr(ws.range("Q2").Value)
    proveedor1 = CStr(ws.range("DE2").Value)
    ruc1 = CStr(ws.range("DF2").Value)
    entidad = CStr(ws.range("A2").Value)
    presidente2 = CStr(ws.range("B2").Value)
    cargoPresidente2 = CStr(ws.range("C2").Value)
    compras = CStr(ws.range("G2").Value)
    siglaEntidad = CStr(ws.range("HA2").Value)
    periodo = CStr(ws.range("HB2").Value)
    fecha = CStr(ws.range("GZ2").Value)
    lugar = CStr(ws.range("FQ2").Value)

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
        If .Bookmarks.Exists("Objeto_de_Contratacion") Then .Bookmarks("Objeto_de_Contratacion").range.Text = objetoDeContratacion
        If .Bookmarks.Exists("Requerimiento") Then .Bookmarks("Requerimiento").range.Text = Requerimiento
        If .Bookmarks.Exists("Fecha_requerimiento") Then .Bookmarks("Fecha_requerimiento").range.Text = fechaRequerimiento
        If .Bookmarks.Exists("Presidente") Then .Bookmarks("Presidente").range.Text = presidente
        If .Bookmarks.Exists("Cargo_presidente") Then .Bookmarks("Cargo_presidente").range.Text = cargoPresidente
        If .Bookmarks.Exists("Objeto_de_Contratacion1") Then .Bookmarks("Objeto_de_Contratacion1").range.Text = objetoDeContratacion1
        If .Bookmarks.Exists("Certificacion_CATE") Then .Bookmarks("Certificacion_CATE").range.Text = certificacionCATE
        If .Bookmarks.Exists("Fecha_certificacion") Then .Bookmarks("Fecha_certificacion").range.Text = fechaCertificacion
        If .Bookmarks.Exists("Tecnico_Unidad") Then .Bookmarks("Tecnico_Unidad").range.Text = tecnicoUnidad
        If .Bookmarks.Exists("Cargo_Tecnico") Then .Bookmarks("Cargo_Tecnico").range.Text = cargoTecnico
        If .Bookmarks.Exists("Codigo_CPC") Then .Bookmarks("Codigo_CPC").range.Text = codigoCPC
        If .Bookmarks.Exists("Autorizacion") Then .Bookmarks("Autorizacion").range.Text = autorizacion
        If .Bookmarks.Exists("Fecha_Autorizacion") Then .Bookmarks("Fecha_Autorizacion").range.Text = fechaAutorizacion
        If .Bookmarks.Exists("Presidente1") Then .Bookmarks("Presidente1").range.Text = presidente1
        If .Bookmarks.Exists("Cargo_presidente1") Then .Bookmarks("Cargo_presidente1").range.Text = cargoPresidente1
        If .Bookmarks.Exists("Objeto_de_Contratacion2") Then .Bookmarks("Objeto_de_Contratacion2").range.Text = objetoDeContratacion2
        If .Bookmarks.Exists("Fecha_publicacion") Then .Bookmarks("Fecha_publicacion").range.Text = fechaPublicacion
        If .Bookmarks.Exists("Nro_NIC") Then .Bookmarks("Nro_NIC").range.Text = nroNIC
        If .Bookmarks.Exists("Titulo") Then .Bookmarks("Titulo").range.Text = titulo
        If .Bookmarks.Exists("Objeto_de_Contratacion3") Then .Bookmarks("Objeto_de_Contratacion3").range.Text = objetoDeContratacion3
        If .Bookmarks.Exists("Nro_Cuadro") Then .Bookmarks("Nro_Cuadro").range.Text = nroCuadro
        If .Bookmarks.Exists("Financiero") Then .Bookmarks("Financiero").range.Text = financiero
        If .Bookmarks.Exists("Cargo_financiero") Then .Bookmarks("Cargo_financiero").range.Text = cargoFinanciero
        If .Bookmarks.Exists("Certificacio_presupuestaria") Then .Bookmarks("Certificacio_presupuestaria").range.Text = certificacioPresupuestaria
        If .Bookmarks.Exists("Fecha_Certificacion") Then .Bookmarks("Fecha_Certificacion").range.Text = fechaCertificacionPresupuesto
        If .Bookmarks.Exists("Partida") Then .Bookmarks("Partida").range.Text = partida
        If .Bookmarks.Exists("Denominación") Then .Bookmarks("Denominación").range.Text = denominacion
        If .Bookmarks.Exists("Presupuesto") Then .Bookmarks("Presupuesto").range.Text = presupuesto
        If .Bookmarks.Exists("Valor_letras") Then .Bookmarks("Valor_letras").range.Text = valorLetras
        If .Bookmarks.Exists("Objeto_de_Contratacion4") Then .Bookmarks("Objeto_de_Contratacion4").range.Text = objetoDeContratacion4
        If .Bookmarks.Exists("Nro_Cuadro1") Then .Bookmarks("Nro_Cuadro1").range.Text = nroCuadro1
        If .Bookmarks.Exists("Nro_NIC1") Then .Bookmarks("Nro_NIC1").range.Text = nroNIC1
        If .Bookmarks.Exists("Objeto_de_Contratacion5") Then .Bookmarks("Objeto_de_Contratacion5").range.Text = objetoDeContratacion5
        If .Bookmarks.Exists("Nro_NIC2") Then .Bookmarks("Nro_NIC2").range.Text = nroNIC2
        If .Bookmarks.Exists("Objeto_de_Contratacion6") Then .Bookmarks("Objeto_de_Contratacion6").range.Text = objetoDeContratacion6
        If .Bookmarks.Exists("Proveedor") Then .Bookmarks("Proveedor").range.Text = proveedor
        If .Bookmarks.Exists("Ruc") Then .Bookmarks("Ruc").range.Text = ruc
        If .Bookmarks.Exists("Presupuesto1") Then .Bookmarks("Presupuesto1").range.Text = presupuesto1
        If .Bookmarks.Exists("Valor_letras1") Then .Bookmarks("Valor_letras1").range.Text = valorLetras1
        If .Bookmarks.Exists("Plazo") Then .Bookmarks("Plazo").range.Text = plazo
        If .Bookmarks.Exists("Administrador") Then .Bookmarks("Administrador").range.Text = administrador
        If .Bookmarks.Exists("Cargo_Administrador") Then .Bookmarks("Cargo_Administrador").range.Text = cargoAdministrador
        If .Bookmarks.Exists("Objeto_de_Contratacion7") Then .Bookmarks("Objeto_de_Contratacion7").range.Text = objetoDeContratacion7
        If .Bookmarks.Exists("Proveedor1") Then .Bookmarks("Proveedor1").range.Text = proveedor1
        If .Bookmarks.Exists("Ruc1") Then .Bookmarks("Ruc1").range.Text = ruc1
        If .Bookmarks.Exists("Entidad") Then .Bookmarks("Entidad").range.Text = entidad
        If .Bookmarks.Exists("Presidente2") Then .Bookmarks("Presidente2").range.Text = presidente2
        If .Bookmarks.Exists("Cargo_presidente2") Then .Bookmarks("Cargo_presidente2").range.Text = cargoPresidente2
        If .Bookmarks.Exists("Compras") Then .Bookmarks("Compras").range.Text = compras
        If .Bookmarks.Exists("Sigla_entidad") Then .Bookmarks("Sigla_entidad").range.Text = siglaEntidad
        If .Bookmarks.Exists("Periodo") Then .Bookmarks("Periodo").range.Text = periodo
        If .Bookmarks.Exists("Fecha") Then .Bookmarks("Fecha").range.Text = fecha
        If .Bookmarks.Exists("Lugar") Then .Bookmarks("Lugar").range.Text = lugar
    End With

    ' Guardar y cerrar documento
    wdDoc.SaveAs2 fileName:=guardarRuta
    wdDoc.Close
    wdApp.Quit

    ' Proteger la estructura del libro
    ThisWorkbook.Protect password:="PROEST2023", Structure:=True

    ' Ubicarse en la hoja "ORDEN"
    ThisWorkbook.Sheets("ORDEN").Activate

    ' Liberar objetos
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing

End Sub

