Attribute VB_Name = "modArmarios"
Option Explicit

' LISTADO DE ARMARIOS ELECTRICOS
' Importar este módulo desde el Editor VBA (Alt+F11 > Archivo > Importar archivo).
' Guardar después el libro como .XLSM.
'
' NOTA: La cadena original se conserva íntegra. Su descomposición por ";" queda pendiente
' de conocer el esquema exacto de campos de la aplicación de origen.

Private Const SH_CONSOLA As String = "CONSOLA"
Private Const SH_CATALOGO As String = "CATALOGO"
Private Const SH_HISTORICO As String = "HISTORICO"
Private Const SH_CONFIG As String = "CONFIGURACION"


' -----------------------------------------------------------------------------
' SISTEMA DE ARCHIVOS
' Se crea junto al libro: ARMARIOS ELECTRICOS\FOTOS\[PREFIJO] y REGISTROS.
' El catálogo TXT es una exportación completa; el registro TXT es un diario acumulativo.
' -----------------------------------------------------------------------------
Private Const NOMBRE_RAIZ As String = "ARMARIOS ELECTRICOS"
Private Const NOMBRE_FOTOS As String = "FOTOS"
Private Const NOMBRE_REGISTROS As String = "REGISTROS"
Private Const ARCHIVO_CATALOGO As String = "CATALOGO_ARMARIOS.txt"
Private Const ARCHIVO_REGISTRO As String = "REGISTRO_ARMARIOS.txt"

Private Function RutaRaizSistema() As String
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 710, "modArmarios", _
                  "Guarda primero el libro Excel en una carpeta para poder crear el sistema de archivos."
    End If
    RutaRaizSistema = ThisWorkbook.Path & Application.PathSeparator & NOMBRE_RAIZ
End Function

Private Sub CrearCarpetaSiNoExiste(ByVal ruta As String)
    If Len(Dir$(ruta, vbDirectory)) = 0 Then MkDir ruta
End Sub

Public Sub InicializarSistemaArchivos()
    On Error GoTo ErrHandler
    Dim fso As Object, raiz As String, rutaFotos As String, rutaRegistros As String
    Dim ws As Worksheet, ultima As Long, r As Long, prefijo As String
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    raiz = RutaRaizSistema()
    CrearCarpetaSiNoExiste raiz
    rutaFotos = raiz & Application.PathSeparator & NOMBRE_FOTOS
    rutaRegistros = raiz & Application.PathSeparator & NOMBRE_REGISTROS
    CrearCarpetaSiNoExiste rutaFotos
    CrearCarpetaSiNoExiste rutaRegistros

    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    ultima = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For r = 2 To ultima
        prefijo = Trim$(CStr(ws.Cells(r, "D").Value))
        If Len(prefijo) > 0 Then CrearCarpetaSiNoExiste rutaFotos & Application.PathSeparator & prefijo
    Next r

    ' Asegura que los dos archivos existan desde la primera inicialización.
    If Not fso.FileExists(rutaRegistros & Application.PathSeparator & ARCHIVO_CATALOGO) Then
        ExportarCatalogoTXT
    End If
    If Not fso.FileExists(rutaRegistros & Application.PathSeparator & ARCHIVO_REGISTRO) Then
        Set ts = fso.CreateTextFile(rutaRegistros & Application.PathSeparator & ARCHIVO_REGISTRO, True, True)
        ts.WriteLine "Fecha y hora" & vbTab & "Usuario Windows" & vbTab & "Operacion" & vbTab & "Referencia" & vbTab & "Descripcion" & vbTab & "Observacion"
        ts.Close
    End If
    Exit Sub
ErrHandler:
    MsgBox "No se ha podido inicializar la estructura de carpetas:" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function TextoSeguro(ByVal valor As Variant) As String
    Dim t As String
    t = CStr(valor)
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    TextoSeguro = t
End Function

Public Sub ExportarCatalogoTXT()
    On Error GoTo ErrHandler
    Dim fso As Object, ts As Object, ws As Worksheet, ruta As String
    Dim ultima As Long, r As Long, c As Long, linea As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Worksheets(SH_CATALOGO)
    ruta = RutaRaizSistema() & Application.PathSeparator & NOMBRE_REGISTROS & _
           Application.PathSeparator & ARCHIVO_CATALOGO
    CrearCarpetaSiNoExiste RutaRaizSistema()
    CrearCarpetaSiNoExiste RutaRaizSistema() & Application.PathSeparator & NOMBRE_REGISTROS
    Set ts = fso.CreateTextFile(ruta, True, True)
    ts.WriteLine "Referencia" & vbTab & "Descripcion" & vbTab & "Descripcion antigua 1" & vbTab & _
                 "Descripcion antigua 2" & vbTab & "Observacion" & vbTab & "Fecha de alta" & vbTab & _
                 "Usuario de alta" & vbTab & "Fecha de modificacion" & vbTab & "Usuario de modificacion" & vbTab & _
                 "Cadena original" & vbTab & "Tipo de registro" & vbTab & "Referencia historica 1" & vbTab & "Referencia historica 2"
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To ultima
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            linea = vbNullString
            For c = 1 To 13
                If c > 1 Then linea = linea & vbTab
                If c = 6 Or c = 8 Then
                    If IsDate(ws.Cells(r, c).Value) Then
                        linea = linea & Format$(ws.Cells(r, c).Value, "dd/mm/yyyy hh:nn:ss")
                    Else
                        linea = linea & TextoSeguro(ws.Cells(r, c).Value)
                    End If
                Else
                    linea = linea & TextoSeguro(ws.Cells(r, c).Value)
                End If
            Next c
            ts.WriteLine linea
        End If
    Next r
    ts.Close
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    MsgBox "No se ha podido actualizar CATALOGO_ARMARIOS.txt:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub AnotarRegistroArchivo(ByVal operacion As String, ByVal referencia As String, ByVal descripcion As String, ByVal observacion As String)
    On Error GoTo ErrHandler
    Dim fso As Object, ts As Object, ruta As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    ruta = RutaRaizSistema() & Application.PathSeparator & NOMBRE_REGISTROS & _
           Application.PathSeparator & ARCHIVO_REGISTRO
    If Not fso.FolderExists(RutaRaizSistema() & Application.PathSeparator & NOMBRE_REGISTROS) Then
        InicializarSistemaArchivos
    End If
    Set ts = fso.OpenTextFile(ruta, 8, True, -1)
    ts.WriteLine Format$(Now, "dd/mm/yyyy hh:nn:ss") & vbTab & TextoSeguro(UsuarioWindows()) & vbTab & _
                 TextoSeguro(operacion) & vbTab & TextoSeguro(referencia) & vbTab & _
                 TextoSeguro(descripcion) & vbTab & TextoSeguro(observacion)
    ts.Close
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    MsgBox "No se ha podido añadir la operación a REGISTRO_ARMARIOS.txt:" & vbCrLf & Err.Description, vbExclamation
End Sub

' Excel ejecuta Auto_Open al abrir el libro con macros habilitadas.
Public Sub Auto_Open()
    InicializarSistemaArchivos
    ExportarCatalogoTXT
End Sub


Private Function UsuarioWindows() As String
    UsuarioWindows = Environ$("USERNAME")
    If Len(Trim$(UsuarioWindows)) = 0 Then UsuarioWindows = Application.UserName
End Function

Private Function SiguienteFila(ByVal ws As Worksheet, ByVal col As Long) As Long
    SiguienteFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row + 1
    If SiguienteFila < 2 Then SiguienteFila = 2
End Function

Private Function PrefijoValido(ByVal prefijo As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, "D").Value)), Trim$(prefijo), vbTextCompare) = 0 Then
            PrefijoValido = True
            Exit Function
        End If
    Next r
End Function

Private Function ExisteReferencia(ByVal referencia As String) As Boolean
    Dim ws As Worksheet, f As Range
    Set ws = ThisWorkbook.Worksheets(SH_CATALOGO)
    Set f = ws.Columns("A").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then ExisteReferencia = True: Exit Function
    Set f = ws.Columns("L").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then ExisteReferencia = True: Exit Function
    Set f = ws.Columns("M").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then ExisteReferencia = True
End Function

Private Function GenerarReferencia(ByVal prefijo As String, ByVal contador As Long) As String
    GenerarReferencia = UCase$(Trim$(prefijo)) & "-" & Format$(contador, "0000")
    ' Formato acordado: PREFIJO-XX-XX. Ejemplo: F01-00-01
    GenerarReferencia = UCase$(Trim$(prefijo)) & "-" & Left$(Format$(contador, "0000"), 2) & "-" & Right$(Format$(contador, "0000"), 2)
End Function

Private Sub RegistrarHistorico(ByVal operacion As String, ByVal referencia As String, _
                               ByVal campo As String, ByVal anterior As Variant, ByVal nuevo As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_HISTORICO)
    r = SiguienteFila(ws, 1)
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = UsuarioWindows()
    ws.Cells(r, 3).Value = operacion
    ws.Cells(r, 4).Value = referencia
    ws.Cells(r, 5).Value = campo
    ws.Cells(r, 6).Value = anterior
    ws.Cells(r, 7).Value = nuevo
End Sub

Public Sub AltaArmario()
    Dim wsUI As Worksheet, wsCat As Worksheet, wsCfg As Worksheet
    Dim prefijo As String, referencia As String, descripcion As String
    Dim descAnt1 As String, descAnt2 As String, observacion As String
    Dim cadena As String, refAnt1 As String, refAnt2 As String
    Dim contador As Long, fila As Long, intento As Long

    InicializarSistemaArchivos
    Set wsUI = ThisWorkbook.Worksheets(SH_CONSOLA)
    Set wsCat = ThisWorkbook.Worksheets(SH_CATALOGO)
    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)

    prefijo = Trim$(CStr(wsUI.Range("B9").Value))
    descripcion = Trim$(CStr(wsUI.Range("B10").Value))
    descAnt1 = CStr(wsUI.Range("B11").Value)
    descAnt2 = CStr(wsUI.Range("B12").Value)
    observacion = CStr(wsUI.Range("B13").Value)
    cadena = CStr(wsUI.Range("B14").Value)
    refAnt1 = Trim$(CStr(wsUI.Range("B15").Value))
    refAnt2 = Trim$(CStr(wsUI.Range("B16").Value))

    If Not PrefijoValido(prefijo) Then
        MsgBox "Selecciona un prefijo válido de la lista.", vbExclamation
        Exit Sub
    End If
    If Len(descripcion) = 0 Then
        MsgBox "La descripción es obligatoria.", vbExclamation
        Exit Sub
    End If
    If Len(refAnt1) > 0 And StrComp(refAnt1, refAnt2, vbTextCompare) = 0 Then
        MsgBox "Las dos referencias históricas no pueden ser iguales.", vbExclamation
        Exit Sub
    End If
    If (Len(refAnt1) > 0 And ExisteReferencia(refAnt1)) Or (Len(refAnt2) > 0 And ExisteReferencia(refAnt2)) Then
        MsgBox "Una de las referencias históricas ya está asignada a otro armario.", vbCritical
        Exit Sub
    End If

    contador = CLng(Val(wsCfg.Range("B2").Value))
    Do
        contador = contador + 1
        If contador > 9999 Then
            MsgBox "Se ha alcanzado el máximo de 9.999 referencias.", vbCritical
            Exit Sub
        End If
        referencia = GenerarReferencia(prefijo, contador)
        intento = intento + 1
        If intento > 9999 Then
            MsgBox "No se ha podido encontrar una referencia disponible.", vbCritical
            Exit Sub
        End If
    Loop While ExisteReferencia(referencia)

    fila = SiguienteFila(wsCat, 1)
    Application.ScreenUpdating = False
    wsCat.Cells(fila, 1).Value = referencia
    wsCat.Cells(fila, 2).Value = descripcion
    wsCat.Cells(fila, 3).Value = descAnt1
    wsCat.Cells(fila, 4).Value = descAnt2
    wsCat.Cells(fila, 5).Value = observacion
    wsCat.Cells(fila, 6).Value = Date
    wsCat.Cells(fila, 7).Value = UsuarioWindows()
    wsCat.Cells(fila, 8).Value = vbNullString
    wsCat.Cells(fila, 9).Value = vbNullString
    wsCat.Cells(fila, 10).Value = cadena
    wsCat.Cells(fila, 11).Value = "NUEVA"
    wsCat.Cells(fila, 12).Value = refAnt1
    wsCat.Cells(fila, 13).Value = refAnt2
    wsCat.Cells(fila, 6).NumberFormat = "dd/mm/yyyy"

    wsCfg.Range("B2").Value = contador
    RegistrarHistorico "ALTA", referencia, "REGISTRO COMPLETO", "", "Alta de armario"
    AnotarRegistroArchivo "ALTA", referencia, descripcion, observacion
    ExportarCatalogoTXT
    Application.ScreenUpdating = True

    wsUI.Range("B5").Value = referencia
    wsUI.Range("B6").Value = "Alta realizada: " & referencia & " | Usuario: " & UsuarioWindows()
    MsgBox "Armario creado correctamente: " & referencia, vbInformation
End Sub

Public Sub ConsultarArmario()
    Dim wsUI As Worksheet, wsCat As Worksheet, f As Range, referencia As String
    Dim fila As Long
    Set wsUI = ThisWorkbook.Worksheets(SH_CONSOLA)
    Set wsCat = ThisWorkbook.Worksheets(SH_CATALOGO)
    referencia = Trim$(CStr(wsUI.Range("B5").Value))
    If Len(referencia) = 0 Then
        MsgBox "Introduce una referencia nueva o histórica en B5.", vbExclamation
        Exit Sub
    End If

    Set f = wsCat.Columns("A").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Set f = wsCat.Columns("L").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Set f = wsCat.Columns("M").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If f Is Nothing Then
        wsUI.Range("B6").Value = "No se ha encontrado la referencia."
        MsgBox "No se ha encontrado el armario.", vbInformation
        Exit Sub
    End If

    fila = f.Row
    wsUI.Range("B6").Value = "Encontrado: " & wsCat.Cells(fila, 1).Value & " | " & wsCat.Cells(fila, 2).Value
    wsUI.Range("B9").Value = IIf(wsCat.Cells(fila, 11).Value = "NUEVA", Split(CStr(wsCat.Cells(fila, 1).Value), "-")(0), "")
    wsUI.Range("B10").Value = wsCat.Cells(fila, 2).Value
    wsUI.Range("B11").Value = wsCat.Cells(fila, 3).Value
    wsUI.Range("B12").Value = wsCat.Cells(fila, 4).Value
    wsUI.Range("B13").Value = wsCat.Cells(fila, 5).Value
    wsUI.Range("B14").Value = wsCat.Cells(fila, 10).Value
    wsUI.Range("B15").Value = wsCat.Cells(fila, 12).Value
    wsUI.Range("B16").Value = wsCat.Cells(fila, 13).Value
End Sub

Public Sub ModificarArmario()
    Dim wsUI As Worksheet, wsCat As Worksheet, f As Range
    Dim referencia As String, fila As Long
    Dim campos As Variant, c As Long, anterior As Variant, nuevo As Variant
    Dim columnas As Variant, nombres As Variant

    InicializarSistemaArchivos
    Set wsUI = ThisWorkbook.Worksheets(SH_CONSOLA)
    Set wsCat = ThisWorkbook.Worksheets(SH_CATALOGO)
    referencia = Trim$(CStr(wsUI.Range("B5").Value))
    If Len(referencia) = 0 Then
        MsgBox "Consulta primero el armario que deseas modificar.", vbExclamation
        Exit Sub
    End If
    Set f = wsCat.Columns("A").Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        MsgBox "Solo se pueden modificar registros localizados por su referencia nueva. Las referencias históricas no son editables.", vbExclamation
        Exit Sub
    End If
    fila = f.Row
    If UCase$(CStr(wsCat.Cells(fila, 11).Value)) <> "NUEVA" Then
        MsgBox "Este registro no es editable.", vbExclamation
        Exit Sub
    End If

    columnas = Array(2, 3, 4, 5, 10, 12, 13)
    nombres = Array("Descripción", "Descripción antigua 1", "Descripción antigua 2", "Observación", "Cadena original", "Referencia histórica 1", "Referencia histórica 2")
    campos = Array("B10", "B11", "B12", "B13", "B14", "B15", "B16")

    ' Las celdas de consola representan el nuevo contenido completo de la ficha.
    For c = LBound(columnas) To UBound(columnas)
        anterior = wsCat.Cells(fila, columnas(c)).Value
        nuevo = wsUI.Range(CStr(campos(c))).Value
        If CStr(anterior) <> CStr(nuevo) Then
            If (columnas(c) = 12 Or columnas(c) = 13) And Len(Trim$(CStr(nuevo))) > 0 Then
                If ExisteReferencia(CStr(nuevo)) Then
                    If StrComp(CStr(anterior), CStr(nuevo), vbTextCompare) <> 0 Then
                        MsgBox "La referencia histórica " & CStr(nuevo) & " ya está asignada.", vbCritical
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next c

    For c = LBound(columnas) To UBound(columnas)
        anterior = wsCat.Cells(fila, columnas(c)).Value
        nuevo = wsUI.Range(CStr(campos(c))).Value
        If CStr(anterior) <> CStr(nuevo) Then
            wsCat.Cells(fila, columnas(c)).Value = nuevo
            RegistrarHistorico "MODIFICACIÓN", referencia, CStr(nombres(c)), anterior, nuevo
        End If
    Next c
    wsCat.Cells(fila, 8).Value = Now
    wsCat.Cells(fila, 9).Value = UsuarioWindows()
    RegistrarHistorico "MODIFICACIÓN", referencia, "USUARIO / FECHA", "", UsuarioWindows() & " - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    AnotarRegistroArchivo "MODIFICACION", referencia, CStr(wsCat.Cells(fila, 2).Value), CStr(wsCat.Cells(fila, 5).Value)
    ExportarCatalogoTXT
    wsUI.Range("B6").Value = "Modificación guardada: " & referencia
    MsgBox "Modificación guardada correctamente.", vbInformation
End Sub

Public Sub CargarPortapapeles()
    On Error GoTo ErrHandler
    Dim datos As Object
    Set datos = CreateObject("MSForms.DataObject")
    datos.GetFromClipboard
    ThisWorkbook.Worksheets(SH_CONSOLA).Range("B14").Value = datos.GetText(1)
    MsgBox "Cadena copiada desde el portapapeles.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se ha podido leer el portapapeles. Comprueba que has copiado previamente el texto.", vbExclamation
End Sub

Public Sub LimpiarConsola()
    With ThisWorkbook.Worksheets(SH_CONSOLA)
        .Range("B5:B6").ClearContents
        .Range("B10:B16").ClearContents
        .Range("B9").Value = "F01"
    End With
End Sub
