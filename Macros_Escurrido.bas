' =============================================================================
'  MIRFE Primavera-Rosa — soporte para CE Escurrido y pH Escurrido
'
'  CÓMO USARLO
'    1. En Excel: Programador → Visual Basic (o Alt+F11).
'    2. Si ya existe un módulo con una versión anterior de este archivo,
'       bórralo y vuelve a insertar este.
'    3. Ejecutar UNA SOLA VEZ la macro  Agregar_Escurrido_A_Dinamicas
'    4. Reemplazar las macros Historico_Valvulas y Limpiar del Módulo1 por las
'       versiones de más abajo.
'
'  SOBRE EL NOMBRE DE LAS COLUMNAS
'  Hoy los encabezados no están iguales en las dos hojas:
'       Reporte   (Tabla1)    → "C.E Escurridos"   con punto
'       Historico (Historico) → "C,E Escurridos"   con coma
'  No hace falta corregirlo: Historico_Valvulas copia por posición, no por
'  nombre, y las macros de abajo buscan el campo ignorando signos, acentos,
'  espacios y el plural. Así funcionan con "C,E Escurridos", "C.E Escurrido"
'  o "CE Escurrido" indistintamente.
'
'  Lo único que SÍ importa es que las dos columnas estén en la misma posición
'  en ambas hojas (Z y AA) y dentro de las tablas, que es como están ahora.
' =============================================================================

Option Explicit

' Se busca por forma compacta, no por texto exacto (ver nota de arriba)
Private Const CAMPO_CE As String = "ceescurrido"
Private Const CAMPO_PH As String = "phescurrido"


' --- Agrega los dos campos a las 53 tablas dinámicas ------------------------
'  Las 53 comparten una sola caché, así que basta refrescarla una vez para que
'  los campos nuevos existan; pero cada tabla necesita que se le agreguen al
'  área de valores para que aparezcan como filas. Eso es lo que hace esto.
Sub Agregar_Escurrido_A_Dinamicas()

    Dim ws As Worksheet, pt As PivotTable
    Dim claves As Variant, clave As Variant
    Dim pf As PivotField
    Dim agregados As Long, existentes As Long, sinCampo As Long
    Dim primera As Boolean

    claves = Array(CAMPO_CE, CAMPO_PH)
    primera = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables

            ' La primera refresca la caché compartida; las demás ya la ven
            If primera Then
                pt.PivotCache.Refresh
                primera = False
            End If

            For Each clave In claves
                Set pf = BuscarCampo(pt, CStr(clave))

                If pf Is Nothing Then
                    sinCampo = sinCampo + 1
                ElseIf YaEsDato(pt, pf.SourceName) Then
                    existentes = existentes + 1
                Else
                    pf.Orientation = xlDataField
                    pf.Function = xlSum
                    agregados = agregados + 1
                End If
            Next clave

        Next pt
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Campos agregados: " & agregados & vbCrLf & _
          "Ya estaban:       " & existentes
    If sinCampo > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "ATENCIÓN: en " & sinCampo & " casos no se encontró el campo." & vbCrLf & _
              "Revisa que las dos columnas del escurrido estén DENTRO de la" & vbCrLf & _
              "tabla Historico y que su encabezado contenga 'Escurrido'."
    End If
    MsgBox msg, vbInformation, "Escurrido en tablas dinámicas"

End Sub


' Quita signos, acentos y espacios, y vuelve singular, para poder comparar
' encabezados escritos de distinta manera:
'   "C,E Escurridos" → "ceescurrido"      "pH Escurrido" → "phescurrido"
Private Function Compactar(ByVal texto As String) As String

    Dim i As Long, ch As String, r As String
    texto = LCase$(texto)

    ' acentos que podrían aparecer en un encabezado
    texto = Replace(texto, "á", "a"): texto = Replace(texto, "é", "e")
    texto = Replace(texto, "í", "i"): texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ú", "u"): texto = Replace(texto, "ñ", "n")

    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then r = r & ch
    Next i

    ' plural: "escurridos" y "escurrido" deben dar lo mismo
    If Len(r) > 1 And Right$(r, 1) = "s" Then r = Left$(r, Len(r) - 1)

    Compactar = r

End Function


' Busca en la tabla dinámica el campo cuya forma compacta contenga la clave
Private Function BuscarCampo(pt As PivotTable, ByVal clave As String) As PivotField

    Dim pf As PivotField
    For Each pf In pt.PivotFields
        If InStr(Compactar(pf.Name), clave) > 0 Then
            Set BuscarCampo = pf
            Exit Function
        End If
    Next pf
    Set BuscarCampo = Nothing

End Function


' ¿La tabla dinámica ya muestra ese campo en el área de valores?
Private Function YaEsDato(pt As PivotTable, ByVal origen As String) As Boolean

    Dim df As PivotField
    For Each df In pt.DataFields
        If df.SourceName = origen Then
            YaEsDato = True
            Exit Function
        End If
    Next df
    YaEsDato = False

End Function


' =============================================================================
'  REEMPLAZOS PARA LAS MACROS EXISTENTES DEL Módulo1
'  Borrar las versiones actuales y dejar estas.
' =============================================================================

' --- Antes copiaba A4:Y4; ahora llega hasta AA4 para incluir el escurrido ----
Sub Historico_Valvulas()

    Dim rep As Worksheet, his As Worksheet
    Dim ultima As Long, nFilas As Long

    Set rep = ThisWorkbook.Worksheets("Reporte")
    Set his = ThisWorkbook.Worksheets("Historico")

    ultima = rep.Cells(rep.Rows.Count, 1).End(xlUp).Row
    If ultima < 4 Then
        MsgBox "No hay filas que cargar en la hoja Reporte.", vbExclamation
        Exit Sub
    End If
    nFilas = ultima - 3

    rep.Range("A4:AA" & ultima).Copy

    his.Range("A2").Resize(nFilas, 27).Insert Shift:=xlDown
    his.Range("A2").Resize(nFilas, 27).PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False
    rep.Activate
    rep.Range("B1").Select

End Sub


' --- Antes borraba C4:W; ahora borra también las dos columnas nuevas --------
'  Se calcula la última fila por la columna A en vez de usar End(xlDown): como
'  el escurrido solo trae dato una vez por semana, sus columnas vienen casi
'  vacías y End(xlDown) se iría hasta el final de la hoja.
Sub Limpiar()

    Dim rep As Worksheet
    Dim ultima As Long

    Set rep = ThisWorkbook.Worksheets("Reporte")
    ultima = rep.Cells(rep.Rows.Count, 1).End(xlUp).Row
    If ultima < 4 Then Exit Sub

    rep.Range("C4:W" & ultima).ClearContents     ' mediciones del día
    rep.Range("Z4:AA" & ultima).ClearContents    ' CE y pH escurrido

    rep.Range("B1").Select

End Sub
