' =============================================================================
'  MIRFE Primavera-Rosa — soporte para CE Escurrido y pH Escurrido
'
'  CÓMO USARLO
'    1. En Excel: Programador → Visual Basic (o Alt+F11).
'    2. Insertar → Módulo, y pegar todo este contenido.
'    3. Ejecutar UNA SOLA VEZ la macro  Agregar_Escurrido_A_Dinamicas
'    4. Reemplazar las macros Historico_Valvulas y Limpiar del Módulo1 por las
'       versiones de más abajo.
'
'  ANTES de ejecutar nada, agregar las dos columnas en las dos tablas:
'    · Hoja Reporte,  tabla Tabla1    → Z3 = "CE Escurrido",  AA3 = "pH Escurrido"
'    · Hoja Historico, tabla Historico → Z1 = "CE Escurrido",  AA1 = "pH Escurrido"
'  Al escribir el encabezado junto a la última columna, Excel extiende la tabla
'  sola. Las posiciones quedan alineadas entre las dos hojas, que es lo que
'  permite que Historico_Valvulas siga copiando un rango contiguo.
' =============================================================================

Option Explicit

Private Const CAMPO_CE As String = "CE Escurrido"
Private Const CAMPO_PH As String = "pH Escurrido"


' --- Agrega los dos campos a las 53 tablas dinámicas ------------------------
'  Las 53 comparten una sola caché, así que basta refrescarla una vez para que
'  los campos nuevos existan; pero cada tabla necesita que se le agreguen al
'  área de valores para que aparezcan como filas. Eso es lo que hace esto.
Sub Agregar_Escurrido_A_Dinamicas()

    Dim ws As Worksheet, pt As PivotTable
    Dim campos As Variant, nombre As Variant
    Dim agregados As Long, existentes As Long, sinCampo As Long
    Dim primera As Boolean

    campos = Array(CAMPO_CE, CAMPO_PH)
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

            For Each nombre In campos
                If TieneCampoDeDatos(pt, CStr(nombre)) Then
                    existentes = existentes + 1
                ElseIf Not ExisteCampo(pt, CStr(nombre)) Then
                    sinCampo = sinCampo + 1
                Else
                    With pt.PivotFields(CStr(nombre))
                        .Orientation = xlDataField
                        .Function = xlSum
                    End With
                    agregados = agregados + 1
                End If
            Next nombre

        Next pt
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Campos agregados: " & agregados & vbCrLf & _
          "Ya estaban:       " & existentes
    If sinCampo > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "ATENCIÓN: en " & sinCampo & " casos el campo no existe en la caché." & vbCrLf & _
              "Revisa que las columnas estén dentro de la tabla Historico y que" & vbCrLf & _
              "los encabezados digan exactamente '" & CAMPO_CE & "' y '" & CAMPO_PH & "'."
    End If
    MsgBox msg, vbInformation, "Escurrido en tablas dinámicas"

End Sub


' ¿La tabla dinámica ya muestra ese campo en el área de valores?
Private Function TieneCampoDeDatos(pt As PivotTable, nombre As String) As Boolean
    Dim df As PivotField
    For Each df In pt.DataFields
        If df.SourceName = nombre Then
            TieneCampoDeDatos = True
            Exit Function
        End If
    Next df
End Function


' ¿El campo existe en la caché, es decir, la columna llegó al origen de datos?
Private Function ExisteCampo(pt As PivotTable, nombre As String) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nombre)
    On Error GoTo 0
    ExisteCampo = Not pf Is Nothing
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
