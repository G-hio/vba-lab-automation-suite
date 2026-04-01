' ******************************************************************************
' PROYECTO: Automatización de Procesos de Laboratorio - AVALQUIMICO SAS
' DESCRIPCIÓN: Normalización de siglas estandar e inserción automática
'              de parámetros para Análisis Proximal.
' AUTOR: Giancarlo - Ingeniero de Sistemas (8vo Semestre)
' ******************************************************************************

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' Optimización de rendimiento para procesos masivos
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' 1. MÓDULO: Normalización de Códigos en Columna G
    ' Detecta cambios en la columna de códigos y aplica formato estándar
    If Not Intersect(Target, Me.Columns("G")) Is Nothing Then
        Call NormalizarCodigosFacturacion(Me)
    End If
    
    ' 2. MÓDULO: Desglose de Análisis Proximal en Columna F
    ' Inserta filas automáticamente cuando se detecta "ANALISIS PROXIMAL"
    If Not Intersect(Target, Me.Columns("F")) Is Nothing Then
        Call DesglosarAnálisisProximal(Me)
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Exit Sub

ErrorHandler:
    MsgBox "Error en la automatización: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' --- Procedimiento para normalizar prefijos de laboratorio ---
Private Sub NormalizarCodigosFacturacion(ws As Worksheet)
    Dim ultimaFila As Long, i As Long
    Dim valor As String
    
    ultimaFila = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 1 To ultimaFila
        valor = UCase(Trim(ws.Cells(i, "G").Value))
        
        ' Lógica de prefijos estándar del laboratorio
        If valor Like "QH*" Then 
            ws.Cells(i, "G").Value = "QH"
        ElseIf valor Like "QICL*" Then 
            ws.Cells(i, "G").Value = "QICL"
        ElseIf valor Like "QICG*" Then 
            ws.Cells(i, "G").Value = "QICG"
        ElseIf valor Like "AAMT*" Then 
            ws.Cells(i, "G").Value = "AAMT"
        ElseIf valor Like "LEXT*" Then 
            ws.Cells(i, "G").Value = "LEXT"
        End If
    Next i
End Sub

' --- Procedimiento para generar sub-parámetros técnicos ---
Private Sub DesglosarAnálisisProximal(ws As Worksheet)
    Dim ultimaFila As Long, i As Long
    
    ultimaFila = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Recorrido inverso (Step -1) es fundamental al insertar filas
    For i = ultimaFila To 1 Step -1
        If InStr(UCase(Trim(ws.Cells(i, "F").Value)), "ANALISIS PROXIMAL") > 0 Then
            ' Inserta espacio para los 3 parámetros
            ws.Rows(i).Resize(3).Insert Shift:=xlShiftDown
            
            ' Copia el formato y datos de la fila original a las nuevas 3 filas
            ws.Rows(i + 3).Copy Destination:=ws.Rows(i).Resize(3)
            
            ' Asigna los nombres específicos de los análisis
            ws.Cells(i, "F").Value = "MATERIA GRASA TOTAL"
            ws.Cells(i + 1, "F").Value = "HUMEDAD"
            ws.Cells(i + 2, "F").Value = "CENIZAS"
        End If
    Next i
End Sub
