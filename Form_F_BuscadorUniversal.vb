Option Compare Database
Option Explicit

Private origenDatos As String
Private campoKey As String
Private campoMostrar As String
Private columnasMostrar As String
Private anchos As String


Private Sub cmdAceptar_Click()

    If IsNull(Me.lstResultados.Value) Then
        MsgBox "Seleccione un registro.", vbExclamation
        Exit Sub
    End If
    
    TempVars("ResultadoBuscador") = Me.lstResultados.Value
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdCancelar_Click()

    TempVars("ResultadoBuscador") = ""
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()

    'si el formulario se abre sin parámetros (p. ej. en diseño), no fallar
    On Error Resume Next
    origenDatos = TempVars("BU_tabla")
    campoKey = TempVars("BU_campoID")
    campoMostrar = TempVars("BU_campoTexto")
    columnasMostrar = TempVars("BU_camposMostrar")
    anchos = TempVars("BU_anchos")
    On Error GoTo 0

    'abrir buscador vacío (la búsqueda la dispara txtBuscar)
    Me.lstResultados.RowSource = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    TempVars.Remove "BU_tabla"
    TempVars.Remove "BU_campoID"
    TempVars.Remove "BU_campoTexto"
    TempVars.Remove "BU_camposMostrar"
    TempVars.Remove "BU_anchos"
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)

    If IsNull(Me.lstResultados.Value) Then
        MsgBox "Seleccione un registro.", vbExclamation
        Exit Sub
    End If
    
    TempVars("ResultadoBuscador") = Me.lstResultados.Value
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub lstResultados_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If IsNull(Me.lstResultados.Value) Then
        MsgBox "Seleccione un registro.", vbExclamation
        Exit Sub
    End If
    
    TempVars("ResultadoBuscador") = Me.lstResultados.Value
    DoCmd.Close acForm, Me.Name
End If
End Sub

Private Sub txtBuscar_AfterUpdate()
Dim sql As String
Dim criterio As String
Dim campos As Variant
Dim lanchos As Variant
Dim anc As Variant
Dim anchosfinal As String
criterio = Nz(Me.txtBuscar, "")

sql = "SELECT " & campoKey & ", " & columnasMostrar & _
      " FROM " & origenDatos

If criterio <> "" Then
    'escapar comodines y comillas simples para búsqueda literal
    criterio = Replace(criterio, "[", "[[]")
    criterio = Replace(criterio, "*", "[*]")
    criterio = Replace(criterio, "?", "[?]")
    criterio = Replace(criterio, "#", "[#]")
    criterio = Replace(criterio, "'", "''")

    sql = sql & " WHERE " & campoMostrar & _
          " LIKE '*" & criterio & "*'"
End If

sql = sql & " ORDER BY " & campoMostrar & ";"
campos = Split(columnasMostrar, ",")

Me.lstResultados.ColumnCount = UBound(campos) + 2 'el id no va en la lista aunque se muestra
lanchos = Split(anchos, ",")
'ancho 0 para la columna de la clave; los anchos recibidos van a las columnas visibles
anchosfinal = "0cm;"
For Each anc In lanchos
    anchosfinal = anchosfinal & anc / 100 * Me.lstResultados.Width / 1000 & "cm;"
Next
anchosfinal = Left(anchosfinal, Len(anchosfinal) - 1)
Me.lstResultados.ColumnWidths = anchosfinal
Me.lstResultados.RowSource = sql

End Sub

Private Sub txtBuscar_Change()
'búsqueda progresiva: refiltra la lista en cada pulsación
    txtBuscar_AfterUpdate
End Sub
