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

    TempVars!Buscador_Seleccion = Null
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_Load()
Dim sql As String

    
    origenDatos = TempVars("BU_tabla")
    campoKey = TempVars("BU_campoID")
    campoMostrar = TempVars("BU_campoTexto")
    columnasMostrar = TempVars("BU_camposMostrar")
    anchos = TempVars("BU_anchos")
 'abrir buscador vacío
    Me.lstResultados.RowSource = ""

'    sql = "SELECT " & campoKey & ", " & columnasMostrar & _
          " FROM " & origenDatos & _
        " ORDER BY " & campoMostrar & ";"

'    Me.lstResultados.RowSource = sql
    
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
    sql = sql & " WHERE " & campoMostrar & _
          " LIKE '*" & Replace(criterio, "'", "''") & "*'"
End If

sql = sql & " ORDER BY " & campoMostrar & ";"
campos = Split(columnasMostrar, ",")

Me.lstResultados.ColumnCount = UBound(campos) + 2 'el id no va en la lista aunque se muestra
lanchos = Split(anchos, ",")
anchosfinal = ""
For Each anc In lanchos
    anchosfinal = anchosfinal & anc / 100 * Me.lstResultados.Width / 1000 & "cm;"
Next
anchosfinal = Left(anchosfinal, Len(anchosfinal) - 1)
Me.lstResultados.ColumnWidths = anchosfinal
Me.lstResultados.RowSource = sql

End Sub
