Option Compare Database

Option Explicit

'===============================================================
'  Módulo Global de Utilidades de Formularios  (modFormUtils)
'  Versión 1.0
'===============================================================

'---------------------------------------------------------------
'  1. LIMPIAR FORMULARIO
'---------------------------------------------------------------
Public Sub Form_Limpiar(frm As Form)

    Dim ctl As Control

    For Each ctl In frm.Controls
        Select Case ctl.ControlType
            Case acTextBox, acComboBox
            'he quitado los bloqueados porque creo que también hay que limpiarlos and not ctl.locked
            'el ID está locked y no se puede limpiar
                If ctl.Enabled And Not ctl.Name Like "txtID_*" Then
                    ctl.Value = Null
                End If
            Case acListBox
                ctl.Requery
        End Select
    Next ctl

End Sub

'==============================================================
'  BLOQUEAR O DESBLOQUEAR FORMULARIO
'==============================================================
'
'  bloquear = True  ? modo consulta (solo lectura)
'  bloquear = False ? modo edición
'
'  Los campos enumerados en frm.camposSiempreBloqueados
'  quedarán SIEMPRE Bloqueados (Locked=True, Enabled=True)
'
'==============================================================

Public Sub Form_Bloquear(frm As Form, bloquear As Boolean, ParamArray controles())

    Dim ctl As Control
    Dim i As Long
    Dim estaSiempreBloqueado As Boolean
    Dim arr() As Variant

    ' Obtener array interno correcto
    If UBound(controles) < 0 Then
        arr = Array()
    ElseIf IsArray(controles(0)) Then
        arr = controles(0)
    Else
        arr = controles
    End If
    For Each ctl In frm.Controls

        Select Case ctl.ControlType

            Case acTextBox, acComboBox, acListBox
                estaSiempreBloqueado = False

                'Comprobar si está en la lista de controles siempre bloqueados
                On Error Resume Next
                For i = LBound(arr) To UBound(arr)
                    If ctl.Name = arr(i) Then
                        estaSiempreBloqueado = True
                        Exit For
                    End If
                Next i
                On Error GoTo 0

                If bloquear Then
                    'Bloquear todo
                    ctl.Locked = True
                    ctl.Enabled = True

                Else
                    'Desbloquear salvo los siempre bloqueados
                    If estaSiempreBloqueado Then
                        ctl.Locked = True
                        ctl.Enabled = True
                    Else
                        ctl.Locked = False
                        ctl.Enabled = True
                    End If
                End If
    'pendiente pensar qué hacer con los commandButton
            'Case acCommandButton
                'Botones: deshabilitar en modo bloqueo
            '    If bloquear Then
            '        ctl.Enabled = True   ' Permitimos todos por defecto, se ajusta en cada formulario si se necesita
            '    Else
            '        ctl.Enabled = True
            '    End If

        End Select

    Next ctl

End Sub

'---------------------------------------------------------------
'  3. VALIDAR CAMPOS OBLIGATORIOS
'---------------------------------------------------------------
Public Function Form_Validar(frm As Form, ParamArray controles()) As Boolean

    Dim i As Long
    Dim Nombre As String
    Dim ctl As Control
    Dim arr() As Variant

    ' Obtener array interno correcto
    If UBound(controles) < 0 Then
        arr = Array()
    ElseIf IsArray(controles(0)) Then
        arr = controles(0)
    Else
        arr = controles
    End If

    For i = LBound(arr) To UBound(arr)

        Nombre = arr(i)
        On Error Resume Next
        Set ctl = frm.Controls(Nombre)
        If ctl Is Nothing Then
            On Error GoTo 0
            MsgBox "El control '" & Nombre & "' no se encuentra en el formulario.", vbCritical
            Form_Validar = False
            Exit Function
        End If
        On Error GoTo 0
        
        If Nz(ctl.Value, "") = "" Then
        'If IsNull(ctl) Or ctl.Value = "" Then
            MsgBox "El campo '" & ctl.Name & "' es obligatorio.", vbExclamation
            ctl.SetFocus
            Form_Validar = False
            Exit Function
        End If

    Next i

    Form_Validar = True

End Function
Public Function Form_ValidarAvanzado(frm As Form) As Boolean
    Dim campo As Variant
    Dim ctl As Control

    ' Validar números positivos
    'On Error Resume Next
    If Not IsEmpty(frm.camposNumericosPositivos) Then
        For Each campo In frm.camposNumericosPositivos
            Set ctl = frm.Controls(campo)
            If Not Val_NumeroPositivo(ctl.Value) Then
                Msg_Error "El campo '" & ctl.Name & "' debe ser un número mayor que 0."
                ctl.SetFocus
                Form_ValidarAvanzado = False
                Exit Function
            End If
        Next campo
    End If

    ' Validar fechas completas
    If Not IsEmpty(frm.camposFechaCompleta) Then
        For Each campo In frm.camposFechaCompleta
            Set ctl = frm.Controls(campo)
            If Not Val_FechaCompleta(ctl.Value) Then
                Msg_Error "El campo '" & ctl.Name & "' debe incluir fecha y hora completas."
                ctl.SetFocus
                Form_ValidarAvanzado = False
                Exit Function
            End If
        Next campo
    End If

    Form_ValidarAvanzado = True
End Function

'---------------------------------------------------------------
'  4. CARGAR COMBO DEPENDIENTE
'---------------------------------------------------------------
Public Sub Form_CargarComboDependiente(cboPadre As ComboBox, cboHijo As ComboBox, _
                                       tabla As String, campoIDPadre As String, campoID As String, campoTexto As String)

    If Nz(cboPadre.Value, "") = "" Then
        cboHijo.RowSource = ""
        cboHijo.Requery
        Exit Sub
    End If

    cboHijo.RowSource = "SELECT " & campoID & ", " & campoTexto & _
                        " FROM " & tabla & _
                        " WHERE " & campoIDPadre & " = " & cboPadre.Value & _
                        " ORDER BY " & campoTexto & ";"

    cboHijo.Requery

End Sub

Public Sub Form_CargarComboTabla(cbo As ComboBox, tabla As String, campoID As String, campoTexto As String)

    cbo.RowSource = "SELECT " & campoID & ", " & campoTexto & _
                    " FROM " & tabla & _
                    " ORDER BY " & campoTexto & ";"

    cbo.Requery
    'Modificado 20/09/26 para visualizar el dato y ocultar el id
    cbo.ColumnCount = 2
    cbo.ColumnWidths = "0cm;2cm"
End Sub
Public Sub Form_CargaTextBoxTabla(txt As TextBox, tabla As String, campoID As String, valorID As Long, campoTexto As String)
Dim valor As String

valor = Nz(DLookup(campoTexto, tabla, campoID & "=" & valorID), "")
txt = valor
End Sub
'---------------------------------------------------------------
'  5. COMPROBAR SI UN COMBO TIENE REGISTROS
'---------------------------------------------------------------
Public Function Combo_TieneDatos(cbo As ComboBox) As Boolean

    If cbo.ListCount > 0 Then
        Combo_TieneDatos = True
    Else
        Combo_TieneDatos = False
    End If

End Function

'---------------------------------------------------------------
'  6. MOSTRAR MENSAJES COMUNES
'---------------------------------------------------------------


Public Sub Msg_Advertencia(texto As String)
    MsgBox texto, vbExclamation
End Sub

Public Function Msg_Confirmacion(texto As String) As Boolean
    Msg_Confirmacion = (MsgBox(texto, vbYesNo + vbQuestion) = vbYes)
End Function
' Mensajes estándar
'-----------------------------
Public Sub Msg_Info(txt As String)
MsgBox txt, vbInformation, "Información"
End Sub
Public Sub Msg_Error(txt As String)
MsgBox txt, vbCritical, "Error"
End Sub
Public Function Msg_Pregunta(txt As String) As Boolean
Msg_Pregunta = (MsgBox(txt, vbQuestion + vbYesNo, "Confirmar") = vbYes)
End Function
'---------------------------------------------------------------
'  7. ABRIR FORMULARIO BUSCADOR (MODAL)
'---------------------------------------------------------------
Public Function Form_AbrirBuscador(nombreBuscador As String) As Variant

    DoCmd.OpenForm nombreBuscador, , , , , acDialog

    'IMPORTANTE: el formulario modal debe establecer un valor antes de cerrar.
    'Este valor queda disponible en esta función:
    Form_AbrirBuscador = TempVars("ResultadoBuscador")

End Function
Public Function AbrirBuscadorUniversal(tabla As String, campoID As String, _
                                       campoTexto As String, camposMostrar As String, anchos As String) As Variant
    'control básico de errores
    If tabla = "" Or campoID = "" Or campoTexto = "" Or camposMostrar = "" Or anchos = "" Then
        MsgBox "Parámetros del buscador incompletos.", vbCritical
        Exit Function
    End If
    
    'Parametros para el buscador
    TempVars("BU_tabla") = tabla
    TempVars("BU_campoID") = campoID
    TempVars("BU_campoTexto") = campoTexto
    TempVars("BU_camposMostrar") = camposMostrar
    TempVars("BU_anchos") = anchos

    TempVars("ResultadoBuscador") = ""

    'Ahora sí: abrir modal (y aquí se detiene la ejecución)
    DoCmd.OpenForm "F_BuscadorUniversal", , , , , acDialog

    'Se ejecutará al CERRAR el formulario
    AbrirBuscadorUniversal = TempVars("ResultadoBuscador")

End Function


'---------------------------------------------------------------
'  8. NAVEGACIÓN ENTRE REGISTROS
'---------------------------------------------------------------
Public Sub Form_Siguiente(frm As Form)
    DoCmd.GoToRecord frm.Name, , acNext
End Sub

Public Sub Form_Anterior(frm As Form)
    DoCmd.GoToRecord frm.Name, , acPrevious
End Sub

'---------------------------------------------------------------
'  9. ABRIR FORMULARIO EN UN REGISTRO ESPECÍFICO
'---------------------------------------------------------------
Public Sub Form_AbrirEnRegistro(nombreForm As String, campoClave As String, valorClave As Variant)

    DoCmd.OpenForm nombreForm, , , campoClave & "=" & valorClave

End Sub

'==============================================================
'  CARGAR CAMPOS RELACIONADOS BASÁNDOSE EN UN ID
'==============================================================
'
'   Uso típico:
'
'   Call Form_CargarCamposRelacionados( _
'          "SELECT Equipo, Averia FROM vIntervenciones WHERE ID = " & Me.ID, _
'          Array(Me.txtEquipo, Me.txtAveria) _
'       )
'
'==============================================================

Public Sub Form_CargarCamposRelacionados(sql As String, controles() As Control)
    Dim rs As DAO.Recordset
    Dim i As Long

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst

        For i = LBound(controles) To UBound(controles)
            If Not IsNull(rs(i)) Then
                controles(i).Value = rs(i)
            Else
                controles(i).Value = Null
            End If
        Next i
    Else
        'sin resultados: limpiar para no dejar valores obsoletos
        For i = LBound(controles) To UBound(controles)
            controles(i).Value = Null
        Next i
    End If

    rs.Close
    Set rs = Nothing
End Sub

Public Sub Form_CargarCamposRelacionadosEnCadena(frm As Form, _
        tablaOrigen As String, campoClaveOrigen As String, valorClave As Variant, _
        relaciones As Variant)

    Dim rsOrigen As DAO.Recordset
    Dim rsDestino As DAO.Recordset
    Dim sql As String
    Dim i As Long

    '1. Obtener el registro en la tabla origen (Mantenimientos)
    sql = "SELECT * FROM " & tablaOrigen & " WHERE " & campoClaveOrigen & "=" & valorClave
    Set rsOrigen = CurrentDb.OpenRecordset(sql)

    If rsOrigen.EOF Then
        rsOrigen.Close
        Exit Sub
    End If

    '2. Recorrer cada relación
    For i = LBound(relaciones) To UBound(relaciones)

        Dim campoID As String
        Dim tablaDestino As String
        Dim campoClaveDestino As String
        Dim controlFormulario As String
        Dim campoTextoDestino As String
        Dim valorRelacion As Variant

        campoID = relaciones(i)(0)
        tablaDestino = relaciones(i)(1)
        campoClaveDestino = relaciones(i)(2)
        controlFormulario = relaciones(i)(3)
        campoTextoDestino = relaciones(i)(4)

        valorRelacion = rsOrigen.Fields(campoID).Value

        If Not IsNull(valorRelacion) Then

            'Consultar tabla destino
            sql = "SELECT " & campoTextoDestino & _
                  " FROM " & tablaDestino & _
                  " WHERE " & campoClaveDestino & "=" & valorRelacion

            Set rsDestino = CurrentDb.OpenRecordset(sql)

            If Not rsDestino.EOF Then
                frm.Controls(controlFormulario).Value = rsDestino.Fields(0).Value
            Else
                frm.Controls(controlFormulario).Value = Null
            End If

            rsDestino.Close

        Else
            frm.Controls(controlFormulario).Value = Null
        End If

    Next i

    rsOrigen.Close

End Sub

Public Function Val_LongitudMinima(valor As Variant, minimo As Integer) As Boolean
    Val_LongitudMinima = (Len(Nz(valor, "")) >= minimo)
End Function
Public Function Val_HoraCompleta(valor As Variant) As Boolean
    Dim h As Integer, m As Integer

    'vacío: se considera válido; la obligatoriedad la decide Form_Validar
    If Nz(valor, "") = "" Then
        Val_HoraCompleta = True
        Exit Function
    End If

    If Not IsDate("01/01/2000 " & Nz(valor, "")) Then
        Val_HoraCompleta = False
        Exit Function
    End If

    h = Hour("01/01/2000 " & valor)
    m = Minute("01/01/2000 " & valor)

    Val_HoraCompleta = (h >= 0 And h <= 23 And m >= 0 And m <= 59)
End Function
Public Function Val_FechaCompleta(valor As Variant) As Boolean

    'vacío: se considera válido; la obligatoriedad la decide Form_Validar
    If Nz(valor, "") = "" Then
        Val_FechaCompleta = True
        Exit Function
    End If

    If Not IsDate(Nz(valor, "")) Then
        Val_FechaCompleta = False
        Exit Function
    End If

    ' Si la hora es cero ? falta información de hora/minuto
    If Hour(valor) = 0 And Minute(valor) = 0 Then
        Val_FechaCompleta = False
        Exit Function
    End If

    Val_FechaCompleta = True

End Function
Public Function Val_NumeroPositivo(valor As Variant) As Boolean

    If Not IsNumeric(Nz(valor, "")) Then
        Val_NumeroPositivo = False
        Exit Function
    End If

    Val_NumeroPositivo = (CDbl(valor) > 0)

End Function

Public Function Val_EsNumero(valor As Variant) As Boolean
    Val_EsNumero = (IsNumeric(Nz(valor, "")))
End Function



