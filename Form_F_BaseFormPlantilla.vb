Option Compare Database

Option Explicit


Private camposObligatorios As Variant
Private camposSiempreBloqueados As Variant
'los siguientes hay que definirlos Public para poder acceder a ellos desde modFormUtils
Public camposNumericosPositivos As Variant
Public camposFechaCompleta As Variant
Private tablaBase As String


Private Sub Form_Open(Cancel As Integer)


If Nz(Me.OpenArgs, "") = "Vacio" Then
Me.RecordSource = "SELECT * FROM " & tablaBase & " WHERE 1=2;"
End If


End Sub


Private Sub Form_Load()
'introducir la tabla principal que gestiona el formulario
tablaBase = ""

camposObligatorios = Array()
camposSiempreBloqueados = Array()
camposNumericosPositivos = Array()
camposFechaCompleta = Array()
Me.RecordSource = tablaBase
Me.Requery
If Me.RecordsetClone.RecordCount = 0 Then
    DoCmd.GoToRecord , , acNewRec
Else
    Call ModoLectura
End If

End Sub


Private Sub Form_Current()


If Me.NewRecord Then
    Call ModoEdicion
    'en algún campo puede que haya que personalizar el limpiado del formulario
    'desde esta opción habría que llamar a esa función
Else
    Call ModoLectura
    RellenarFormulario
End If


End Sub


Private Sub cmdNuevo_Click()


' Abrir vacío para nuevo
'El código comentado llama a Form_Current(), con lo que oculta los botones para guardar
' y luego da error en DoCmd.GoToRecord , , acNewRec
'Me.RecordSource = tablaBase
'Me.Requery

Call ModoEdicion
DoCmd.GoToRecord , , acNewRec





End Sub


Private Sub cmdEditar_Click()
    Call ModoEdicion
End Sub


Private Sub cmdGuardar_Click()


    If Not Form_Validar(Me, camposObligatorios) Then Exit Sub
    If Not Form_ValidarAvanzado(Me) Then Exit Sub
    
    DoCmd.RunCommand acCmdSaveRecord


    Msg_Info "Registro guardado."


    Call ModoLectura


End Sub


Private Sub cmdCancelar_Click()


    If Me.Dirty Then
        DoCmd.RunCommand acCmdUndo
    End If

    Call ModoLectura


End Sub
Private Sub cmdEliminar_Click()
On Error GoTo Err_cmdEliminar_Click


    'DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdEliminar_Click:
    Exit Sub

Err_cmdEliminar_Click:
    MsgBox Err.Description
    Resume Exit_cmdEliminar_Click
    
End Sub


Private Sub cmdBuscar_Click()


Dim r As Variant
'SUSTITUIR LOS CAMPOS ID POR EL CAMPO QUE CORRESPONDA
'PARÁMETROS
'TABLA-CAMPO ID-CAMPO TEXTO-CAMPOS A MOSTRAR-ANCHOS

r = AbrirBuscadorUniversal(tablaBase, "ID", "Nombre", "ID, Nombre", "0,100")


If r <> "" And Not IsNull(r) Then
    Me.RecordSource = "SELECT * FROM " & tablaBase & " WHERE ID = " & r & ";"
    Me.Requery
    Call ModoLectura
End If


End Sub
Private Sub cmdCerrar_Click()
DoCmd.Close acForm, Me.Name
End Sub


Public Sub RellenarFormulario()
' Cada formulario hijo debe implementar su propia lógica
End Sub
Private Sub cmdNextRegis_Click()
On Error GoTo Err_btn_NextRegis_Click


    DoCmd.GoToRecord , , acNext

Exit_btn_NextRegis_Click:
    Exit Sub

Err_btn_NextRegis_Click:
    MsgBox Err.Description
    Resume Exit_btn_NextRegis_Click
    
End Sub
Private Sub cmdPrevRegs_Click()
On Error GoTo Err_btn_PrevRegs_Click


    DoCmd.GoToRecord , , acPrevious

Exit_btn_PrevRegs_Click:
    Exit Sub

Err_btn_PrevRegs_Click:
    MsgBox Err.Description
    Resume Exit_btn_PrevRegs_Click
    
End Sub
Private Sub cmdFirstRegs_Click()
On Error GoTo Err_cmdFirstRegs_Click


    DoCmd.GoToRecord , , acFirst

Exit_cmdFirstRegs_Click:
    Exit Sub

Err_cmdFirstRegs_Click:
    MsgBox Err.Description
    Resume Exit_cmdFirstRegs_Click
    
End Sub
Private Sub cmdLastRegs_Click()
On Error GoTo Err_cmdLastRegs_Click


    DoCmd.GoToRecord , , acLast

Exit_cmdLastRegs_Click:
    Exit Sub

Err_cmdLastRegs_Click:
    MsgBox Err.Description
    Resume Exit_cmdLastRegs_Click
    
End Sub
Private Sub ModoLectura()

    Me.AllowEdits = False
    Me.AllowAdditions = False
    Me.AllowDeletions = False

    Form_Bloquear Me, True, camposSiempreBloqueados

    ' Gestión de botones
    Me.cmdEditar.Visible = True
    If Me.Recordset.RecordCount > 0 Then
        Me.cmdEditar.SetFocus
    End If
    Me.cmdGuardar.Visible = False
    Me.cmdCancelar.Visible = False

    Me.cmdNuevo.Enabled = True
    Me.cmdEliminar.Enabled = True
    Me.cmdBuscar.Enabled = True

End Sub
Private Sub ModoEdicion()

    Me.AllowEdits = True
    Me.AllowAdditions = True

    Form_Bloquear Me, False, camposSiempreBloqueados

    ' Botones
    Me.cmdGuardar.Visible = True
    Me.cmdGuardar.SetFocus
    Me.cmdCancelar.Visible = True
    Me.cmdEditar.Visible = False
       

    Me.cmdNuevo.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.cmdBuscar.Enabled = False

End Sub
