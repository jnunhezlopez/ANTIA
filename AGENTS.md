# ANTIA — Índice del proyecto

Aplicación Microsoft Access (VBA) para la gestión de inventario de equipos.
ENTORNO Windows / PowerShell 5.1. El repositorio NO es un proyecto git.

## Estructura del repositorio

| Archivo | Descripción |
|---|---|
| `ANTIA_two.accdb` | Base de datos principal (toda la lógica vive aquí). Contiene imágenes/objetos y código VBA. |
| `ANTIA_two.laccdb` | Archivo de bloqueo de Access (residual; ignorar, no versionar). |
| `Form_F_BaseFormPlantilla.vb` | Export del módulo del formulario base `F_BaseFormPlantilla` (leer/escribir siempre fuera del accdb). |
| `Form_F_BuscadorUniversal.vb` | Export del módulo del formulario `F_BuscadorUniversal`. |
| `modFormUtils.vb` | Export del módulo estándar `modFormUtils`. |
| `AGENTS.md` | Este índice. |
| `.gitignore` | Excluye `*.accdb` (binario de Access) y `*.laccdb` (archivo de bloqueo). Solo se versiona el código fuente. |

Regla crítica: el .vb es SOLO una copia de trabajo. Cualquier cambio debe
re-importarse en el accdb (Access) para tener efecto; y al exportar código
desde Access hay que volcar el resultado a estos .vb para que quede sincronizado.

## Estado lógico del accdb (leído con ACE OLEDB + Access COM)

Objetos actualmente presentes:
- Formularios (2): `F_BaseFormPlantilla`, `F_BuscadorUniversal`
- Módulos estándar: ninguno en el proyecto VBA (la entrada `modFormUtils` del panel de navegación es un fantasma; el código real está exportado en `modFormUtils.vb`)
- Tablas de usuario: NINGUNA (solo las 12 `MSys*` de sistema). El esquema hay que reconstruirlo.

## Esquema previsto (congelado en `MSysNameMap` del accdb, procedente de la v1)

El accdb guarda el mapeo de nombres de los objetos borrados. Con él se reconstruye el esquema objetivo:

- `T_Categorias` — ID (PK), Nombre
- `T_Subcategorias` — ID (PK), IDCategoria (FK → T_Categorias.ID), Nombre
- `T_Tipos` — ID_Tipo (PK), Nombre
- `T_Ubicaciones` — ID_Ubicacion (PK), Nombre
- `T_Fabricantes` — ID_Fabricante (PK), Nombre
- `T_EstadosEQ` — ID_EstadosEQ (PK), Nombre
- `T_Equipos` — ID_Equipo (PK), Nombre, ID_Tipo (FK), ID_Ubicacion (FK), ID_Fabricante (FK), Modelo, NumeroSerie, FechaInstalacion, ID_EstadoEQ (FK), FechaCompra, DatosTecnicos, Observaciones
- `T_Prueba` — Nombre, Observaciones

Formularios heredados de la v1: `F_BuscadorResultados` (record source `T_Categorias`), `F_Fabricantes` (`T_Fabricantes`), `F_Equipos` (`T_Equipos`).

Query fantasma heredada: `~sq_cF_BuscadorUniversal_old_subform~sq_csubResultados` → `SELECT DISTINCTROW * FROM T_Categorias;` (pertenecía a un subformulario antiguo ya suprimido del buscador).

## Convenciones de código

- Idioma: castellano (mensajes de UI y comentarios).
- Prefijos de objetos: tablas `T_`, formularios `F_`, controles `txt` (TextBox), `cbo` (ComboBox), `lst` (ListBox), `cmd` (CommandButton), `subf` (SubForm).
- Claves: `ID_*` (enteros Autonumérico/PK). Los FKs se nombran igual que la PK de la tabla destino (`ID_Tipo`, `ID_Ubicacion`, …).
- `Option Compare Database` + `Option Explicit` al inicio de cada módulo.
- Los formularios de edición heredan (copian) el patrón del formulario base `F_BaseFormPlantilla`; no existe herencia real en VBA.

## Índice de código

### `Form_F_BaseFormPlantilla.vb` — formulario base/plantilla (modo lectura–edición)
Variables de configuración que cada formulario hijo sobreescribe en `Form_Load`:
- `tablaBase` (String): tabla gestionada por el formulario.
- `camposObligatorios`, `camposSiempreBloqueados` (Private Variant array).
- `camposNumericosPositivos`, `camposFechaCompleta` (Public Variant array; Public para que `modFormUtils` los lea).

Procedimientos:
- `Form_Open` — stub; la apertura con `OpenArgs="Vacio"` se resuelve en `Form_Load` (en `Open`, `tablaBase` aún no está inicializado).
- `Form_Load` — inicializa arrays y RecordSource; con `OpenArgs="Vacio"` abre en 0 registros (`WHERE 1=2`); si no hay registros va a nuevo, si no → `ModoLectura`.
- `Form_Current` — registro nuevo → `ModoEdicion`; si no → `ModoLectura` + `RellenarFormulario`.
- `Form_BeforeUpdate` — valida antes de guardar cualquier vía (PgDn/PgUp, botones de navegación, `cmdGuardar`, cierre); si la validación falla `Cancel = True` → el registro no se guarda y se permanece en él.
- `Form_KeyDown` — con `KeyPreview=True` (activado en `Form_Load`) bloquea PgDn/PgUp (`vbKeyNext`/`vbKeyPrior`) mientras se está en modo edición (`AllowEdits=True`), para que no abandonen el registro a mitad de la introducción; en modo lectura la navegación por teclado sigue funcionando. Ojo: `KeyPreview` es una propiedad de diseño, NO se hereda — hay que ponerla a `True` en cada formulario hijo (propiedad u hoja "Key Preview = Yes") y copiar también `Form_KeyDown` a su módulo.
- `RellenarFormulario` — stub público; cada formulario hijo implementa su lógica (cargas de componentes relacionados).
- `cmdNuevo_Click` / `cmdEditar_Click` / `cmdGuardar_Click` / `cmdCancelar_Click` / `cmdEliminar_Click` / `cmdBuscar_Click` / `cmdCerrar_Click`.
- `cmdCancelar_Click` — deshace los cambios (`Undo`); si queda en un registro nuevo sin guardar salta a uno existente.
- `cmdEliminar_Click` — pide confirmación (`Msg_Confirmacion`) y habilita temporalmente `AllowDeletions` (en `ModoLectura` está a `False`).
- `cmdBuscar_Click` — usa `AbrirBuscadorUniversal(tablaBase, "ID", "Nombre", "ID, Nombre", "0,100")`; ojo: asume campo "ID" y "Nombre" (adaptar al esquema objetivo con `ID_*`).
- Navegación: `cmdNextRegis_Click`, `cmdPrevRegs_Click`, `cmdFirstRegs_Click`, `cmdLastRegs_Click`.
- `ModoLectura` — Form.AllowEdits/Additions/Deletions=False, `Form_Bloquear Me, True`, muestra botones de edición/nuevo/eliminar/buscar.
- `ModoEdicion` — AllowEdits/Additions/Deletions=True, desbloquea, muestra Guardar/Cancelar, deshabilita Nuevo/Eliminar/Buscar.

### `Form_F_BuscadorUniversal.vb` — buscador universal (modal, vía TempVars)
Recibe parámetros desde `modFormUtils.AbrirBuscadorUniversal` mediante TempVars:
`BU_tabla`, `BU_campoID`, `BU_campoTexto`, `BU_camposMostrar`, `BU_anchos`; devuelve el ID seleccionado en `TempVars("ResultadoBuscador")`.
- `Form_Load` — lee los TempVars (con protección si no existen); RowSource vacío a propósito (la búsqueda la dispara `txtBuscar`).
- `txtBuscar_Change` — búsqueda progresiva: relanza el filtrado en cada pulsación.
- `txtBuscar_AfterUpdate` — SQL con `LIKE '*…*'`, escapa comillas simples y comodines (`* ? # [ ]`), calcula ColumnCount/ColumnWidths desde `camposMostrar`/`anchos` (la columna de la clave, primera, se oculta a 0 cm).
- `cmdAceptar_Click`, `lstResultados_DblClick`, `lstResultados_KeyDown` — fijan `ResultadoBuscador` y cierran; `cmdCancelar_Click` anula (deja `ResultadoBuscador` vacío).
- `Form_Unload` — limpia los TempVars `BU_*` (NO `ResultadoBuscador`, que lee el llamador tras el modal).
- Pendiente: falta aplicar el desbloqueo del `lstResultados`.

### `modFormUtils.vb` — utilidades globales de formularios
- `Form_Limpiar(frm)` — limpia TextBox/ComboBox (evita controles `txtID_*`), Requery de ListBox.
- `Form_Bloquear(frm, bloquear, camposSiempreBloqueados)` — bloquea/desbloquea TextBox/ComboBox/ListBox; los "siempre bloqueados" quedan `Locked=True, Enabled=True`.
- `Form_Validar(frm, camposObligatorios)` — valida campos no vacíos; si un control no existe avisa y no falla.
- `Form_ValidarAvanzado(frm)` — valida `camposNumericosPositivos` y `camposFechaCompleta` (usa `Val_*`).
- Combos: `Form_CargarComboDependiente`, `Form_CargarComboTabla` (ColumnCount=2, ancho "0cm;2cm"), `Form_CargaTextBoxTabla`.
- `Combo_TieneDatos(cbo)`.
- Mensajes: `Msg_Info`, `Msg_Error`, `Msg_Advertencia`, `Msg_Confirmacion`, `Msg_Pregunta`.
- Buscador: `Form_AbrirBuscador(nombreBuscador)`, `AbrirBuscadorUniversal(tabla, campoID, campoTexto, camposMostrar, anchos)`.
- Navegación: `Form_Siguiente`, `Form_Anterior`.
- `Form_AbrirEnRegistro(form, campoClave, valorClave)`.
- Campos relacionados: `Form_CargarCamposRelacionados(sql, controles())`, `Form_CargarCamposRelacionadosEnCadena(frm, tablaOrigen, …)`.
- Validadores: `Val_LongitudMinima`, `Val_HoraCompleta`, `Val_FechaCompleta`, `Val_NumeroPositivo`, `Val_EsNumero` (`Val_HoraCompleta` y `Val_FechaCompleta` tratan el vacío como válido; la obligatoriedad la decide `Form_Validar`).

## Notas de mantenimiento para agentes

- No hay build/lint/tests: el proyecto se valida abriendo el accdb en Microsoft Access.
- Para inspeccionar el accdb sin abrir Access: ACE OLEDB vía PowerShell
  (`Provider=Microsoft.ACE.OLEDB.12.0;Data Source=…`). Tablas visibles vía `OpenSchema(20)`.
- Para leer/escribir VBA real o la estructura interna: Access COM
  (`New-Object -ComObject Access.Application`, `$db.Modules`, `$db.Containers("Forms").Documents`).
- `.gitignore` ya excluye `*.accdb` y `*.laccdb`; el repositorio versiona solo el código fuente (`.vb`, `AGENTS.md`).
- Antes de reconstruir el esquema, confirmar con el usuario los nombres exactos de tablas/campos del esquema objetivo (columna "Esquema previsto").