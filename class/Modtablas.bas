Attribute VB_Name = "ModificarCampos"
Option Explicit

Public VGCNx As New ADODB.Connection             'Conexion de la BD empresa
Public VGcnxCT As New ADODB.Connection        'Conexion de Contabilidad
Public VGGeneral As New ADODB.Connection      'Conexion de la BD Generales
Public VGConfig As New ADODB.Connection      'Conexion de la BD de configuracion
Public VGCadenaReport2 As String

Public VGnumniveles As Integer               'N�mero de Niveles del Plan de Cuentas
Public VGnumnivgas As Integer               'N�mero de Niveles del Plan de gastos
Public VGnumnivcos As Integer               'N�mero de Niveles de centro de costos

Public VGUsuario As String
Public VGPass  As String
Public VGcomputer As String                  'Nombre de la computadora
Public VGtipolicencia As String
Public VGfechalicencia As Date
Public VGCodEmpresa As String
Public SQL As String
Public RSQL As New ADODB.Recordset


Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Enum TIPOSISTEMA
   INVENTARIOS = 1
   compras = 2
   pagar = 3
   caja = 4
   contab = 5
   facturacion = 6
   cobrar = 7
   activos = 8
   costos = 9
   planillas = 10
   PyMe = 11
   
End Enum
Public VGsql As String * 1

Public Const NUMMAGICO As Integer = 5

'Constantes de mensajes para visualizar
Public mensaje1 As String
Public Const g_tiposol = "01"
Public Const g_tipodolar = "02"
Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"

Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum
Public Enum tipocambio
    Compra = "01"
    Venta = "02"
    Promedio = "03"
End Enum
Public Sub adicionarcamposinmuebles()

If Not ExisteElem(1, VGCNx, "maeart", "longitudderecha") Then
        VGCNx.Execute "ALTER TABLE maeart ADD longitudderecha float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "longitudizquierda") Then
        VGCNx.Execute "ALTER TABLE maeart ADD longitudizquierda float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "longitudfrontal") Then
        VGCNx.Execute "ALTER TABLE maeart ADD longitudfrontal float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "longitudposterior") Then
        VGCNx.Execute "ALTER TABLE maeart ADD longitudposterior float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "areaterreno") Then
        VGCNx.Execute "ALTER TABLE maeart ADD areaterreno float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "areaconstruida") Then
        VGCNx.Execute "ALTER TABLE maeart ADD areaconstruida float NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "numerodepisos") Then
        VGCNx.Execute "ALTER TABLE maeart ADD numerodepisos integer NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "numerodehabitaciones") Then
        VGCNx.Execute "ALTER TABLE maeart ADD numerodehabitaciones integer NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "numerodeservicios") Then
        VGCNx.Execute "ALTER TABLE maeart ADD numerodeservicios integer NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "linderofrontera") Then
        VGCNx.Execute "ALTER TABLE maeart ADD linderofrontera nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "linderoposterior") Then
        VGCNx.Execute "ALTER TABLE maeart ADD linderoposterior nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "linderoizquierdo") Then
        VGCNx.Execute "ALTER TABLE maeart ADD linderoizquierdo nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "linderoderecho") Then
        VGCNx.Execute "ALTER TABLE maeart ADD linderoderecho nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGCNx, "maeart", "proyectocodigo") Then
        VGCNx.Execute "ALTER TABLE maeart ADD proyectocodigo nvarchar(3) NULL"
End If

End Sub
Public Sub adicionarcamposCT()
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresaruc") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresaruc nvarchar(11) NULL"
   End If
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresadireccion") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresadireccion nvarchar(50) NULL"
   End If
    If Not ExisteElem(1, VGCNx, "co_multiempresas", "cajacodigo") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD cajacodigo varchar(50) default('01')"
   End If
   If Not ExisteElem(1, VGCNx, "ct_operacion", "facturacionanticipada") Then
        VGCNx.Execute "ALTER TABLE ct_operacion ADD facturacionanticipada bit default('0')"
   End If
    If Not ExisteElem(1, VGCNx, "ct_centrocosto", "estructuranumerolinea") Then
        VGCNx.Execute "ALTER TABLE ct_centrocosto ADD estructuranumerolinea varchar(10) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.AnoProceso & "", "saldoacumdebe00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.AnoProceso & " ADD saldoacumdebe00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.AnoProceso & "", "saldoacumhaber00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.AnoProceso & " ADD saldoacumhaber00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.AnoProceso & "", "saldoacumussdebe00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.AnoProceso & " ADD saldoacumussdebe00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.AnoProceso & "", "saldoacumussHaber00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.AnoProceso & " ADD saldoacumussHaber00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_cuenta", "cuentaadicionacargo") Then
        VGCNx.Execute "ALTER TABLE ct_cuenta ADD cuentaadicionacargo char(1) default ('0') "
   End If    'JCGI
   If Not ExisteElem(1, VGCNx, "ct_asiento", "asientoadicionacargo") Then
        VGCNx.Execute "ALTER TABLE ct_asiento ADD asientoadicionacargo char(1) default ('0') "
   End If
   If Not ExisteElem(1, VGCNx, "vt_asientodet", "cuentaventadiferida") Then
        VGCNx.Execute "ALTER TABLE vt_asientodet ADD cuentaventadiferida varchar(20) default ('00') "
        VGCNx.Execute (" update vt_asientodet set cuentaventadiferida=cuenta ")
   End If
   If Not ExisteElem(1, VGCNx, "vt_pedido", "pedidoventadiferida") Then
        VGCNx.Execute "ALTER TABLE vt_pedido ADD pedidoventadiferida integer  default (0) "
        VGCNx.Execute (" update vt_pedido set pedidoventadiferida=0 ")
   End If    'JCGI
   If Not ExisteElem(1, VGCNx, "ct_importarventas", "procedimientoasiento") Then
        VGCNx.Execute "ALTER TABLE ct_importarventas ADD procedimientoasiento varchar(40) default ('') "
        VGCNx.Execute (" update ct_importarventas set procedimientoasiento='' ")
   End If    'JCGI

End Sub
Public Sub adicionarcamposcostos()
   If Not ExisteElem(1, VGCNx, "cs_sistema", "baseorigen") Then
        VGCNx.Execute "ALTER TABLE cs_sistema ADD baseorigen varchar(30) default(' ')"
   End If
   If Not ExisteElem(1, VGCNx, "cs_resumenxmesplantillas", "importedolares") Then
        VGCNx.Execute "ALTER TABLE cs_resumenxmesplantillas ADD importedolares float default('0')"
   End If
   If Not ExisteElem(1, VGCNx, "cs_sistema", "codigopersonalplantilla") Then
        VGCNx.Execute "ALTER TABLE cs_sistema ADD codigopersonalplantilla varchar(2) default('00')"
   End If
   If Not ExisteElem(1, VGCNx, "cs_sistema", "mesesreferencia") Then
      VGCNx.Execute "ALTER TABLE cs_sistema ADD mesesreferencia integer default('12')"
  End If
  If Not ExisteElem(1, VGCNx, "cs_estructurapresentacion", "tipodegastosfijos") Then
        VGCNx.Execute "ALTER TABLE cs_estructurapresentacion ADD tipodegastosfijos bit default('0') "
 End If
If Not ExisteElem(1, VGCNx, "cs_sistema", "mesdecierre") Then
        VGCNx.Execute "ALTER TABLE cs_sistema ADD mesdecierre nvarchar(6) default('') "
End If
End Sub
Public Sub adicionarcampos()
On Error GoTo err2
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresaruc") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresaruc nvarchar(11) NULL"
   End If
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresadireccion") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresadireccion nvarchar(50) NULL"
   End If
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "codigocuenta") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD codigocuenta nvarchar(14) NULL"
   End If
   If ExisteElem(1, VGCNx, "cc_tipodocumento", "tdocumentonumerador") Then
        VGCNx.Execute "ALTER TABLE cc_tipodocumento ALTER COLUMN tdocumentonumerador nvarchar(15) NULL"
   End If
   If Not ExisteElem(1, VGCNx, "te_codigocaja", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE te_codigocaja ADD empresacodigo varchar(2) default('01')"
   End If
   If Not ExisteElem(1, VGCNx, "vt_cargo", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE vt_cargo ADD empresacodigo varchar(2) default('01')"
   End If
   If Not ExisteElem(1, VGCNx, "vt_abono", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE vt_abono ADD empresacodigo varchar(2) default('01')"
   End If
   If Not ExisteElem(1, VGCNx, "vt_puntovtadocumento", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE vt_puntovtadocumento ADD empresacodigo varchar(2) default('01')"
   End If
    If Not ExisteElem(1, VGCNx, "vt_seriedocumento", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE vt_seriedocumento ADD empresacodigo varchar(2) default('01')"
   End If
    If Not ExisteElem(1, VGCNx, "te_saldosmensuales", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE te_saldosmensuales ADD empresacodigo varchar(2) default('01')"
   End If
    If Not ExisteElem(1, VGCNx, "co_multiempresas", "cajacodigo") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD cajacodigo varchar(50) default('01')"
   End If
    If Not ExisteElem(1, VGCNx, "ct_operacion", "facturacionanticipada") Then
        VGCNx.Execute "ALTER TABLE ct_operacion ADD facturacionanticipada bit default('0')"
   End If
    If Not ExisteElem(1, VGCNx, "ct_centrocosto", "estructuranumerolinea") Then
        VGCNx.Execute "ALTER TABLE ct_centrocosto ADD estructuranumerolinea varchar(10) "
   End If
    If Not ExisteElem(1, VGCNx, "co_tipocompra", "modosprovisionescodigo") Then
        VGCNx.Execute "ALTER TABLE co_tipocompra ADD modosprovisionescodigo varchar(30) default('01,05')"
   End If
   If Not ExisteElem(1, VGCNx, "al_sistema", "flagconversioncodigo") Then
        VGCNx.Execute "ALTER TABLE al_sistema ADD flagconversioncodigo bit default('0')"
   End If
If Not ExisteElem(0, VGCNx, "al_tipoalmacen") Then
   SQL = " Create Table al_tipoalmacen "
   SQL = SQL & "( tipoalmacencodigo VarChar(1),"
   SQL = SQL & "tipoalmacendescripcion VarChar(30),"
   SQL = SQL & "usuariocodigo varchar(8),fechaact datetime "
   SQL = SQL & " CONSTRAINT PK_al_tipoalmacen "
   SQL = SQL & " PRIMARY KEY (tipoalmacencodigo))  "
   VGCNx.Execute SQL
End If
If Not ExisteElem(1, VGCNx, "al_sistema", "flagtipoalmacen") Then
        VGCNx.Execute "ALTER TABLE al_sistema ADD flagtipoalmacen bit default('0')"
End If
If Not ExisteElem(1, VGCNx, "tabalm", "tipoalmacencodigo") Then
        VGCNx.Execute "ALTER TABLE tabalm ADD tipoalmacencodigo varchar(1) default('0')"
End If
If Not ExisteElem(1, VGCNx, "co_gastos", "gastosgeneractacte") Then
        VGCNx.Execute "ALTER TABLE co_gastos ADD gastosgeneractacte bit default('0')"
End If
If Not ExisteElem(1, VGCNx, "co_gastos", "tipodocumentocodigo") Then
        VGCNx.Execute "ALTER TABLE co_gastos ADD tipodocumentocodigo varchar(2) default('00')"
End If
If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresadescrcorta") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresadescrcorta varchar(15) "
End If
If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresatelefonos") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresatelefonos varchar(20) "
End If
If Not ExisteElem(1, VGConfig, "empresa", "multiguiasremision") Then
        VGConfig.Execute "ALTER TABLE empresa ADD multiguiasremision bit default('0')"
End If
If Not ExisteElem(1, VGConfig, "empresa", "multifacturas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD multifacturas bit default('0') "
End If
If Not ExisteElem(1, VGConfig, "empresa", "multiboletas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD multiboletas bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "maeart", "estadodetraccion") Then
        VGCNx.Execute "ALTER TABLE maeart ADD estadodetraccion bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "kitvirtual") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD kitvirtual bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "vt_pedido", "pedidoobserva") Then
        VGCNx.Execute "ALTER TABLE vt_pedido ADD pedidoobserva varchar(200) default('0') "
End If
If Not ExisteElem(1, VGCNx, "tabtransa", "ingresosfuturos") Then
        VGCNx.Execute "ALTER TABLE tabtransa ADD ingresosfuturos bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "minimodetraccion") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD minimodetraccion float default('999999') "
End If
If Not ExisteElem(1, VGCNx, "co_sistema", "codigopercepcion") Then
        VGCNx.Execute "ALTER TABLE co_sistema ADD codigopercepcion nvarchar(20) "
End If
    If Not ExisteElem(1, VGCNx, "cp_tipodocumento", "tdocumentointerempresa") Then
        VGCNx.Execute "ALTER TABLE cp_tipodocumento ADD tdocumentointerempresa bit default('0')"
   End If
    If Not ExisteElem(1, VGCNx, "te_cuentabancos", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE te_cuentabancos ADD empresacodigo char(2) default('01')"
   End If
   If Not ExisteElem(1, VGCNx, "co_modoprovi", "modoprovianalitico") Then
        VGCNx.Execute "ALTER TABLE co_modoprovi ADD modoprovianalitico bit default('0')"
   End If
   If Not ExisteElem(1, VGCNx, "co_cabeceraprovisiones", "cabprovianalitico") Then
        VGCNx.Execute "ALTER TABLE co_cabeceraprovisiones ADD cabprovianalitico varchar(11)"
   End If
   If Not ExisteElem(1, VGCNx, "co_sistema", "TipoDocAcuenta") Then
        VGCNx.Execute "ALTER TABLE co_sistema ADD TipoDocAcuenta char(2)"
   End If
   If Not ExisteElem(1, VGCNx, "co_sistema", "TipoDocRetencion") Then
        VGCNx.Execute "ALTER TABLE co_sistema ADD TipoDocRetencion char(2)"
   End If
   If Not ExisteElem(1, VGCNx, "co_modoprovi", "librocodigo") Then
        VGCNx.Execute "ALTER TABLE co_modoprovi ADD librocodigo char(2) default('00')"
   End If
   If Not ExisteElem(1, VGCNx, "co_modoprovi", "mesproceso") Then
        VGCNx.Execute "ALTER TABLE co_modoprovi ADD mesproceso char(6) default('000000')"
   End If
   If Not ExisteElem(1, VGCNx, "te_cabecerarecibos", "cabprovinumaux") Then
     VGCNx.Execute "ALTER TABLE te_cabecerarecibos ADD cabprovinumaux varchar(10) default('')"
   End If
   If Not ExisteElem(1, VGCNx, "co_cabeceraprovisiones", "cabprovinumlibro") Then
     VGCNx.Execute "ALTER TABLE co_cabeceraprovisiones ADD cabprovinumlibro varchar(20) default('')"
   End If
   If Not ExisteElem(0, VGConfig, "si_usuario") Then
     VGConfig.Execute "select * into si_usuario  from usuario"
   End If
   If Not ExisteElem(1, VGCNx, "vt_parametroventa", "PedidosSinfacturar") Then
     VGCNx.Execute "ALTER TABLE vt_parametroventa ADD PedidosSinfacturar bit default(0)"
     VGCNx.Execute "update vt_parametroventa SET PedidosSinfacturar=0"
   End If
   If Not ExisteElem(1, VGCNx, "co_cabordcompra", "puntovtacodigo") Then
     VGCNx.Execute "ALTER TABLE co_cabordcompra ADD puntovtacodigo char(2) default('00')"
     VGCNx.Execute "update co_cabordcompra SET puntovtacodigo='00'"
   End If
   If Not ExisteElem(1, VGCNx, "co_cabordcompra", "trasladofisico") Then
     VGCNx.Execute "ALTER TABLE co_cabordcompra ADD trasladofisico bit default(0)"
     VGCNx.Execute "update co_cabordcompra SET trasladofisico=0"
   End If
   If Not ExisteElem(1, VGCNx, "co_estadorequerimiento", "NivelRequerimiento") Then
     VGCNx.Execute "ALTER TABLE co_estadorequerimiento ADD NivelRequerimiento char(1) default('0')"
     VGCNx.Execute "update co_estadorequerimiento SET NivelRequerimiento='0'"
   End If
   If Not ExisteElem(1, VGCNx, "co_tipodeorden", "flagrequerimientosPedidos") Then
     VGCNx.Execute "ALTER TABLE co_tipodeorden ADD flagrequerimientosPedidos char(1) default('0')"
     VGCNx.Execute "update co_tipodeorden SET flagrequerimientosPedidos='0'"
   End If
   If Not ExisteElem(1, VGCNx, "co_cabordcompra", "estadoordencodigo") Then
     VGCNx.Execute "ALTER TABLE co_cabordcompra ADD estadoordencodigo integer default(0)"
     VGCNx.Execute "update co_cabordcompra SET estadoordencodigo=0"
   End If
   If ExisteElem(1, VGCNx, "VT_detallepedido", "unidadcodigo") Then
     VGCNx.Execute "ALTER TABLE VT_detallepedido ALTER COLUMN unidadcodigo varchar(5) "
   End If
   If Not ExisteElem(1, VGCNx, "al_sistema", "SaldoConsolidadoxPedidos") Then
     VGCNx.Execute "ALTER TABLE al_sistema ADD SaldoConsolidadoxPedidos integer default(0) "
   End If
   If Not ExisteElem(1, VGCNx, "al_sistema", "SaldoConsolidadoxPedidos") Then
     VGCNx.Execute "ALTER TABLE al_sistema ADD SaldoConsolidadoxPedidos integer default(0) "
   End If

   Exit Sub
err2:
 MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
Resume Next
End Sub
Public Property Get ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    Dim NombrePC As String
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Property
    End If
    ipos = InStr(sName, Chr$(0))
    Randomize
    NombrePC = Trim$(Str(CLng(Rnd * 10000000)))

    ComputerName = "##" + Left$(sName, ipos - 1) + NombrePC
End Property
Public Sub central(f As Form)
    f.Left = (Screen.Width - f.Width) / 2
    f.Top = (Screen.Height / 1.19 - f.Height)
End Sub

Public Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase$(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim$(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim$(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim$(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim$(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGConfig, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function
Public Function Validar_RUC(xRuc As String) As Boolean
 Dim flag As Boolean
 Dim TAB_VAL(1 To 7) As Integer
 Dim nX As Integer, NY As Integer, NR As Integer, i As Integer
 Dim CadNR As String
 
' TAB_VAL(1) = 2
' TAB_VAL(2) = 7
' TAB_VAL(3) = 6
' TAB_VAL(4) = 5
' TAB_VAL(5) = 4
' TAB_VAL(6) = 3
' TAB_VAL(7) = 2
 flag = True
 xRuc = Trim$(xRuc)
 
' If xRuc <> " " Then
  'If xRuc <> "00000002" Then
     If Len(RTrim$(xRuc)) < 11 Then
         MsgBox "N�mero de R.U.C. no tiene 11 d�gitos", vbExclamation, "Ingreso de Datos"
         flag = False
      Else
'         nX = 0
'         NR = 0
'         NY = 0
'         CadNR = ""
'         For i = 1 To 7
'             nX = nX + Val(mid$(xRuc, i, 1)) * TAB_VAL(i)
'         Next i
'         NY = nX \ 11
'         NR = 11 - (nX - (NY * 11))
'         CadNR = TRIM$(String(10 - Len(Str(NR)) + 1, "0")) & TRIM$(Str(NR))
'         If mid$(CadNR, 10, 1) = mid$(xRuc, 8, 1) Then
'            flag = True
''         Else
'            MsgBox "N�mero de R.U.C. invalido", vbExclamation, "Ingreso de Datos"
'            flag = False
'         End If
      End If
'   Else
'      MsgBox "Anexo emite Liquidaciones de compra", vbExclamation, "Ingreso de Datos"
 '  End If
 'End If
 Validar_RUC = flag
End Function
'*************************************************
'Elimina de ( ' ) de una Cadena
'para Grabarla en una instrucci�n SQL
'*************************************************
Public Function SupCadSQL(s As String) As String
 Dim Aux As String
 If Not IsNull(s) Then
     Aux = Replace(s, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Public Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport
        If Right$(VGParamSistem.RutaReport, 1) <> "\" Then
           .ReportFileName = VGParamSistem.RutaReport & "\"
        End If
        .ReportFileName = .ReportFileName & VGParamSistem.carpetareportes
        
        If Right$(.ReportFileName, 1) <> "\" Then
        .ReportFileName = .ReportFileName & "\"
        End If
        '.ReportFileName &
        .ReportFileName = .ReportFileName & cNombreReporte
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
        Else
           .Connect = VGCadenaReport2
        End If
           
        .formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"     'aki va el ruc
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, i As Integer
Dim Valor As String
    i = 0
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        'I = 0
        If pos = 0 Then Exit Do
        Valor = Left$(cad, pos - 1)
        cry.SortFields(i) = Valor
        i = i + 1
        cad = Right$(cad, (Len(cad) - pos))
    Loop
End Sub

Sub ImpresionRptbase(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport & "\" & cNombreReporte
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
 
        End If
           
        .formulas(0) = "@Emp='" & VGParametros.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Public Sub PropCrystal(ByRef CrystalRpt As CrystalReport)
    CrystalRpt.WindowShowCancelBtn = True
    CrystalRpt.WindowShowCloseBtn = True
    CrystalRpt.WindowShowExportBtn = True
    CrystalRpt.WindowShowGroupTree = True
    CrystalRpt.WindowShowNavigationCtls = True
    CrystalRpt.WindowShowPrintBtn = True
    CrystalRpt.WindowShowPrintSetupBtn = True
    CrystalRpt.WindowShowProgressCtls = True
    CrystalRpt.WindowShowSearchBtn = True
    CrystalRpt.WindowShowZoomCtl = True
    CrystalRpt.Destination = crptToWindow
    CrystalRpt.WindowState = crptMaximized


End Sub

Sub ImpresionRpt_SubRpt_Proc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional ORDEN As String, Optional titulo As String)
Dim strBuscar As New dll_apis
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        If Right$(VGParamSistem.RutaReport, 1) <> "\" Then VGParamSistem.RutaReport = VGParamSistem.RutaReport & "\"
        .ReportFileName = VGParamSistem.RutaReport + cNombreReporte
        
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2

        End If
           
        .formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
         .DiscardSavedData = True
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2

        End If

        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Public Function XRecuperaTipoCambio(Fecha As Date, tipo As tipocambio, cnx As ADODB.Connection) As Double
Dim RSAUX As ADODB.Recordset
Set RSAUX = New ADODB.Recordset
Dim Campo As String
    XRecuperaTipoCambio = 0
    Select Case tipo
        Case Compra
            Campo = "tipocambiocompra"
        Case Venta
            Campo = "tipocambioventa"
        Case Promedio
            Campo = "tipocambiopromedio"
        Case Else
            Campo = "tipocambioventa"
    End Select
    SQL = "Select Valor=isnull(" & Campo & ",0)  from ct_tipocambio where convert(varchar(10),tipocambiofecha,103) ='" & Fecha & "'"
    Set RSAUX = VGCNx.Execute(SQL)
    If RSAUX.RecordCount > 0 Then
        XRecuperaTipoCambio = RSAUX!Valor
    End If
End Function
Public Function ExisteSQL(ByVal cnx As ADODB.Connection, ByVal SentenciaSQL As String) As Boolean
On Error GoTo SaliError
    Screen.MousePointer = 11
    ExisteSQL = False
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open SentenciaSQL, cnx, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        ExisteSQL = True
    End If
    Screen.MousePointer = 1
    Exit Function
SaliError:
    Screen.MousePointer = 1
    ExisteSQL = False
    MsgBox Err.Description
    Exit Function
    Resume
End Function

Public Sub ADOConectar()
On Error GoTo error
Set VGGeneral = New ADODB.Connection
VGGeneral.CursorLocation = adUseClient
VGGeneral.CommandTimeout = 0
VGGeneral.ConnectionTimeout = 200
VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.ServidorGEN
VGGeneral.Open

   
'Conexion de Cofiguracion

Set VGConfig = New ADODB.Connection
VGConfig.CursorLocation = adUseClient
VGConfig.CommandTimeout = 0
VGConfig.ConnectionTimeout = 0
VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
VGConfig.Open
    
'Conexion de inventarios

If VGParamSistem.BDEmpresa = "" Or VGParamSistem.BDEmpresa = "?" Then
   Set RSQL = VGConfig.Execute("select empresabaseinventarios from empresa where empresaflaginventarios=1")
   VGParamSistem.BDEmpresa = RSQL!empresabaseinventarios
End If
Set VGCNx = New ADODB.Connection
VGCNx.CursorLocation = adUseClient
VGCNx.CommandTimeout = 0
VGCNx.ConnectionTimeout = 0
VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
VGCNx.Open
    
'Conexion de Contabilidad

Set VGcnxCT = New ADODB.Connection
VGcnxCT.CursorLocation = adUseClient
VGcnxCT.CommandTimeout = 0
VGcnxCT.ConnectionTimeout = 0
VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
VGcnxCT.Open
    
'Call adicionacamposct
Exit Sub

error:
    
MsgBox Err.Description, vbExclamation
Exit Sub
Resume
End Sub

Public Function Fecha(ByVal tipo As Integer, dato As Date) As Date
Dim fecha1 As Date
fecha1 = Format("01/" & Format(Month(dato), "00") & "/" & Year(dato), "dd/mm/yyyy")
Select Case tipo
        Case 1
          Fecha = fecha1
        Case 2
          fecha1 = fecha1 + 31
          fecha1 = Format("01/" & Format(Month(fecha1), "00") & "/" & Year(fecha1), "dd/mm/yyyy")
          Fecha = fecha1 - 1
        Case 3
          fecha1 = fecha1 - 31
          Fecha = Format("01/" & Format(Month(fecha1), "00") & "/" & Year(fecha1), "dd/mm/yyyy")
End Select
End Function

Public Function ESNULO(EXPRESION As Variant, Valor As Variant) As Variant
On Error GoTo errfun
   If IsNull(EXPRESION) Or Trim$(EXPRESION) = Empty Then
      ESNULO = Valor
     Else: ESNULO = EXPRESION
   End If
   Exit Function
errfun:
   ESNULO = 0
End Function
Public Function ExisteElem(ByRef Tip As Integer, ByRef VGCN As ADODB.Connection, ByRef Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim RSAUX As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   Tabla = UCase$(Tabla): Campo = UCase$(Campo)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & Tabla
        Case 1:
            SQL = "Select Top 1 " & Campo & " From " & Tabla
    End Select
    Set RSAUX = VGCN.Execute(SQL)
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function
Public Function DateSQL(ByVal Fecha As String) As String
    'On Error GoTo ERR
    If IsNull(Fecha) Then Exit Function
        Select Case VGformatofecha
            Case "DMY"
            DateSQL = "'" & Format(Fecha, "dd/mm/yyyy") & "'"
            Case "MDY"
            DateSQL = "'" & Format(Fecha, "mm/dd/yyyy") & "'"
        End Select
'ERR:
 '    DateSQL = "'" & Day(FECHA) & "/" & Month(FECHA) & "/" & Year(FECHA) & "'"
End Function

Function DesMes(nMes As String) As String
Dim DescriMes As String

Select Case nMes
   Case "01"
               DescriMes = "ENERO "
   Case "02"
               DescriMes = "FEBRERO  "
   Case "03"
               DescriMes = "MARZO "
   Case "04"
               DescriMes = "ABRIL "
    Case "05"
               DescriMes = "MAYO "
    Case "06"
               DescriMes = "JUNIO "
    Case "07"
               DescriMes = "JULIO "
    Case "08"
               DescriMes = "AGOSTO "
    Case "09"
               DescriMes = "SETIEMBRE "
    Case "10"
               DescriMes = "OCTUBRE "
    Case "11"
               DescriMes = "NOVIEMBRE "
    Case "12"
               DescriMes = "DICIEMBRE "
End Select

DesMes = DescriMes
End Function

'Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
' With EsteGrid
'  .AllowAddNew = False
'  .AllowDelete = False
'  .AllowUpdate = False
'  .AllowRowSizing = False
'  .TabAction = dbgControlNavigation
'  .MarqueeStyle = dbgHighlightRow
 ' .Font =
' End With
'End Sub

Public Function Devolver_Dato(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, CampDev As String, Optional Cod2 As String, Optional Campo2 As String, Optional Cod3 As String, Optional Campo3 As String, Optional Cod4 As Double, Optional Campo4 As String) As String
Dim cSel1 As ADODB.Recordset, cF As String
Set cSel1 = New ADODB.Recordset

If Trim$(Campo) <> "" Then
    If Fecha = False Then
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  '" & Cod & "' "
    Else
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  #" & Format(Cod, "mm/dd/yyyy") & "#"
    End If
End If
If Trim$(Campo2) <> "" Then
    cF = cF & " and " & Campo2 & " = '" & Cod2 & "' "
End If
If Trim$(Campo3) <> "" Then
    cF = cF & " and " & Campo3 & " = '" & Cod3 & "' "
End If
If Trim$(Campo4) <> "" Then
    cF = cF & " and " & Campo4 & " = '" & Cod4 & "' "
End If
Select Case tipo
  Case 1 'Bd. Comun
              cSel1.Open cF, VGCNx, adOpenStatic
  Case 2 'Bd. Config
              cSel1.Open cF, VGConfig, adOpenStatic
  Case 3 'Bd. Contabilidad
              cSel1.Open cF, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Devolver_Dato = IIf(Not IsNull(cSel1(0)), cSel1(0), "")
Else
     Devolver_Dato = ""
End If
End Function

Public Function NUMLET(num As String) As String
Dim cLET As String
Dim cWork As String
Dim cUNIDAD As String
Dim cDECENA As String
Dim cCENTENA As String
Dim nMODULUS As Integer
Dim nI As Integer
Dim nK As Integer
Dim Lit1 As String
Dim Lit2 As String
Dim Lit3 As String
Dim Lit4 As String
Dim Lit5 As String
Lit1 = "Uno    Dosc   Trec   Cuatroc  Quin   Seisc  Setec  Ochoc  Novec  "
Lit2 = "Diez     Veinte   Treinta  Cuarenta CincuentaSesenta  Setenta  Ochenta  Noventa  "
Lit3 = "Once      Doce      Trece     Catorce   Quince    Dieciseis DiecisieteDieciocho Diecinueve"
Lit4 = "Uno   Dos   Tres  CuatroCinco Seis  Siete Ocho  Nueve "
Lit5 = "Millon    Billon    Trillon   CuatrillonQuintillon"
'Proceso Input = Num , Output = Let

cLET = ""

'Dim NUM As Double
'NUM = Val(NUMx)

If num > 0.99 Then
    'Separa los Enteros en una Cadena de Caracteres
     If InStr(1, Trim$(Str(num)), ".", 0) > 0 Then
        cWork = Mid$(Trim$(Str(num)), 1, InStr(1, Trim$(Str(num)), ".", 0) - 1)
     Else
        cWork = Str(num)
     End If
     nMODULUS = Int(Len(Trim$(cWork)) / 3)
     nMODULUS = Len(Trim$(cWork)) - (nMODULUS * 3)
     
     If nMODULUS > 0 Then
        cWork = String(3 - nMODULUS, "0") & Trim$(cWork)
     End If
     
     nK = (Len(Trim$(cWork)) / 3) - 1
    'Procesa de Mil en Mil
     nI = 1
     Do While nI < Len(Trim$(cWork)) - 1
        cCENTENA = Mid$(Trim$(cWork), nI, 1)
        cDECENA = Mid$(Trim$(cWork), nI + 1, 1)
        cUNIDAD = Mid$(Trim$(cWork), nI + 2, 1)
        'Centenas
        If cCENTENA <> "0" Then
            If cCENTENA = "1" Then
                cLET = cLET & "Cien "
                If cDECENA <> "0" Or cUNIDAD <> "0" Then
                    cLET = Mid$(cLET, 1, (Len(cLET) - 1)) & "to "
                End If
            Else
                cLET = cLET & Trim$(Mid$(Lit1, ((Val(cCENTENA) - 1) * 7) + 1, 7)) & "ientos "
            End If
        End If
        'Decenas
        If cDECENA <> "0" Then
            If cDECENA = "1" And cUNIDAD <> "0" Then
                If ((Val(cUNIDAD) - 1) * 10) + 1 > 0 Then cLET = cLET + Trim$(Mid$(Lit3, ((Val(cUNIDAD) - 1) * 10) + 1, 10))
            Else
                If ((Val(cDECENA) - 1) * 9) + 1 > 0 Then cLET = cLET + Trim$(Mid$(Lit2, ((Val(cDECENA) - 1) * 9) + 1, 9))
            End If
        End If
        'Unidades
        If cUNIDAD <> "0" Then
            If cDECENA > "1" Then
                cLET = Mid$(cLET, 1, (Len(cLET) - 1)) & "i"
                If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + LCase(Trim$(Mid$(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6)))
            Else
                If cDECENA < "1" Then
                    If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + Trim$(Mid$(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6))
                End If
            End If
        End If
        cLET = cLET & " "
        'Pone Miles o Millones
        If nK > 0 Then
            If cCENTENA & cDECENA & cUNIDAD = "001" Then
                cLET = Mid$(cLET, 1, Len(cLET) - 2) & " "
            End If
            nMODULUS = Int(nK / 2)
            nMODULUS = nK - (nMODULUS * 2)
            If nMODULUS = 0 Then
                cLET = cLET + Trim$(Mid$(Lit5, (((nK / 2) - 1) * 10) + 1, 10))
                If cCENTENA & cDECENA & cUNIDAD = "001" Or num > 1999999 Then
                    cLET = cLET & "es "
                Else
                    cLET = cLET & " "
                End If
            Else
                If cCENTENA & cDECENA & cUNIDAD > "000" Then
                    cLET = cLET & "Mil "
                End If
            End If
            nK = nK - 1
        End If
        nI = nI + 3
    Loop
    cLET = cLET & "con "
End If
If InStr(1, Trim$(Str(num)), ".", 0) > 0 Then
    cLET = cLET + Mid$(Trim$(Str(num)), InStr(1, Trim$(Str(num)), ".", 0) + 1, 2) & "/100" & " "
Else
    cLET = cLET + "00/100" & " "
End If
NUMLET = cLET
End Function

Public Function CODIFICA(CADENA As String, Valor As Integer) As String
    Dim ciclo As Integer, posic As Integer
    Dim utl_sal As Integer
    Dim carac As String, cadena_cod As String, cad As String
    posic = 0: utl_sal = 0
    carac = "": cadena_cod = "": cad = ""
    CADENA = UCase$(Trim$(CADENA))
    For ciclo = 1 To Len(CADENA)
     carac = Mid$(CADENA, ciclo, 1)
     If (ciclo Mod 2) = 0 Then
      carac = UCase$(carac)
     Else
      carac = LCase(carac)
     End If
     cadena_cod = cadena_cod & carac
    Next ciclo
    
    For ciclo = 1 To Len(cadena_cod)
     posic = ciclo Mod 7
     carac = Mid$(cadena_cod, ciclo, 1)
     Select Case posic
     Case 0:
            carac = Chr(Asc(carac) * 2)
     Case 1:
            carac = Chr(Asc(carac) - Valor)
     Case 2:
            carac = Chr(Asc(carac) - (ciclo * 2))
            utl_sal = Asc(carac)
     Case 3:
            If utl_sal > 10 Then utl_sal = utl_sal - (Int(utl_sal / 10) * 10)
            carac = Chr(Asc(carac) - Valor + utl_sal)
     Case 4:
            carac = Chr(Asc(carac) - ciclo)
            utl_sal = Asc(carac)
     Case 5:
            If utl_sal > 10 Then utl_sal = utl_sal - (Int(utl_sal / 10) * 10)
            carac = Chr(Asc(carac) - Valor + utl_sal)
     End Select
     cad = cad + carac
    Next ciclo
    CODIFICA = cad
End Function
'funci�n que desencripta una cadena
Public Function DECODIFICA(CADENA As String, Valor As Integer) As String
    Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
    Dim carac As String, cad As String
    CADENA = Trim$(CADENA)
    cad = ""
    val_n = 0: val_an = 0
    For ciclo = 1 To Len(CADENA)
     carac = Mid$(CADENA, ciclo, 1)
     posic = ciclo Mod 7
     Select Case posic
     Case 0:
            val_n = Asc(carac) / 2
     Case 1:
            val_n = Asc(carac) + Valor
     Case 2:
            val_n = Asc(carac) + (ciclo * 2)
            val_an = Asc(carac)
     Case 3:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + Valor - val_an
     Case 4:
            val_n = Asc(carac) + ciclo
     Case 5:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + Valor - val_an
     Case 6:
           val_n = Asc(carac)
     End Select
     cad = cad + Chr(val_n)
    Next ciclo
    DECODIFICA = UCase$(cad)
End Function
Public Function numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     If IsNumeric(Number) Then
        aValor = Number
     Else
      aValor = Val(Number)
     End If
   End If
   numero = Trim(Format(aValor, "#######0.00"))
End Function

Public Function MostrarForm(pVentana As Form, pPos As String)
    
   'pVentana.Icon = LoadPicture(App.Path & "\Cuenta.ico")
   
   If pPos = "C" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   ElseIf pPos = "M" And pVentana.Visible = False Then
      pVentana.Caption = pVentana.Caption & "  " & VGParametros.NomEmpresa
      pVentana.Width = Screen.Width
   ElseIf pPos = "C1" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   ElseIf pPos = "C2" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   End If
   pVentana.Panel.Panels(1).Width = (pVentana.Width / 4)
   If pPos = "M" Then
      pVentana.Panel.Panels(1).Width = ((pVentana.Width - 2600) / 4)
      pVentana.Panel.Panels(1).text = "EMPRESA: " & VGParametros.NomEmpresa
      pVentana.Panel.Panels(2).text = "PTO. VENTA: " & VGParametros.puntovta
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
   Else
      pVentana.Panel.Panels(1).text = "FORMATO : " & Escadena(pVentana.Caption)
      pVentana.Panel.Panels(2).text = "USUARIO: " & VGUsuario
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
   End If
   pVentana.Panel.Panels(1).Alignment = sbrLeft
   pVentana.Panel.Panels(3).text = "FECHA :" & Format(Date, "dd/mm/yyyy")
   pVentana.Panel.Panels(3).Alignment = sbrRight
   pVentana.Panel.Panels(3).Width = (pVentana.Width / 4)
   pVentana.Panel.Panels(4).text = "HORA :" & Format(Time, "hh:mm:ss")
   pVentana.Panel.Panels(4).Alignment = sbrRight
   pVentana.Panel.Panels(4).Width = (pVentana.Width / 4)

End Function

Public Function MostrarFormVentas(pVentana As Form, pPos As String)
    
   'pVentana.Icon = LoadPicture(App.Path & "\Cuenta.ico")
   
   If pPos = "C" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   ElseIf pPos = "M" And pVentana.Visible = False Then
      pVentana.Caption = pVentana.Caption & "  " & VGParametros.NomEmpresa
      pVentana.Width = Screen.Width
   ElseIf pPos = "C1" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   ElseIf pPos = "C2" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   End If


End Function

Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
End Function




Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer, Optional Noincluir1, Optional Noincluir2 As Integer)
 Dim J As Integer
 If IsMissing(Noincluir1) Then
    Noincluir1 = -1
 End If
 If IsMissing(Noincluir2) Then
    Noincluir2 = -1
 End If
   
 For J = ninicio To nfin
   If J = Noincluir1 Or J = Noincluir2 Then
   Else
      MBox(J) = Empty
   End If
 Next J
End Function
Public Function TraeDataSerie(nsql As String, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute(nsql)
    If rsbuscn.RecordCount > 0 Then
        TraeDataSerie = rsbuscn!puntovtadoccorr
    Else
        TraeDataSerie = "1"
    End If
    Set rsbuscn = Nothing

End Function

Public Function VerificaCombo(xcombo As ComboBox, ncadena As String) As Long
    Dim J, k As Long
    On Error GoTo nerror
    VerificaCombo = -1
    If xcombo.ListCount > 0 Then
      VerificaCombo = 0
      For J = 0 To xcombo.ListCount - 1
         xcombo.ListIndex = J
         k = InStr(xcombo, "-")
         If k > 1 Then
           If Left(xcombo.text, k - 1) = ncadena Then
             VerificaCombo = J
             Exit For
           End If
         Else
           If xcombo.text = ncadena Then
             VerificaCombo = J
             Exit For
           End If
         End If
      Next J

    End If
    
nerror:
  If Err Then
    MsgBox Err.Number & "-" & Err.Description
    Err = 0
    Resume Next
  End If
End Function

Public Sub CargarTipo(xcombo As ComboBox, xtipo)
  Dim adll2 As New dllgeneral.dll_general
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento", VGCNx)
'     xcombo.AddItem g_tipobol & "-Boleta"
'     xcombo.AddItem g_tipofac & "-Factura"
'     xcombo.AddItem g_tipoguia & "-B.O."
     xcombo.ListIndex = 0
   Case 3   '---estado
     xcombo.Clear
     xcombo.AddItem "S-SI"
     xcombo.AddItem "N-NO"
     xcombo.ListIndex = 0
   Case 4  '-- Tipo persona
     xcombo.Clear
     xcombo.AddItem "1-NATURAL"
     xcombo.AddItem "2-JURIDICA"
     xcombo.ListIndex = 0
   Case 5  '-tipo pais
     xcombo.Clear
     xcombo.AddItem "1-PERUANA"
     xcombo.AddItem "2-EXTRANJERA"
     xcombo.ListIndex = 0
   Case 6   '--todos los tipos documentos --
     xcombo.Clear
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento ", VGCNx)
     'xcombo.AddItem g_tipobol & "-Boleta"
     'xcombo.AddItem g_tipofac & "-Factura"
     'xcombo.AddItem g_tipoguia & "-B.O."
     'xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.ListIndex = 0
     
  End Select
End Sub
Public Function Escadena(pdato) As String
   If IsNull(pdato) Then
      Escadena = ""
    ElseIf Len(Trim(pdato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim$(pdato)
   End If
End Function


Public Function DatoTipoCambio(xCn As ADODB.Connection, xfecha As String) As Double
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  SQL = "select tipocambiocompra,tipocambioventa from ct_tipocambio "
  SQL = SQL & "Where tipocambiofecha='" & Format(xfecha, "dd/mm/yyyy") & "'"
  Set rs = xCn.Execute(SQL)
  If Not (rs.EOF Or rs.BOF) Then
     DatoTipoCambio = Format(rs(1), "#####0.###0")
  Else
     DatoTipoCambio = Format(1, "#####0.###0")
  End If
  Set rs = Nothing
End Function


Public Sub imprimir(cNombreReporte As String)
Dim VGdllApi As New dll_apisgen.dll_apis
On Error GoTo Errores

With MDIPrincipal.CryRptProc
   Call PropCrystal(MDIPrincipal.CryRptProc)
   .ReportFileName = VGParamSistem.RutaReport
   If Right$(VGParamSistem.RutaReport, 1) <> "\" Then
     .ReportFileName = VGParamSistem.RutaReport & "\"
   End If
  .ReportFileName = .ReportFileName & VGParamSistem.carpetareportes
  If Right$(.ReportFileName, 1) <> "\" Then
        .ReportFileName = .ReportFileName & "\"
  End If
  .ReportFileName = .ReportFileName & cNombreReporte
  .Connect = "Provider=SQLOLEDB;PWD=" & VGParamSistem.Pwd & ";UID=" & VGParamSistem.Usuario & ";DSQ=" & VGParamSistem.BDEmpresa & ";DSN=" & VGParamSistem.Servidor
  .formulas(0) = "Empresa='" & VGParametros.NomEmpresa & "'"
  .Action = 1
End With
Exit Sub
    
Errores:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
  
End Sub
Public Sub GeneraAsientoEnlineaTesorTransfer(empresa As String, Fecha As Date, Nrecibo As String)
Dim rsparimpo As ADODB.Recordset
Dim Comando As ADODB.Command
On Error GoTo Procesotransf
    Set rsparimpo = New ADODB.Recordset
    rsparimpo.Open "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='T'", VGcnxCT, adOpenKeyset, adLockReadOnly
    Set Comando = New ADODB.Command
        With Comando
            .CommandType = adCmdStoredProc
            .CommandText = "te_GeneraAsientosTesoreriaTransflinea_pro"
            .ActiveConnection = VGGeneral
            .Parameters.Refresh
            .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
            .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
            .Parameters("@empresa") = empresa
            .Parameters("@Asiento") = rsparimpo!asiento
            .Parameters("@SubAsiento") = rsparimpo!SubAsiento
            .Parameters("@Libro") = rsparimpo!libro
            
            .Parameters("@Mes") = Format(Month(Fecha), "00")
            .Parameters("@Ano") = Year(Fecha)
            .Parameters("@Compu") = VGcomputer
            .Parameters("@Usuario") = VGParamSistem.Usuario
            .Parameters("@Ntransfer") = Nrecibo
            .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
            .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
            .Execute
        End With
        Screen.MousePointer = 1
        MsgBox "La Contabilizacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
        Exit Sub
Procesotransf:
        Screen.MousePointer = 1
        MsgBox Err.Description
        Exit Sub
        Resume
End Sub
Public Sub GeneraAsientoEnlineaTesor(Fecha As Date, empresa As String, m_Opcion As String, Nrecibo As String, op As Integer, comprobconta As String, monedacodigo As String, cajabanco As String, m_tipovoucher As String)
Dim rsparimpo As ADODB.Recordset
Dim numerror As Integer
Dim Comando As ADODB.Command
numerror = 0
On Error GoTo Proceso

   VGCNx.BeginTrans

Set rsparimpo = New ADODB.Recordset

rsparimpo.Open "Select * From  ct_importartesoreria Where tipooperacion ='" & UCase(m_Opcion) & "' ", VGcnxCT, adOpenKeyset, adLockReadOnly
If rsparimpo.RecordCount() > 0 Then

   Set Comando = New ADODB.Command
   With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
        .CommandTimeout = 0
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
        .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
        .Parameters("@empresa") = empresa
        .Parameters("@Asiento") = rsparimpo!asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!libro
         
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@Ano") = Year(Fecha)
            
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGcomputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@TipoMov") = Trim(UCase(m_tipovoucher))
        .Parameters("@Nrecibo") = Nrecibo
        .Parameters("@op") = op
        .Parameters("@comprobconta") = comprobconta
        .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
        .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
        .Execute
   End With
   If numerror = 0 Then
        VGCNx.CommitTrans
        Screen.MousePointer = 1
        MsgBox "La Contabilizacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
   End If
End If
Exit Sub
Proceso:
   numerror = 1
   Screen.MousePointer = 1
    MsgBox Err.Description
    VGCNx.RollbackTrans
   Exit Sub
   Resume
End Sub

Public Function DatoMoneda(xValor As String) As String
   Dim rmone As New ADODB.Recordset
   
   Set rmone = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & xValor & "'")
   If rmone.RecordCount > 0 Then
       DatoMoneda = Escadena(rmone!monedasimbolo) & " ."
   Else
       DatoMoneda = " "
   End If
   rmone.Close
   Set rmone = Nothing

End Function

