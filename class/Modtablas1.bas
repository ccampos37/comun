Attribute VB_Name = "ModificarCampos"
Option Explicit
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Enum TIPOSISTEMA
   inventarios = 1
   compras = 2
   pagar = 3
   caja = 4
   contab = 5
   facturacion = 6
End Enum

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

If Not ExisteElem(1, VGcnx, "maeart", "longitudderecha") Then
        VGcnx.Execute "ALTER TABLE maeart ADD longitudderecha float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "longitudizquierda") Then
        VGcnx.Execute "ALTER TABLE maeart ADD longitudizquierda float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "longitudfrontal") Then
        VGcnx.Execute "ALTER TABLE maeart ADD longitudfrontal float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "longitudposterior") Then
        VGcnx.Execute "ALTER TABLE maeart ADD longitudposterior float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "areaterreno") Then
        VGcnx.Execute "ALTER TABLE maeart ADD areaterreno float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "areaconstruida") Then
        VGcnx.Execute "ALTER TABLE maeart ADD areaconstruida float NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "numerodepisos") Then
        VGcnx.Execute "ALTER TABLE maeart ADD numerodepisos integer NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "numerodehabitaciones") Then
        VGcnx.Execute "ALTER TABLE maeart ADD numerodehabitaciones integer NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "numerodeservicios") Then
        VGcnx.Execute "ALTER TABLE maeart ADD numerodeservicios integer NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "linderofrontera") Then
        VGcnx.Execute "ALTER TABLE maeart ADD linderofrontera nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "linderoposterior") Then
        VGcnx.Execute "ALTER TABLE maeart ADD linderoposterior nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "linderoizquierdo") Then
        VGcnx.Execute "ALTER TABLE maeart ADD linderoizquierdo nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "linderoderecho") Then
        VGcnx.Execute "ALTER TABLE maeart ADD linderoderecho nvarchar(30) NULL"
End If
If Not ExisteElem(1, VGcnx, "maeart", "proyectocodigo") Then
        VGcnx.Execute "ALTER TABLE maeart ADD proyectocodigo nvarchar(3) NULL"
End If

End Sub
Public Sub adicionarcampos()
Dim SQL As String
Dim rsql As New Recordset
On Error GoTo ERROR1
VGtipo = 0
If VGtipo <> contab Then
   If Not ExisteElem(1, VGConfig, "empresa", "empresabasecontabilidad") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabasecontabilidad nvarchar(20) NULL"
   End If
   If Not ExisteElem(1, VGCnxCT, "ct_sistema", "sistemaconfiguracentrocostos") Then
        VGCnxCT.Execute "ALTER TABLE ct_sistema ADD sistemaconfiguracentrocostos VARCHAR(20) NULL"
   End If
   If Not ExisteElem(1, VGCnxCT, "ct_sistema", "sistemaultimonivelcostos") Then
        VGCnxCT.Execute "ALTER TABLE ct_sistema ADD sistemaultimonivelcostos INTEGER NULL"
   End If
   If Not ExisteElem(1, VGCnxCT, "ct_centrocosto", "centrocostonivel") Then
        VGCnxCT.Execute "Alter table ct_centrocosto ADD centrocostonivel integer NULL"
   End If
End If
If Not ExisteElem(1, VGcnx, "co_gastos", "gastosequivalente") Then
        VGcnx.Execute "ALTER TABLE co_gastos ADD gastosequivalente varchar(4) NULL"
End If
5 If Not ExisteElem(1, VGConfig, "empresa", "digitoscodigo") Then
        VGConfig.Execute "ALTER TABLE empresa ADD digitoscodigo integer NULL"
End If
If Not ExisteElem(1, VGcnx, "familia", "correlativocodigo") Then
        VGcnx.Execute "ALTER TABLE familia ADD correlativocodigo integer NULL"
End If
If Not ExisteElem(1, VGcnx, "co_gastos", "tipoanaliticocodigo") Then
   VGcnx.Execute "Alter table co_gastos ADD tipoanaliticocodigo VARCHAR(3) NULL"
End If
If Not ExisteElem(0, VGcnx, "co_multiempresas") Then
   SQL = " Create Table co_multiempresas "
   SQL = SQL & "( empresacodigo VarChar(2),"
   SQL = SQL & "empresadescripcion VarChar(20),"
   SQL = SQL & "usuariocodigo varchar(8),fechaact datetime "
   SQL = SQL & " CONSTRAINT PK_co_multiempresas "
   SQL = SQL & " PRIMARY KEY (empresacodigo))  "
   VGcnx.Execute SQL
End If
If Not ExisteElem(1, VGcnx, "co_multiempresas", "usuariocodigo") Then
        VGcnx.Execute "ALTER TABLE co_multiempresas ADD usuariocodigo VARCHAR(8) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_multiempresas", "fechaact") Then
        VGcnx.Execute "ALTER TABLE co_multiempresas ADD fechaact datetime NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "sistemamultiempresas") Then
       VGcnx.Execute "ALTER TABLE co_sistema ADD sistemamultiempresas VARCHAR(1) NULL"
    End If
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD empresacodigo VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "entidadcodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD entidadcodigo VARCHAR(11) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "detrec_gastos") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD detrec_gastos VARCHAR(10) NULL"
End If
If Not ExisteElem(1, VGcnx, "ct_centrocosto", "centrocostonivel") Then
   VGcnx.Execute "Alter table ct_centrocosto ADD centrocostonivel integer NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "centrocostocodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD centrocostocodigo VARCHAR(10) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "cabprovinumero") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD cabprovinumero INTEGER NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "sistemaminimoretencion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD sistemaminimoretencion INTEGER NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_proveedor", "proveedorcontribuyente") Then
        VGcnx.Execute "ALTER TABLE cp_proveedor ADD proveedorcontribuyente bit NULL"
End If
If ExisteElem(1, VGcnx, "cp_cargo", "cargoemiteretencion") Then
   VGcnx.Execute "ALTER TABLE cp_cargo ALTER COLUMN cargoemiteretencion varchar(1) NULL"
 Else
   VGcnx.Execute "ALTER TABLE cp_cargo ADD cargoemiteretencion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipocompra", "tipocomprainafecta") Then
        VGcnx.Execute "ALTER TABLE co_tipocompra ADD tipocomprainafecta varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_multiempresas", "agentederetencion") Then
        VGcnx.Execute "ALTER TABLE co_multiempresas ADD agentederetencion varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "detalle_no_saldos") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD detalle_no_saldos varchar(1) NULL"
End If
If ExisteElem(1, VGcnx, "cp_tipodocumento", "tdocumentonumerador") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ALTER COLUMN tdocumentonumerador VARCHAR(11) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_tipodocumento", "documentoretencion") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento add documentoretencion VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "cargoemitedetraccion") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD cargoemitedetraccion varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_gastos", "habilitadodetraccion") Then
        VGcnx.Execute "ALTER TABLE co_gastos ADD habilitadodetraccion varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD empresacodigo varchar(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "Bancarizacion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD Bancarizacion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion01") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion01 float NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion02") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion02 float NULL"
End If
'----

If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD empresacodigo VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "cabcomprobnumero") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD cabcomprobnumero INT NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "entidadcodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD entidadcodigo VARCHAR(11) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "centrocostocodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD centrocostocodigo VARCHAR(10) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_conceptocaja", "conceptosiccosto") Then
        VGcnx.Execute "ALTER TABLE te_conceptocaja ADD conceptosiccosto VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresalistaestadoclientes") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresalistaestadoclientes VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresalistaestadoproveedor") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresalistaestadoproveedor VARCHAR(1) NULL"
End If
If ExisteElem(1, VGcnx, "cp_tipodocumento", "tdocumentonumerador") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ALTER COLUMN tdocumentonumerador VARCHAR(11) NULL"
End If

If Not ExisteElem(0, VGcnx, "te_rendiciones") Then
   SQL = " Create Table te_rendiciones (oficinacodigo VarChar(3),"
   SQL = SQL & "codigocaja varchar(2),monedacodigo VarChar(2),"
   SQL = SQL & "rendicionnumero VarChar(6),"
   SQL = SQL & "rendicionsaldoinicial float,"
   SQL = SQL & "rendicioningresos float,"
   SQL = SQL & "rendicionegresos float,"
   SQL = SQL & "rendicionsaldofinal float,"
   SQL = SQL & "rendicionfecha datetime,"
   SQL = SQL & "usuariocodigo varchar(8),fechaact datetime "
   SQL = SQL & " CONSTRAINT PK_te_rendiciones "
   SQL = SQL & " PRIMARY KEY (oficinacodigo,monedacodigo,rendicionnumero))  "
   VGcnx.Execute SQL
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "rendicionnumero") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD rendicionnumero VARCHAR(6) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "cajarendiciones") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD cajarendiciones bit NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "rendicionnumero01") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD rendicionnumero01 VARCHAR(6) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "rendicionnumero02") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD rendicionnumero02 VARCHAR(6) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "cargoemiteretencion") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD cargoemiteretencion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_tipodocumento", "documentoretencion") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ADD documentoretencion varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "detalle_no_saldos") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD detalle_no_saldos varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresaretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresaretencion VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresacodigoretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresacodigoretencion VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "porcentajeretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD porcentajeretencion  float  NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "sistemaminimoretencion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD sistemaminimoretencion INTEGER NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresacodigodetraccion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresacodigodetraccion VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD empresacodigo varchar(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "Bancarizacion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD Bancarizacion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion01") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion01 float NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion02") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion02 float NULL"
End If

'---

If Not ExisteElem(1, VGConfig, "empresa", "empresaflagcompras") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagcompras nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabasecompras") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabasecompras nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflaginventarios") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflaginventarios nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabaseinventarios") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabaseinventarios nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagpagar") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagpagar nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabasepagar") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabasepagar nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagtesoreria") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagtesoreria nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabasetesoreria") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabasetesoreria nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagventas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagventas nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabaseventas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabaseventas nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagcobrar") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagcobrar nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabasecobrar") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabasecobrar nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipodeorden", "flagrequerimientos") Then
        VGcnx.Execute "ALTER TABLE co_tipodeorden ADD flagrequerimientos nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipodeorden", "ordendebienes") Then
        VGcnx.Execute "ALTER TABLE co_tipodeorden ADD ordendebienes nvarchar(1) NULL"
End If
If ExisteElem(1, VGcnx, "co_cabordcompra", "OC_CENTREG") Then
        VGcnx.Execute "ALTER TABLE co_cabordcompra alter column OC_CENTREG nvarchar(80) NULL"
End If
If ExisteElem(1, VGcnx, "co_cabordcompra", "OC_COBSERV") Then
        VGcnx.Execute "ALTER TABLE co_cabordcompra alter column OC_COBSERV nvarchar(80) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipodeorden", "usuariocodigo") Then
        VGcnx.Execute "ALTER TABLE co_tipodeorden add usuariocodigo nvarchar(8) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipodeorden", "fechaact") Then
        VGcnx.Execute "ALTER TABLE co_tipodeorden add fechaact datetime NULL"
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "centrocostocodigo") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra add centrocostocodigo nvarchar(10) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "entidadcodigo") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra add entidadcodigo nvarchar(11) NULL"
End If
If Not ExisteElem(1, VGcnx, "kardexaux", "alma") Then
        VGcnx.Execute "ALTER TABLE kardexaux add alma nvarchar(02) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "fam_codigo") Then
   VGcnx.Execute "ALTER TABLE co_detordcompra add fam_codigo nvarchar(10) NULL"
End If
If Not ExisteElem(0, VGcnx, "co_estadoRequerimiento") Then
   SQL = "CREATE TABLE [co_estadorequerimiento] (estadooccodigo nvarchar(1) NOT NULL ,"
   SQL = SQL & "estadoocdescripcion nvarchar(40) NOT NULL ,"
   SQL = SQL & "estadoocatendido bit NOT NULL ,"
   SQL = SQL & "NivelRequerimientocodigo nvarchar(1) not null "
   SQL = SQL & "CONSTRAINT PK_co_estadorequerimiento "
   SQL = SQL & " PRIMARY KEY  CLUSTERED(estadooccodigo))"
   VGcnx.Execute (SQL)
End If
If Not ExisteElem(0, VGcnx, "co_NivelRequerimiento") Then
   SQL = "CREATE TABLE [co_NivelRequerimiento] (NivelRequerimientocodigo nvarchar(1) NOT NULL ,"
   SQL = SQL & "NivelRequerimientodescripcion nvarchar(20) NOT NULL "
   SQL = SQL & "CONSTRAINT PK_co_NivelRequerimiento"
   SQL = SQL & " PRIMARY KEY  CLUSTERED(NivelRequerimientocodigo))"
   VGcnx.Execute (SQL)
End If

If Not ExisteElem(1, VGcnx, "co_estadoRequerimiento", "NivelRequerimientocodigo") Then
   VGcnx.Execute "ALTER TABLE co_estadoRequerimiento add NivelRequerimientocodigo nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "vt_parametroventa", "codigotransaccionventa") Then
   VGcnx.Execute "ALTER TABLE vt_parametroventa add codigotransaccionventa nvarchar(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_tipodeorden", "AprobacionGerencia") Then
        VGcnx.Execute "ALTER TABLE co_tipodeorden add AprobacionGerencia bit"
End If
If Not ExisteElem(1, VGcnx, "co_estadoRequerimiento", "Nivelsiguientedeaprobacion") Then
        VGcnx.Execute "ALTER TABLE co_estadoRequerimiento add Nivelsiguientedeaprobacion char(1)"
End If
If Not ExisteElem(1, VGcnx, "co_NivelRequerimiento", "NivelaprobacionGerencia") Then
        VGcnx.Execute "ALTER TABLE co_NivelRequerimiento add NivelaprobacionGerencia bit "
End If
If Not ExisteElem(1, VGcnx, "co_NivelRequerimiento", "Nivelaprobacionmantenimiento") Then
        VGcnx.Execute "ALTER TABLE co_NivelRequerimiento add Nivelaprobacionmantenimiento bit "
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "estadooccodigo") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra add estadooccodigo char(1) "
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "oc_estadoorden") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra add oc_estadoorden char(1) "
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "fechaanulacion") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra add fechaanulacion datetime null "
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagventas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagventas nvarchar(1) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresabaseventas") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresabaseventas nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "NOMBREGUIAREMISION") Then
        VGConfig.Execute "ALTER TABLE empresa ADD  NOMBREGUIAREMISION NVARCHAR(15) NULL"
End If
If Not ExisteElem(1, VGConfig, "empresa", "NOMBREFACTURA") Then
        VGConfig.Execute "ALTER TABLE empresa ADD  NOMBREFACTURA NVARCHAR(15) NULL"
End If
If Not ExisteElem(1, VGcnx, "vt_pedido", "pedidotiporefe") Then
        VGcnx.Execute "ALTER TABLE vt_pedido add pedidotiporefe NVARCHAR(2) NULL "
End If
If Not ExisteElem(1, VGcnx, "vt_pedido", "pedidonrorefe") Then
        VGcnx.Execute "ALTER TABLE vt_pedido add pedidonrorefe NVARCHAR(11) NULL "
End If
If Not ExisteElem(1, VGcnx, "vt_pedido", "transportecodigo") Then
        VGcnx.Execute "ALTER TABLE vt_pedido add transportecodigo NVARCHAR(11) NULL "
End If
If Not ExisteElem(0, VGcnx, "al_sistema") Then
   SQL = " drop table configuracion "
   VGcnx.Execute SQL
   SQL = " CREATE TABLE [al_sistema] ( "
   SQL = SQL & " [Almacenpredeterminado] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,"
   SQL = SQL & " [cuentacostoven_debe] [nvarchar] (10) , [cuentacostoven_Habe] [nvarchar] (10) , "
   SQL = SQL & " [tipoalmacen] [nvarchar] (1),[valorizaliqcompra] [nvarchar] (1)) ON [PRIMARY]  "
   VGcnx.Execute SQL
   SQL = " update movalmcab set cacodmon='01' where cacodmon='MN'"
   VGcnx.Execute SQL
   SQL = " update movalmcab set cacodmon='02' where cacodmon='ME'"
   VGcnx.Execute SQL
End If
If Not ExisteElem(1, VGcnx, "familia", "gastoscodigo") Then
        VGcnx.Execute "ALTER TABLE familia add gastoscodigo NVARCHAR(20) NULL "
End If
If ExisteElem(1, VGcnx, "co_detordcompra", "oc_cunidad") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra alter column oc_cunidad NVARCHAR(6) NULL "
End If
If ExisteElem(1, VGcnx, "co_detordcompra", "oc_cuniref") Then
        VGcnx.Execute "ALTER TABLE co_detordcompra alter column oc_cuniref NVARCHAR(6) NULL "
End If
If Not ExisteElem(1, VGcnx, "movalmcab", "cabprovinumero") Then
        VGcnx.Execute "ALTER TABLE movalmcab add  cabprovinumero integer NULL "
End If
If Not ExisteElem(1, VGcnx, "movalmcab", "estadoprovision") Then
        VGcnx.Execute "ALTER TABLE movalmcab add  estadoprovision bit NULL "
End If
If ExisteElem(1, VGcnx, "movalmcab", "canroped") Then
        VGcnx.Execute "ALTER TABLE movalmcab alter column  canroped nvarchar(11) NULL "
End If
If ExisteElem(1, VGcnx, "te_saldoini", "codcaja") Then
        VGcnx.Execute "ALTER TABLE te_saldoini drop column  codcaja "
End If
If ExisteElem(1, VGcnx, "te_saldoini", "codBanco") Then
        VGcnx.Execute "ALTER TABLE te_saldoini drop column  codBanco "
End If
If Not ExisteElem(1, VGcnx, "te_saldoini", "codCajaBanco") Then
        VGcnx.Execute "ALTER TABLE te_saldoini add  codCajaBanco nvarchar(2) null "
End If
If Not ExisteElem(0, VGcnx, "al_kardex_cc") Then
        VGcnx.Execute "select top 0 * into al_kardex_cc from al_kardex_val"
End If
If Not ExisteElem(1, VGcnx, "al_kardex_cc", "centrocostocodigo") Then
        VGcnx.Execute "alter table al_kardex_cc add centrocostocodigo nvarchar(10) NULL "
End If
If Not ExisteElem(1, VGcnx, "al_sistema", "permiterequerimientos") Then
        VGcnx.Execute "alter table al_sistema add permiterequerimientos bit NULL "
End If
If Not ExisteElem(1, VGcnx, "al_sistema", "permiteIngresosconrequerimientos") Then
        VGcnx.Execute "alter table al_sistema add permiteIngresosconrequerimientos bit NULL "
End If
If Not ExisteElem(1, VGcnx, "co_detordcompra", "ordenreferencia") Then
        VGcnx.Execute "alter table co_detordcompra add ordenreferencia nvarchar(20) NULL "
End If
If Not ExisteElem(1, VGConfig, "empresa", "empresaflagcontabilidad") Then
        VGConfig.Execute "ALTER TABLE empresa ADD empresaflagcontabilidad nvarchar(1) NULL"
End If
If Not ExisteElem(0, VGcnx, "te_FormadePago") Then
   SQL = "CREATE TABLE [te_FormadePago] ([FormadePagocodigo] [char](2) NOT NULL ,"
   SQL = SQL & "[FormadePagodescripcion] [char](50) NOT NULL , [FormadePagodesccorta] [char](30) NOT NULL ,"
   SQL = SQL & "[FormadePagotipo] [char] (1) NULL ,[FormadePagoingplan] [char] (1) NULL ,"
   SQL = SQL & "[FormadePagoingcobra] [char] (1) NULL ,[FormadePagopermiteaplica] [char](1) NULL ,    [FormadePagorenovarletras] [char] (1)   NULL ,"
   SQL = SQL & "[FormadePagodocrenovaletra] [char](1) NULL,[FormadePagovalidabanco] [char](1) NULL,[FormadePagonumeauto] [char] (1)   NULL ,"
   SQL = SQL & "[FormadePagonumerador] [varchar] (11) NULL,[FormadePagocuentasoles] [char](20) NULL,[FormadePagocuentadolares] [char] (20)   NULL ,"
   SQL = SQL & "[FormadePagoaplicadifcamb] [char] (1) NULL,[FormadePagonotaconta] [char](1) NULL,[FormadePagosunat] [char] (3)   NULL ,"
   SQL = SQL & "[usuariocodigo] [char](8) NOT NULL ,[fechaact] [datetime] NOT NULL ,[FormadePagocancela] [char] (1)   NULL ,"
   SQL = SQL & "[documentoretencion] [varchar](1) NULL,[cargoemitedetraccion] [varchar](1) NULL ,"
   SQL = SQL & "CONSTRAINT [PK_te_FormadePago] PRIMARY KEY  CLUSTERED ([FormadePagocodigo])  ON [PRIMARY])"
   VGcnx.Execute SQL
End If
If ExisteElem(1, VGcnx, "te_codigocaja", "cajacuentasoles") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja alter column cajacuentasoles nvarchar(20) NULL"
End If
If ExisteElem(1, VGcnx, "te_codigocaja", "cajacuentadolares") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja alter column  cajacuentadolares nvarchar(20) NULL"
End If
If ExisteElem(1, VGcnx, "te_conceptoCaja", "conceptocuentasoles") Then
        VGcnx.Execute "ALTER TABLE te_conceptoCaja alter column conceptocuentasoles nvarchar(20) NULL"
End If
If ExisteElem(1, VGcnx, "te_conceptoCaja", "conceptocuentadolar") Then
        VGcnx.Execute "ALTER TABLE te_conceptoCaja alter column  conceptocuentadolar nvarchar(20) NULL"
End If
If ExisteElem(1, VGcnx, "te_cuentabancos", "cbanco_cuenta") Then
        VGcnx.Execute "ALTER TABLE te_cuentabancos alter column  cbanco_cuenta nvarchar(20) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "controlaestadosrendicion") Then
        VGcnx.Execute "ALTER TABLE co_sistema add controlaestadosrendicion bit default 1"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "diasatrazorendicion") Then
        VGcnx.Execute "ALTER TABLE co_sistema add diasatrazorendicion integer default 0 "
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "diacierrerendicion") Then
        VGcnx.Execute "ALTER TABLE co_sistema add diacierrerendicion integer default 0"
End If
If Not ExisteElem(1, VGcnx, "te_rendiciones", "saldoacumuladoxrendir") Then
        VGcnx.Execute "ALTER TABLE te_rendiciones add saldoacumuladoxrendir float default 0"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "CajaCuentaxRendir") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja add CajaCuentaxRendir bit default 0"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "CajaSuspendida") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja add Cajasuspendida bit default 0"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "CajaFondofijo") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja add CajaFondofijo bit default 0"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "codigooperaciontransferencia") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD codigooperaciontransferencia VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "EstadoDocxRendir") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD EstadoDocxRendir VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "SaldoDocxRendir") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD SaldoDocxRendir float default 0"
End If
If Not ExisteElem(1, VGcnx, "cp_abono", "ComprobConta") Then
        VGcnx.Execute "ALTER TABLE cp_abono ADD ComprobConta varchar(10)"
End If
If Not ExisteElem(1, VGcnx, "cp_parametros", "contabilizaenlinea") Then
        VGcnx.Execute "ALTER TABLE cp_parametros ADD contabilizaenlinea varchar(1)"
End If
If Not ExisteElem(1, VGcnx, "vt_parametroventa", "tiporedondeocobrar") Then
        VGcnx.Execute "ALTER TABLE vt_parametroventa ADD tiporedondeocobrar integer default 1"
End If
If Not ExisteElem(1, VGcnx, "cp_tipodocumento", "tdocumentoactualizaxtesoreria") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ADD tdocumentoactualizaxtesoreria integer default 1"
End If
If Not ExisteElem(1, VGcnx, "tabalm", "almacenvalorizado") Then
        VGcnx.Execute "ALTER TABLE tabalm ADD almacenvalorizado integer default 0"
End If
If Not ExisteElem(1, VGcnx, "ct_operacion", "operaciondocumentoanulado") Then
        VGcnx.Execute "ALTER TABLE ct_operacion ADD operaciondocumentoanulado bit  default 0"
End If
If Not ExisteElem(0, VGcnx, "al_cierresmensuales") Then
        VGcnx.Execute "select * into al_cierresmensuales from CIERRMESVALOR"
        VGcnx.Execute ("drop table CIERRMESVALOR")
End If
Exit Sub
ERROR1:
MsgBox "Ocurrio un Error," & error & " debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub

Public Property Get ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Property
    End If
    ipos = InStr(sName, Chr$(0))
    ComputerName = "##" + Left$(sName, ipos - 1)
End Property
Public Sub central(f As Form)
    f.Left = (Screen.Width - f.Width) / 2
    f.Top = (Screen.Height / 1.19 - f.Height)
End Sub

Public Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Public Function Existe(tipo As Integer, Cod As String, TABLA As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & TABLA & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(TABLA) = "PUNTO_VENTA" Then
                cSL = "Select * from " & TABLA & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & TABLA & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGConfig, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGCnxCT, adOpenStatic
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
 xRuc = Trim(xRuc)
 
' If xRuc <> " " Then
  'If xRuc <> "00000002" Then
     If Len(RTrim(xRuc)) < 11 Then
         MsgBox "Número de R.U.C. no tiene 11 dígitos", vbExclamation, "Ingreso de Datos"
         flag = False
      Else
'         nX = 0
'         NR = 0
'         NY = 0
'         CadNR = ""
'         For i = 1 To 7
'             nX = nX + Val(Mid(xRuc, i, 1)) * TAB_VAL(i)
'         Next i
'         NY = nX \ 11
'         NR = 11 - (nX - (NY * 11))
'         CadNR = Trim(String(10 - Len(Str(NR)) + 1, "0")) & Trim(Str(NR))
'         If Mid(CadNR, 10, 1) = Mid(xRuc, 8, 1) Then
'            flag = True
''         Else
'            MsgBox "Número de R.U.C. invalido", vbExclamation, "Ingreso de Datos"
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
'para Grabarla en una instrucción SQL
'*************************************************
Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Public Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional Titulo As String)
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport
        If Right(VGParamSistem.RutaReport, 1) <> "\" Then
           .ReportFileName = VGParamSistem.RutaReport & "\"
        End If
        .ReportFileName = .ReportFileName & VGParamSistem.carpetareportes
        If Right(.ReportFileName, 1) <> "\" Then
        .ReportFileName = .ReportFileName & "\"
        End If
        .ReportFileName = .ReportFileName & cNombreReporte
        .Connect = vgCADENAREPORT2
        .LogOnServer "pdsmon.dll", "", VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, VGParamSistem.PwdGEN
    
        .Formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
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
Dim valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        i = 0
        If pos = 0 Then Exit Do
        valor = Left(cad, pos - 1)
        cry.SortFields(i) = valor
        i = i + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Sub ImpresionRptbase(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional Titulo As String)
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport & "\" & cNombreReporte
        .LogOnServer "pdssql.dll", VGParamSistem.ServidorGEN, VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, ""
        .Connect = vgCADENAREPORT2
        .Formulas(0) = "@Emp='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
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

Sub ImpresionRpt_SubRpt_Proc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional ORDEN As String, Optional Titulo As String)
Dim strBuscar As New dll_apis
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        If Right(VGParamSistem.RutaReport, 1) <> "\" Then VGParamSistem.RutaReport = VGParamSistem.RutaReport & "\"
        .ReportFileName = VGParamSistem.RutaReport + cNombreReporte
        
        .LogOnServer "pdssql.dll", VGParamSistem.ServidorGEN, VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, ""
        .Connect = vgCADENAREPORT2
        .Formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
   '    .DiscardSavedData = True
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
        .LogOnServer "pdssql.dll", VGParamSistem.ServidorGEN, VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, ""
        .Connect = vgCADENAREPORT2
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
Public Function XRecuperaTipoCambio(Fecha As String, tipo As tipocambio, cnx As ADODB.Connection) As Double
Dim rsaux As ADODB.Recordset
Set rsaux = New ADODB.Recordset
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
    rsaux.Open "Select Valor=isnull(" & Campo & ",0)  from ct_tipocambio where tipocambiofecha ='" & Fecha & "'", cnx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        XRecuperaTipoCambio = rsaux!valor
    End If
End Function
Public Function ExisteSQL(ByVal cnx As ADODB.Connection, ByVal SentenciaSQL As String) As Boolean
On Error GoTo SaliError
    Screen.MousePointer = 11
    ExisteSQL = False
    Dim rsaux As ADODB.Recordset
    Set rsaux = New ADODB.Recordset
    rsaux.Open SentenciaSQL, cnx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        ExisteSQL = True
    End If
    Screen.MousePointer = 1
    Exit Function
SaliError:
    Screen.MousePointer = 1
    ExisteSQL = False
    MsgBox Err.Description
 '   Resume
End Function

Public Sub ADOCONECTAR()
On Error GoTo error

Set VGGeneral = New ADODB.Connection  'BD. ConfigFac
VGGeneral.CursorLocation = adUseClient
VGGeneral.CommandTimeout = 0
VGGeneral.ConnectionTimeout = 200
VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.ServidorGEN
VGGeneral.Open

'Conexion de Compras el Principal

Set VGcnx = New ADODB.Connection
VGcnx.CursorLocation = adUseClient
VGcnx.CommandTimeout = 0
VGcnx.ConnectionTimeout = 0
VGcnx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
VGcnx.Open
    
   
'Conexion de Cofiguracion

Set VGConfig = New ADODB.Connection
VGConfig.CursorLocation = adUseClient
VGConfig.CommandTimeout = 0
VGConfig.ConnectionTimeout = 0
VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
VGConfig.Open
    
'Conexion de Contabilidad

Set VGCnxCT = New ADODB.Connection
VGCnxCT.CursorLocation = adUseClient
VGCnxCT.CommandTimeout = 0
VGCnxCT.ConnectionTimeout = 0
VGCnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
VGCnxCT.Open
    
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
          fecha1 = Format("01/" & Format(Month(fecha1), "00") & "/" & Year(fecha1), "dd/mm/yyyy")
End Select
End Function

Public Function ESNULO(EXPRESION As Variant, valor As Variant) As Variant
On Error GoTo errfun
   If IsNull(EXPRESION) Or Trim(EXPRESION) = Empty Then
      ESNULO = valor
     Else: ESNULO = EXPRESION
   End If
   Exit Function
errfun:
   ESNULO = 0
End Function
Public Function ExisteElem(Tip As Integer, VGcnx As ADODB.Connection, TABLA As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim rsaux As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   TABLA = UCase(TABLA): Campo = UCase(Campo)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & TABLA
        Case 1:
            SQL = "Select Top 1 " & Campo & " From " & TABLA
    End Select
    rsaux.Open SQL, VGcnx
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function



