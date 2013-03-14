Attribute VB_Name = "Module1"
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
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.Anoproceso & "", "saldoacumdebe00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.Anoproceso & " ADD saldoacumdebe00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.Anoproceso & "", "saldoacumhaber00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.Anoproceso & " ADD saldoacumhaber00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.Anoproceso & "", "saldoacumussdebe00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.Anoproceso & " ADD saldoacumussdebe00 float default (0) "
   End If
    If Not ExisteElem(1, VGCNx, "ct_saldos" & VGParamSistem.Anoproceso & "", "saldoacumussHaber00") Then
        VGCNx.Execute "ALTER TABLE ct_saldos" & VGParamSistem.Anoproceso & " ADD saldoacumussHaber00 float default (0) "
   End If
If Not ExisteElem(1, VGCNx, "ct_ctacteanalitico" & VGParamSistem.Anoproceso & "", "ctacteanaliticoajustedifcambio") Then
        VGCNx.Execute "ALTER TABLE ct_ctacteanalitico" & VGParamSistem.Anoproceso & " ADD ctacteanaliticoajustedifcambio nvarchar(1) default('0') "
End If
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
On Error GoTo ERROR1
If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresaruc") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresaruc nvarchar(11) NULL"
   End If
   If Not ExisteElem(1, VGCNx, "co_multiempresas", "empresadireccion") Then
        VGCNx.Execute "ALTER TABLE co_multiempresas ADD empresadireccion nvarchar(50) NULL"
   End If
   If ExisteElem(1, VGCNx, "cc_tipodocumento", "tdocumentonumerador") Then
        VGCNx.Execute "ALTER TABLE cc_tipodocumento ALTER COLUMN tdocumentonumerador nvarchar(11) NULL"
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
If Not ExisteElem(1, VGconfig, "empresa", "multiguiasremision") Then
        VGconfig.Execute "ALTER TABLE empresa ADD multiguiasremision bit default('0')"
End If
If Not ExisteElem(1, VGconfig, "empresa", "multifacturas") Then
        VGconfig.Execute "ALTER TABLE empresa ADD multifacturas bit default('0') "
End If
If Not ExisteElem(1, VGconfig, "empresa", "multiboletas") Then
        VGconfig.Execute "ALTER TABLE empresa ADD multiboletas bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "maeart", "estadodetraccion") Then
        VGCNx.Execute "ALTER TABLE maeart ADD estadodetraccion bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "minimodetraccion") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD minimodetraccion float default('999999') "
End If
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "kitvirtual") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD kitvirtual bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "tabtransa", "ingresosfuturos") Then
        VGCNx.Execute "ALTER TABLE tabtransa ADD ingresosfuturos nvarchar(1) default('N') "
End If
If Not ExisteElem(1, VGCNx, "co_sistema", "codigopercepcion") Then
        VGCNx.Execute "ALTER TABLE co_sistema ADD codigopercepcion nvarchar(20) default('00') "
End If
If Not ExisteElem(1, VGCNx, "cc_tipoplanilla", "asientocodigo") Then
        VGCNx.Execute "ALTER TABLE cc_tipoplanilla ADD asientocodigo nvarchar(2) default('00') "
End If

If Not ExisteElem(1, VGCNx, "cp_tipoplanilla", "asientocodigo") Then
        VGCNx.Execute "ALTER TABLE cp_tipoplanilla ADD asientocodigo nvarchar(2) default('00') "
End If

If Not ExisteElem(1, VGCNx, "te_cuentabancos", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE te_cuentabancos ADD empresacodigo nvarchar(2) default('00') "
End If

If Not ExisteElem(1, VGCNx, "cp_tipoplanilla", "tplanillacompensaciones") Then
        VGCNx.Execute "ALTER TABLE cp_tipoplanilla ADD tplanillacompensaciones char(1) default('0') "
End If

If Not ExisteElem(1, VGCNx, "cc_tipoplanilla", "tplanillacompensaciones") Then
        VGCNx.Execute "ALTER TABLE cc_tipoplanilla ADD tplanillacompensaciones char(1) default('0') "
End If
Exit Sub
ERROR1:
MsgBox "Ocurrio un Error," & error & " debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub
Public Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = titulo
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
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = vgCADENAREPORT2
         End If
           
        .Formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"     'aki va el ruc
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
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
Dim pos As Integer, cadaux As String, I As Integer
Dim valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        I = 0
        If pos = 0 Then Exit Do
        valor = Left(cad, pos - 1)
        cry.SortFields(I) = valor
        I = I + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Sub ImpresionRptbase(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim I As Integer
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
           .Connect = vgCADENAREPORT2
 
        End If
           
        .Formulas(0) = "@Emp='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
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
    CrystalRpt.PageZoom (150)

End Sub

Sub ImpresionRpt_SubRpt_Proc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional ORDEN As String, Optional titulo As String)
Dim strBuscar As New dll_apis
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        If Right(VGParamSistem.RutaReport, 1) <> "\" Then VGParamSistem.RutaReport = VGParamSistem.RutaReport & "\"
        .ReportFileName = VGParamSistem.RutaReport + cNombreReporte
        
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = vgCADENAREPORT2

        End If
           
        .Formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
   '     .DiscardSavedData = True
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = vgCADENAREPORT2

        End If

        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
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

Public Sub ADOCONECTAR()
Dim RSQL As New ADODB.Recordset
On Error GoTo error

Set VGgeneral = New ADODB.Connection  'BD. ConfigFac
VGgeneral.CursorLocation = adUseClient
VGgeneral.CommandTimeout = 0
VGgeneral.ConnectionTimeout = 200
VGgeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.ServidorGEN
VGgeneral.Open

   
'Conexion de Cofiguracion

Set VGconfig = New ADODB.Connection
VGconfig.CursorLocation = adUseClient
VGconfig.CommandTimeout = 0
VGconfig.ConnectionTimeout = 0
VGconfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
VGconfig.Open
    
'Conexion de inventarios

If VGParamSistem.BDEmpresa = "" Or VGParamSistem.BDEmpresa = "?" Then
   Set RSQL = VGconfig.Execute("select empresabaseinventarios from empresa where empresaflaginventarios=1")
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

Public Function fecha(ByVal tipo As Integer, dato As Date) As Date
Dim fecha1 As Date
fecha1 = Format("01/" & Format(Month(dato), "00") & "/" & Year(dato), "dd/mm/yyyy")
Select Case tipo
        Case 1
          fecha = fecha1
        Case 2
          fecha1 = fecha1 + 31
          fecha1 = Format("01/" & Format(Month(fecha1), "00") & "/" & Year(fecha1), "dd/mm/yyyy")
          fecha = fecha1 - 1
        Case 3
          fecha1 = fecha1 - 31
          fecha = Format("01/" & Format(Month(fecha1), "00") & "/" & Year(fecha1), "dd/mm/yyyy")
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
Public Function ExisteElem(Tip As Integer, VGCNx As ADODB.Connection, Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim RSAUX As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   Tabla = UCase(Tabla): Campo = UCase(Campo)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & Tabla
        Case 1:
            SQL = "Select Top 1 " & Campo & " From " & Tabla
    End Select
    RSAUX.Open SQL, VGCNx
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function
Public Function DateSQL(ByVal fecha As String) As String
    'On Error GoTo ERR
    If IsNull(fecha) Then Exit Function
        Select Case VGformatofecha
            Case "DMY"
            DateSQL = "'" & Format(fecha, "dd/mm/yyyy") & "'"
            Case "MDY"
            DateSQL = "'" & Format(fecha, "mm/dd/yyyy") & "'"
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

Public Function Devolver_Dato(tipo As Integer, Cod As String, Tabla As String, Campo As String, fecha As Boolean, CampDev As String, Optional Cod2 As String, Optional Campo2 As String, Optional Cod3 As String, Optional Campo3 As String, Optional Cod4 As Double, Optional Campo4 As String) As String
Dim cSel1 As ADODB.Recordset, cF As String
Set cSel1 = New ADODB.Recordset

If Trim(Campo) <> "" Then
    If fecha = False Then
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  '" & Cod & "' "
    Else
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  #" & Format(Cod, "mm/dd/yyyy") & "#"
    End If
End If
If Trim(Campo2) <> "" Then
    cF = cF & " and " & Campo2 & " = '" & Cod2 & "' "
End If
If Trim(Campo3) <> "" Then
    cF = cF & " and " & Campo3 & " = '" & Cod3 & "' "
End If
If Trim(Campo4) <> "" Then
    cF = cF & " and " & Campo4 & " = '" & Cod4 & "' "
End If
Select Case tipo
  Case 1 'Bd. Comun
              cSel1.Open cF, VGCNx, adOpenStatic
  Case 2 'Bd. Config
              cSel1.Open cF, VGconfig, adOpenStatic
  Case 3 'Bd. Contabilidad
              cSel1.Open cF, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Devolver_Dato = IIf(Not IsNull(cSel1(0)), cSel1(0), "")
Else
     Devolver_Dato = ""
End If
End Function

