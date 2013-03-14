Attribute VB_Name = "ModificarCampos"
Option Explicit
Public NombreImpresora As String
Public EsFactura As Boolean

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Enum TIPOSISTEMA
   inventarios = 1
   compras = 2
   pagar = 3
   caja = 4
   contab = 5
   facturacion = 6
End Enum
Public VGsql As String * 1
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
End Sub
Public Sub adicionarcampos()
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
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "kitvirtual") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD kitvirtual bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "tabtransa", "ingresosfuturos") Then
        VGCNx.Execute "ALTER TABLE tabtransa ADD ingresosfuturos bit default('0') "
End If
If Not ExisteElem(1, VGCNx, "co_sistema", "codigopercepcion") Then
        VGCNx.Execute "ALTER TABLE co_sistema ADD codigopercepcion nvarchar(20) default('00') "
End If
If Not ExisteElem(1, VGCNx, "vt_pedido", "tipocontactocodigo") Then
        VGCNx.Execute "ALTER TABLE vt_pedido ADD tipocontactocodigo nvarchar(2) "
End If
If Not ExisteElem(1, VGCNx, "vt_pedido", "profesionalcodigo") Then
        VGCNx.Execute "ALTER TABLE vt_pedido ADD profesionalcodigo nvarchar(11) "
End If
If Not ExisteElem(1, VGCNx, "movalmcab", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE movalmcab ADD empresacodigo nvarchar(2) default('01')"
End If
If Not ExisteElem(1, VGCNx, "vt_puntoventa", "puntovtaticket") Then
        VGCNx.Execute "ALTER TABLE vt_puntoventa ADD puntovtaticket bit default('0')"
End If
If Not ExisteElem(1, VGCNx, "co_cabordcompra", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE co_cabordcompra ADD empresacodigo char(2)"
End If

If Not ExisteElem(1, VGCNx, "co_detordcompra", "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE co_detordcompra ADD empresacodigo char(2)"
End If

If Not ExisteElem(1, VGCNx, "co_detordcompra", "puntovtacodigo") Then
        VGCNx.Execute "ALTER TABLE co_detordcompra ADD puntovtacodigo char(2)"
End If
If Not ExisteElem(1, VGCNx, "vt_conceptosdepago", "bancocodigo") Then
        VGCNx.Execute "ALTER TABLE vt_conceptosdepago ADD bancocodigo char(2) default ('  ')"
End If
If Not ExisteElem(1, VGCNx, "vt_conceptostipodepago", "ctactecodigo") Then
        VGCNx.Execute "ALTER TABLE vt_conceptostipodepago ADD ctactecodigo char(20) default(' ')"
End If
If Not ExisteElem(1, VGCNx, "vt_parametroventa", "tesoreriaenlinea") Then
        VGCNx.Execute "ALTER TABLE vt_parametroventa ADD tesoreriaenlinea bit default('0')"
End If
If Not ExisteElem(1, VGCNx, "te_operaciongeneral", "facturacionactualiza") Then
        VGCNx.Execute "ALTER TABLE te_operaciongeneral ADD facturacionactualiza bit default('0')"
End If
If Not ExisteElem(1, VGCNx, "te_operaciongeneral", "provisionesactualiza") Then
        VGCNx.Execute "ALTER TABLE te_operaciongeneral ADD provisionesactualiza bit default('0')"
End If
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

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
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
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGconfig, adOpenStatic
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

Public Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim i As Integer
Dim RsImp As ADODB.Recordset

On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
                
        '----------------- AÑADIDO ---------------
        If EsFactura = True Then
            If NombreImpresora = "M" Or NombreImpresora = "T" Then
                Set RsImp = VGCNx.Execute("select impresoranombre from vt_configuraimpresora " _
                & " where empresacodigo='" & VGparametros.empresacodigo & "' and " _
                & " puntovtacodigo='" & VGparametros.puntovta & "' " _
                & " and impresoratipo='" & NombreImpresora & "' ")
                If RsImp.RecordCount > 0 Then
                    .PrinterName = RsImp(0)
                    .PrinterPort = Printer.Port
                    .PrinterDriver = "winspool"
                    .ProgressDialog = False
                    Set RsImp = Nothing
                Else
                    MsgBox "Aun no ha configurado la impresora.", vbCritical, "Sistema"
                    Screen.MousePointer = 1
                    Exit Sub
                End If
            End If
        End If
        '-------------------------------------
        
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
           .Connect = VGcadenareport2
         End If
           
        .formulas(0) = "@empresa='" & VGparametros.NomEmpresa & "'"
        .formulas(1) = "@ruc='" & VGparametros.RucEmpresa & "'"     'aki va el ruc
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
    NombreImpresora = Empty
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
           .Connect = VGcadenareport2
 
        End If
           
        .formulas(0) = "@Emp='" & VGparametros.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGparametros.RucEmpresa & "'"
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
    CrystalRpt.PageZoom (150)

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
        If Right(VGParamSistem.RutaReport, 1) <> "\" Then VGParamSistem.RutaReport = VGParamSistem.RutaReport & "\"
        .ReportFileName = VGParamSistem.RutaReport + cNombreReporte
        
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2

        End If
           
        .formulas(0) = "@Empresa='" & VGparametros.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGparametros.RucEmpresa & "'"
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
  '       .DiscardSavedData = True
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2

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
        XRecuperaTipoCambio = RSAUX!valor
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
Dim rsql As New ADODB.Recordset
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
VGconfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
VGconfig.Open
    
'Conexion de inventarios
If VGParamSistem.BDEmpresa = "" Or VGParamSistem.BDEmpresa = "?" Then
   Set rsql = VGconfig.Execute("select empresabaseinventarios from empresa where empresaflaginventarios=1")
   VGParamSistem.BDEmpresa = rsql!empresabaseinventarios
End If
Set VGCNx = New ADODB.Connection
VGCNx.CursorLocation = adUseClient
VGCNx.CommandTimeout = 0
VGCNx.ConnectionTimeout = 0
VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
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

If Trim(Campo) <> "" Then
    If Fecha = False Then
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


Public Function CODIFICA(cadena As String, valor As Integer) As String
    Dim ciclo As Integer, posic As Integer
    Dim utl_sal As Integer
    Dim carac As String, cadena_cod As String, cad As String
    posic = 0: utl_sal = 0
    carac = "": cadena_cod = "": cad = ""
    cadena = UCase(Trim(cadena))
    For ciclo = 1 To Len(cadena)
     carac = Mid(cadena, ciclo, 1)
     If (ciclo Mod 2) = 0 Then
      carac = UCase(carac)
     Else
      carac = LCase(carac)
     End If
     cadena_cod = cadena_cod & carac
    Next ciclo
    
    For ciclo = 1 To Len(cadena_cod)
     posic = ciclo Mod 7
     carac = Mid(cadena_cod, ciclo, 1)
     Select Case posic
     Case 0:
            carac = Chr(Asc(carac) * 2)
     Case 1:
            carac = Chr(Asc(carac) - valor)
     Case 2:
            carac = Chr(Asc(carac) - (ciclo * 2))
            utl_sal = Asc(carac)
     Case 3:
            If utl_sal > 10 Then utl_sal = utl_sal - (Int(utl_sal / 10) * 10)
            carac = Chr(Asc(carac) - valor + utl_sal)
     Case 4:
            carac = Chr(Asc(carac) - ciclo)
            utl_sal = Asc(carac)
     Case 5:
            If utl_sal > 10 Then utl_sal = utl_sal - (Int(utl_sal / 10) * 10)
            carac = Chr(Asc(carac) - valor + utl_sal)
     End Select
     cad = cad + carac
    Next ciclo
    CODIFICA = cad
End Function
'función que desencripta una cadena
Public Function DECODIFICA(cadena As String, valor As Integer) As String
    Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
    Dim carac As String, cad As String
    cadena = Trim(cadena)
    cad = ""
    val_n = 0: val_an = 0
    For ciclo = 1 To Len(cadena)
     carac = Mid(cadena, ciclo, 1)
     posic = ciclo Mod 7
     Select Case posic
     Case 0:
            val_n = Asc(carac) / 2
     Case 1:
            val_n = Asc(carac) + valor
     Case 2:
            val_n = Asc(carac) + (ciclo * 2)
            val_an = Asc(carac)
     Case 3:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + valor - val_an
     Case 4:
            val_n = Asc(carac) + ciclo
     Case 5:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + valor - val_an
     Case 6:
           val_n = Asc(carac)
     End Select
     cad = cad + Chr(val_n)
    Next ciclo
    DECODIFICA = UCase(cad)
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
Lit1 = "Uno    Dosc    Trec   Cuatroc  Quin   Seisc  Setec  Ochoc  Novec  "
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
     If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
        cWork = Mid(Trim(Str(num)), 1, InStr(1, Trim(Str(num)), ".", 0) - 1)
     Else
        cWork = Str(num)
     End If
     nMODULUS = Int(Len(Trim(cWork)) / 3)
     nMODULUS = Len(Trim(cWork)) - (nMODULUS * 3)
     
     If nMODULUS > 0 Then
        cWork = String(3 - nMODULUS, "0") & Trim(cWork)
     End If
     
     nK = (Len(Trim(cWork)) / 3) - 1
    'Procesa de Mil en Mil
     nI = 1
     Do While nI < Len(Trim(cWork)) - 1
        cCENTENA = Mid(Trim(cWork), nI, 1)
        cDECENA = Mid(Trim(cWork), nI + 1, 1)
        cUNIDAD = Mid(Trim(cWork), nI + 2, 1)
        'Centenas
        If cCENTENA <> "0" Then
            If cCENTENA = "1" Then
                cLET = cLET & "Cien "
                If cDECENA <> "0" Or cUNIDAD <> "0" Then
                    cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "to "
                End If
            Else
                cLET = cLET & Trim(Mid(Lit1, ((Val(cCENTENA) - 1) * 7) + 1, 7)) & "ientos "
            End If
        End If
        'Decenas
        If cDECENA <> "0" Then
            If cDECENA = "1" And cUNIDAD <> "0" Then
                If ((Val(cUNIDAD) - 1) * 10) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit3, ((Val(cUNIDAD) - 1) * 10) + 1, 10))
            Else
                If ((Val(cDECENA) - 1) * 9) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit2, ((Val(cDECENA) - 1) * 9) + 1, 9))
            End If
        End If
        'Unidades
        If cUNIDAD <> "0" Then
            If cDECENA > "1" Then
                cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "i"
                If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + LCase(Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6)))
            Else
                If cDECENA < "1" Then
                    If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6))
                End If
            End If
        End If
        cLET = cLET & " "
        'Pone Miles o Millones
        If nK > 0 Then
            If cCENTENA & cDECENA & cUNIDAD = "001" Then
                cLET = Mid(cLET, 1, Len(cLET) - 2) & " "
            End If
            nMODULUS = Int(nK / 2)
            nMODULUS = nK - (nMODULUS * 2)
            If nMODULUS = 0 Then
                cLET = cLET + Trim(Mid(Lit5, (((nK / 2) - 1) * 10) + 1, 10))
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
If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
    cLET = cLET + Mid(Trim(Str(num)), InStr(1, Trim(Str(num)), ".", 0) + 1, 2) & "/100" & " "
Else
    cLET = cLET + "00/100" & " "
End If
NUMLET = cLET
End Function


