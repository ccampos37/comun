Attribute VB_Name = "ModLibroselectronicos"
Public Sub Generadiario(dato As String)
Dim RSQL As New ADODB.Recordset
Dim Archivo As String
Dim li_aRC As Integer
Archivo = NombreArchivoTxt(dato)
li_aRC = FreeFile
Open "C:\libroselectronicos\" & Archivo For Output As #li_aRC
If Mid$(Archivo, 30, 1) = "2" Then
   Close #li_aRC
   Exit Sub
End If
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "ct_libroDiario_rpt"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@cabcomprobmes") = VGParamSistem.Mesproceso
        Set RSQL = .Execute
End With
Call GeneraArchivoDiario(RSQL, Archivo, li_aRC)
End Sub

Public Sub GeneraArchivoDiario(rs As ADODB.Recordset, Archivo As String, li_arc1 As Integer)
Dim reg As String
Dim fecha As String
Dim n As Double
rs.MoveFirst
n = 0
Do While Not rs.EOF
   registro = ""
   With rs
     fecha = rs!cabcomprobfeccontable
     reg = Mid$(Archivo, 14, 8) + "|"
     reg = reg + rs!cabcomprobnumero
     reg = reg + "|01|" + rs!cuentacodigo + "|"
     reg = reg + fecha + "|" + rs!detcomprobglosa + "|"
     reg = reg + LTrim(Str(Round(rs!detcomprobdebe, 2))) + "|"
     reg = reg + LTrim(Str(Round(rs!detcomprobhaber, 2))) + "|1|"
   End With
   Print #li_arc1, reg
   n = n + 1
   rs.MoveNext
Loop
rs.Close
Close #li_arc1
Set rs = Nothing
Exit Sub
Error_PDT:
End Sub
Public Sub GeneradiarioSimplificado(dato As String)
Dim RSQL As New ADODB.Recordset
Dim Archivo As String
Archivo = NombreArchivoTxt(dato)
li_aRC = FreeFile
Open "C:\libroselectronicos\" & Archivo For Output As #li_aRC
If Mid$(Archivo, 30, 1) = "2" Then
   Close #li_aRC
   Exit Sub
End If
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "ct_libroDiario_rpt"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@cabcomprobmes") = VGParamSistem.Mesproceso
        Set RSQL = .Execute
End With
Call GeneraArchivoDiarioSimplificado(RSQL, Archivo)
End Sub
Public Sub GeneraArchivoDiarioSimplificado(rs As ADODB.Recordset, Archivo As String)
Dim li_aRC As Integer
Dim reg As String
Dim fecha As String
li_aRC = FreeFile
rs.MoveFirst
Do While Not rs.EOF
   registro = ""
   With rs
     fecha = rs!cabcomprobfeccontable
     reg = Mid$(Archivo, 14, 8) + "|" + rs!cabcomprobnumero + "|01|" + rs!cuentacodigo + "|"
     reg = reg + fecha + "|" + rs!detcomprobglosa + "|"
     reg = reg + LTrim(Str(Round(rs!detcomprobdebe, 2))) + "|"
     reg = reg + LTrim(Str(Round(rs!detcomprobhaber, 2))) + "|1|"
   End With
   Print #li_aRC, reg
   rs.MoveNext
Loop
rs.Close
Close #li_aRC
Set rs = Nothing
MsgBox "Se ha generado el archivo c:\telecredito\" & "0600" & LBLNUMERO & ".txt  satisfactoriamente ", vbInformation, "Mensaje"
Exit Sub
Error_PDT:
End Sub
Public Sub GeneraCompras(dato As String)
Dim RSQL As New ADODB.Recordset
Dim Archivo As String
Dim li_aRC As Integer
Archivo = NombreArchivoTxt(dato)
li_aRC = FreeFile
Open "C:\libroselectronicos\" & Archivo For Output As #li_aRC
If Mid$(Archivo, 30, 1) = "2" Then
   Close #li_aRC
   Exit Sub
End If
Set RSparCompras = New ADODB.Recordset
SQL = "select * from ct_paramlibaux where paramlibauxtipo='CO'"
Set RSparCompras = VGCNx.Execute(SQL)
If RSparCompras.RecordCount = 0 Then
   MsgBox "No existen parametros para el registros de compras"
   Exit Sub
End If
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "ct_LibroRegistroCompras_rpt"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@mes") = VGParamSistem.Mesproceso
        .Parameters("@ASIENTOSPLAN") = RSparCompras!paramlibauxasiento
        .Parameters("@CTASPLANCOMP") = RSparCompras!paramlibauxcuenta
        .Parameters("@CTASIGV") = RSparCompras!paramlibauxigv
        .Parameters("@CTASIES") = RSparCompras!paramlibauxies
        .Parameters("@CTASRENTA") = RSparCompras!paramlibauxirenta
        Set RSQL = .Execute
End With
Call GeneraArchivoCompras(RSQL, Archivo, li_aRC)
End Sub

Public Sub GeneraArchivoCompras(rs As ADODB.Recordset, Archivo As String, li_arc1 As Integer)
Dim reg As String
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
rs.MoveFirst
Do While Not rs.EOF
   registro = ""
   With rs
     dato3 = rs!detcomprobfechaemision
     reg = Mid$(Archivo, 14, 8) + "|" + rs!cabcomprobnumero + "|" + dato3 + "|"
     
     dato1 = rs!detcomprobfechavencimiento
     dato3 = IIf(rs!documentocodigo = "50", Right(rs!serie, 3), rs!serie)
     reg = reg + dato1 + "|" + rs!documentocodigo + "|" + dato3 + "|"
     
     'campos 7 para adelante
     dato1 = IIf(rs!documentocodigo = "50", VGParamSistem.Anoproceso, "0")
     dato3 = "0"
     reg = reg + dato1 + "|" + rs!detcomprobnumdocumento + "|" + dato3 + "|"
     
     dato1 = rs!identidadcodigo
     dato2 = ESNULO(rs!entidadruc, "-")
     If dato2 = "" Then dato2 = "-"
     dato3 = RTrim(rs!entidadrazonsocial)
     If dato3 = "" Then dato3 = "-"
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 13 para adelante
     dato1 = Round(!baseimpgrab, 2)
     dato2 = Round(!igvimpgrab, 2)
     reg = reg + dato1 + "|" + dato2 + "|" + "0.00" + "|"
     
     reg = reg + "0.00" + "|" + "0.00" + "|" + "0.00" + "|"
     
     'campos 19 para adelante
     dato1 = IIf(Round(!montoinafecto, 2) = 0, "0.00", Round(!montoinafecto, 2))
     reg = reg + dato1 + "|" + "0.00" + "|" + "0.00" + "|"
     
     dato1 = Round(!baseimpgrab, 2) + Round(!igvimpgrab, 2) + Round(!montoinafecto, 2)
     dato2 = Round(!detcomprobtipocambio, 3)
     If Len(dato2) = 4 Then dato2 = dato2 + "0"
     dato3 = ESNULO(!detcomprobfecharef, "01/01/0001")
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 25 para adelante
     dato1 = ESNULO(!tipdocref, "00")
     dato2 = IIf(dato1 = "00", "-", Left(detcomprobnumref, 4))
     dato3 = IIf(dato1 = "00", "-", Right(detcomprobnumref, 10))
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     dato1 = "-"
     dato2 = "01/01/0001"
     dato3 = "0"
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 31 para adelante
     dato1 = "0"
     dato2 = Format(Year(rs!detcomprobfechaemision), "0000") + Format(Month(rs!detcomprobfechaemision), "00")
     dato2 = IIf(dato2 < VGParamSistem.Anoproceso + VGParamSistem.Mesproceso, "6", "1")
     reg = reg + dato1 + "|" + dato2 + "|"
     
   End With
   Print #li_arc1, reg
   rs.MoveNext
Loop
rs.Close
Close #li_arc1
Set rs = Nothing
Exit Sub
Error_PDT:
End Sub
Public Sub GeneraVentas(dato As String)
Dim RSQL As New ADODB.Recordset
Dim Archivo As String
Dim li_aRC As Integer
Archivo = NombreArchivoTxt(dato)
li_aRC = FreeFile
Open "C:\libroselectronicos\" & Archivo For Output As #li_aRC
If Mid$(Archivo, 30, 1) = "2" Then
   Close #li_aRC
   Exit Sub
End If
Set RSparVentas = New ADODB.Recordset
RSparVentas.Open "select * from ct_paramlibaux where paramlibauxtipo='VT'", VGCNx, adOpenKeyset, adLockReadOnly
If RSparVentas.RecordCount = 0 Then
   MsgBox "No existen parametros para el registros de Ventas"
   Exit Sub
End If
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "ct_Libroregistroventas_rpt"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@MES") = VGParamSistem.Mesproceso
        .Parameters("@ASIENTOSPLAN") = RSparVentas!paramlibauxasiento
        .Parameters("@CTASPLANCOMP") = RSparVentas!paramlibauxcuenta
        .Parameters("@CTASIGV") = RSparVentas!paramlibauxigv
        .Parameters("@CTASFLETE") = RSparVentas!paramlibauxies
        .Parameters("@CTASOTROS") = RSparVentas!paramlibauxirenta
        .Parameters("@CTASDEVOL") = "74%"
        Set RSQL = .Execute
End With
Call GeneraArchivoVentas(RSQL, Archivo, li_aRC)
End Sub

Public Sub GeneraArchivoVentas(rs As ADODB.Recordset, Archivo As String, li_arc1 As Integer)
Dim reg As String
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim contador As Double
rs.MoveFirst
contador = 1
Do While Not rs.EOF
   
   registro = ""
   With rs
     dato1 = rs!detcomprobfechaemision
     reg = Mid$(Archivo, 14, 8) + "|" + rs!cabcomprobnumero + "|" + dato1 + "|"
     
     dato1 = rs!detcomprobfechavencimiento
     reg = reg + dato1 + "|" + rs!documentocodigo + "|" + rs!tdserie + "|"
     
     'campos 7 para adelante
     dato1 = Right(rs!detcomprobnumdocumento, 10)
     dato2 = "0"
     dato3 = rs!identidadcodigo
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     dato1 = ESNULO(rs!entidadruc, "-")
     dato2 = RTrim(rs!entidadrazonsocial)
     If dato2 = "" Then dato2 = "-"
     dato3 = "0.00"
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 13 para adelante
     dato1 = IIf(Round(!baseimponible, 2) = 0, "0.00", Round(!baseimponible, 2))
     dato2 = "0.00"
     dato3 = IIf(Round(!montoinafecto, 2) = 0, "0.00", Round(!montoinafecto, 2))
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     dato1 = "0.00"
     dato2 = IIf(Round(!igvimpgrab, 2) = 0, "0.00", Round(!igvimpgrab, 2))
     dato3 = "0.00"
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 19 para adelante
     dato1 = "0.00"
     dato2 = "0.00"
     dato3 = Round(!baseimponible, 2) + Round(!igvimpgrab, 2) + Round(!montoinafecto, 2)
     dato3 = IIf(dato3 = 0, "0.00", dato3)
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     dato1 = IIf(Round(!detcomprobtipocambio, 3) = 0, "0.000", Round(!detcomprobtipocambio, 3))
     If Len(dato1) = 4 Then dato1 = dato1 + "0"
     dato2 = ESNULO(Left(!documentoreferencia, 2), "00")
     dato2 = IIf(dato2 = "00", "01/01/0001", !detcomprobfecharef)
     dato3 = ESNULO(Left(!documentoreferencia, 2), "00")
     dato3 = IIf(dato3 = "", "00", dato3)
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
     'campos 25 para adelante
     
     dato1 = ESNULO(Left(!documentoreferencia, 2), "00")
     dato1 = IIf(dato1 = "00", "-", Left(detcomprobnumref, 4))
     dato2 = ESNULO(Left(!documentoreferencia, 2), "00")
     dato2 = IIf(dato2 = "00", "-", Right(detcomprobnumref, 10))
     dato3 = IIf(!operaciondocumentoanulado = 1, "2", "1")
     reg = reg + dato1 + "|" + dato2 + "|" + dato3 + "|"
     
   End With
   contador = contador + 1
   If contador = 16 Then
  '    MsgBox ("error")
   End If
   Print #li_arc1, reg
   rs.MoveNext

Loop
rs.Close
Close #li_arc1
Set rs = Nothing
Exit Sub
Error_PDT:
End Sub

Public Function NombreArchivoTxt(dato1 As String)
Dim rsql2 As New ADODB.Recordset
Dim nombre As String
Dim dia As String
Dim codoportun As String
Dim llenadato As String
Set rsql2 = Nothing
Set rsql2 = VGCNx.Execute("select * from ct_librossunatcorrelativos where librocodigosunat='" & dato1 & "'")
If rsql2!diaproceso = "DD" Then
   dia = Left(fecha(1, "01/" + VGParamSistem.Mesproceso + "/" + VGParamSistem.Anoproceso + ""), 2)
Else
   dia = rsql2!diaproceso
End If
If rsql2!codigoOportunidad = "CC" Then
   codoportun = Left(fecha(1, "01/" + VGParamSistem.Mesproceso + "/" + VGParamSistem.Anoproceso + ""), 2)
Else
   codoportun = rsql2!codigoOportunidad
End If
nombre = rsql2!identificadorGenerico & VGParametros.RucEmpresa & VGParamSistem.Anoproceso
nombre = nombre & VGParamSistem.Mesproceso & dia & dato1 & codoportun
Set rsql2 = Nothing
SQL = "select * from ct_librosSunatxempresa where empresacodigo='" & VGParametros.empresacodigo & "'"
SQL = SQL & " and librocodigosunat='" & dato1 & "'"
Set rsql2 = VGCNx.Execute(SQL)
If rsql2.RecordCount = 0 Then
   nombre = nombre + "2011.txt"
 Else
   nombre = nombre + "1111.txt"
End If
NombreArchivoTxt = nombre
End Function


