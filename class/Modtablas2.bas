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
Public VGsql As String * 1
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum
Public Enum TipoCambio
    Compra = "01"
    Venta = "02"
    Promedio = "03"
End Enum


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

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If fecha Then
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
 Dim nX As Integer, NY As Integer, NR As Integer, I As Integer
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

Public Function XRecuperaTipoCambio(fecha As Date, tipo As TipoCambio, cnx As ADODB.Connection) As Double
Dim RSAUX As ADODB.Recordset
Set RSAUX = New ADODB.Recordset
Dim Campo As String
    XRecuperaTipoCambio = 1
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
    SQL = "Select Valor=isnull(" & Campo & ",1)  from ct_tipocambio where convert(varchar(10),tipocambiofecha,103) ='" & fecha & "'"
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

