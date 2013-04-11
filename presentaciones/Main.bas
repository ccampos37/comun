Attribute VB_Name = "Modulo"
Option Explicit

Public VGdllApi As dll_apisgen.dll_apis
Public VGfactu As String
Public VGconta As String
Public VGprovi As String
Public VGpaga As String
Public VGalma As String
Public VGcte As String
Public VGTeso As String
Public VGPyme As String
Public VGcosto As String
Public VGPlani As String
Public VGActi As String

Public rsql As New ADODB.Recordset
Public VGConfig As New ADODB.Connection
Public SQL As String
Public VGComputer As String

Public VgSalir As Integer

Public VGParamSistem As ParametrosdeSistema

Public Type ParametrosdeSistema
    Servidor As String
    BDEmpresa As String
    Usuario As String
    PWD      As String
    
    mesproceso As String
    Anoproceso As String
    
End Type
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub Main()
    'Base de Datos General
      
On Error GoTo Xmain
    'VGusuario = "03"
    'Leer Ini
    Set VGdllApi = New dll_apisgen.dll_apis
 
    VGfactu = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "factu", "?")
    VGconta = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "conta", "?")
    VGprovi = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "provi", "?")
    VGpaga = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "paga", "?")
    VGalma = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "alma", "?")
    VGcte = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "cte", "?")
    VGTeso = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Teso", "?")
    VGPyme = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Pyme", "?")
    VGcosto = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Costos", "?")
    VGActi = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "activos", "?")
   VGPlani = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Plani", "?")
  
  
 'Conexion de inventarios
VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
VGParamSistem.PWD = DECODIFICASQL(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")))


  
    FrmIngreso.Show
    Exit Sub
Xmain:
    MsgBox Err.Description, vbExclamation, "Error Main"
End Sub

Public Property Get ComputerName(Optional tipo As Integer) As Variant
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
    If tipo = 0 Then
       Randomize
       NombrePC = Trim$(Str(CLng(Rnd * 10000000)))
       ComputerName = "##" + Left$(sName, ipos - 1) + NombrePC
    ElseIf tipo = 1 Then
       ComputerName = "##" + Left$(sName, ipos - 1)
    Else
       ComputerName = Left$(sName, ipos - 1)
   End If
End Property
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

Public Function CODIFICASQL(cadena As String) As String
    Dim ciclo As Integer, posic As Integer
    Dim utl_sal As Integer
    Dim carac As String, cadena_cod As String, cad As String
    posic = 0: utl_sal = 0
    carac = "": cadena_cod = "": cad = ""
    cadena = UCase$(Trim$(cadena))
    For ciclo = 1 To Len(cadena)
     carac = Mid$(cadena, ciclo, 1)
     If (ciclo Mod 2) = 0 Then
      carac = UCase$(carac)
     Else
      carac = LCase(carac)
     End If
     cadena_cod = cadena_cod & carac
    Next ciclo
    
    For ciclo = 1 To Len(cadena_cod)
        posic = ciclo Mod 2
        carac = Mid$(cadena_cod, ciclo, 1)
        Select Case posic
        Case 0:
            carac = Chr(Asc(carac) * 2)
        Case 1:
            carac = Chr(Asc(carac) - 2)
        End Select
        cad = cad + carac
    Next ciclo
CODIFICASQL = cad
End Function

Public Function DECODIFICASQL(cadena As String) As String
    Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
    Dim carac As String, cad As String
    cadena = Trim$(cadena)
    cad = ""
    val_n = 0: val_an = 0
    For ciclo = 1 To Len(cadena)
        carac = Mid$(cadena, ciclo, 1)
        posic = ciclo Mod 2
        Select Case posic
        Case 0:
            val_n = Asc(carac) / 2
           cad = cad + LCase$(Chr(val_n))
        Case 1:
            val_n = Asc(carac) + 2
            cad = cad + Chr(val_n)
        End Select
    Next ciclo
    DECODIFICASQL = (cad)
End Function




