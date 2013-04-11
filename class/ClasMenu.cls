VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public TablaUsu As String
Public TabaMenuDet As String
Public TablaMenu As String
Public Conexion As New ADODB.Connection
Dim Contr As Control
Public Sub CrearTablasSeguridad()
    Dim SQL As String
    'Crear Tabla Usuarios
    If Not ExisteElem(0, Conexion, TablaUsu) Then
        SQL = "" & _
        " Create Table " & TablaUsu & " ( " & _
        " usuariocodigo Text(5)  , Emp_Codigo Text(15) , UsuarioPassword Text(8), " & _
        " Usu_Nivel Text(1), usuarioNombre Text(32), " & _
        " CONSTRAINT Clave " & _
        " PRIMARY KEY (usuariocodigo, Emp_Codigo)) "
        Conexion.Execute SQL
    End If
    'Crea Tabla Detalle de Menu Usuarios
    If Not ExisteElem(0, Conexion, TabaMenuDet) Then
        SQL = "" & _
        " Create Table " & TabaMenuDet & " ( " & _
        " usuariocodigo Text(5)  , Emp_Codigo Text(15) , Men_Codigo Text(10), " & _
        " Men_Hab bit, " & _
        " CONSTRAINT Clave " & _
        " PRIMARY KEY (usuariocodigo, Emp_Codigo,Men_Codigo)) "
        Conexion.Execute SQL
    End If
End Sub
Public Sub CrearTablaMenu()
    Dim SQL As String
    Dim Contr As Control
    Dim rsaux As New ADODB.Recordset
    Dim ContNiv() As Integer
    Dim Niv As Integer, I As Integer, J As Integer
    Dim CodigoMenu As String
    On Error GoTo errortab
    '*------------------------------------------------------------------*
    If Not ExisteElem(0, VGConfig, TablaMenu) Then
        SQL = " Create Table " & TablaMenu & " (tipodesistema int,Men_Codigo varchar(20),Men_Descri varchar(64), " & _
        " Men_CodEdit varchar(48),  " & _
        " CONSTRAINT Clave PRIMARY KEY (tipodesistema,Men_Codigo))"
        VGConfig.Execute SQL
      Else:
        VGConfig.Execute "Delete From " & TablaMenu & " where tipodesistema=" & VGtipo
    End If
    'Redimensionando los Niveles
    ReDim ContNiv(MaxNivel + 1) 'MaxNivel : Funcion que determina el maximo nivel en el menu
    '*Armando El Menu
    Set rsaux = Nothing
    rsaux.Open TablaMenu, VGConfig, adOpenDynamic, adLockOptimistic
    For Each Contr In MDIPrincipal.Controls
        If TypeOf Contr Is Menu Then
            If Contr.Visible And RTrim$(Contr.Caption) <> "-" And UCase$(RTrim$(Contr.Caption)) <> "Salir" Then
                Niv = NumChar(Contr.Name, "_")
                For I = Niv To Niv
                    ContNiv(I) = ContNiv(I) + 1
                    For J = Niv + 1 To UBound(ContNiv) - 1
                        ContNiv(J) = 0
                    Next
                Next
                rsaux.AddNew
                CodigoMenu = ""
                For I = 0 To UBound(ContNiv) - 1
                    CodigoMenu = CodigoMenu & IIf(ContNiv(I) > 0, Format(ContNiv(I), "00"), "")
                Next
                rsaux!tipodesistema = VGtipo
                rsaux!Men_CodIGO = RTrim$(CodigoMenu)
                rsaux!Men_Descri = DevCad("&", Contr.Caption)
                rsaux!Men_CodEdit = RTrim$(Contr.Name)
                
                rsaux!Men_Visible = -1
                
                rsaux.Update
            End If
        End If
    Next
    Exit Sub
errortab:
'MsgBox "Error : " & Err.Number & "  " & Err.Description, vbExclamation
Resume Next
End Sub
Private Function MaxNivel() As Integer
    Dim Cont As Integer, Aux As Integer
    MaxNivel = 0
    '*---------------------------------*
    For Each Contr In MDIPrincipal.Controls
        If TypeOf Contr Is Menu Then
            If Contr.Visible And RTrim$(Contr.Caption) <> "-" And UCase$(RTrim$(Contr.Caption)) <> "SALIR" Then
                Aux = NumChar(Contr.Name, "_")
                If MaxNivel < Aux Then MaxNivel = Aux
            End If
        End If
    Next
End Function
Public Function NumChar(cadena As String, Caracter As String) As Integer
'Funcion Creada por Fernando Cossio
    Dim I As Integer, Cont As Integer
    Cont = 0
    For I = 1 To Len(cadena)
        If UCase$(Caracter) = Mid$(cadena, I, 1) Then
            Cont = Cont + 1
        End If
    Next
    NumChar = Cont
End Function
Public Function DevCad(CaraterNo As String, cadena As String)
    'Devuelve toda la cadena sin el caracter espcificado en el parametro
    'Funcion creada por fernando cossio
    Dim I As Integer, acum, Cont As String
  
    For I = 1 To Len(cadena)
        If UCase$(CaraterNo) <> Mid$(cadena, I, 1) Then
            acum = acum + Mid$(cadena, I, 1)
        End If
    Next
    DevCad = acum
End Function
Public Sub HabilitarMenuNom(Usuario As String)
'Procedimiento creado por Fernando Cossio
    Dim RsMenUsu As New ADODB.Recordset
    Dim SQL As String
    SQL = " SELECT a.usuariocodigo, b.Men_CodEdit, a.MEN_HAB From " & TabaMenuDet & " a "
    SQL = SQL & " left join " & TablaMenu & " b on a.tipodesistema=b.tipodesistema and a.MEN_CODIGO=b.Men_Codigo "
    SQL = SQL & " where a.usuariocodigo='" & Usuario & "' and a.tipodesistema=" & VGtipo & ""
    Set RsMenUsu = VGConfig.Execute(SQL)
    Dim n As String
    If RsMenUsu.RecordCount = 0 Then Exit Sub
    RsMenUsu.MoveFirst
    Do While Not RsMenUsu.EOF
        For Each Contr In MDIPrincipal.Controls
            If TypeOf Contr Is Menu Then
                n = Contr.Name
                If UCase$(n) = UCase$(ESNULO(RsMenUsu!Men_CodEdit, 1)) Then
                     Contr.Enabled = IIf(RsMenUsu!MEN_HAB, True, False)
                     Exit For
                End If
            End If
        Next
        RsMenUsu.MoveNext
    Loop
End Sub
Public Sub MenuTrue()
'Procedimiento creado por Fernando Cossio
    For Each Contr In MDIPrincipal.Controls
        If TypeOf Contr Is Menu Then
            Contr.Enabled = True
        End If
    Next
End Sub
