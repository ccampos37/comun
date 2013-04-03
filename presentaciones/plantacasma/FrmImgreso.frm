VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmIngreso 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "SISTEMA INTEGRADO DE GESTION ADMINISTRATIVA"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   9840
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmImgreso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   7935
   ScaleLeft       =   2000
   ScaleMode       =   0  'User
   ScaleTop        =   2000
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frameinicio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ususarios"
      Height          =   2175
      Left            =   2640
      TabIndex        =   21
      Top             =   2160
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Niagara Engraved"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmImgreso.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FF8080&
         Caption         =   "&Aceptar"
         Height          =   795
         Left            =   720
         Picture         =   "FrmImgreso.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frameopciones 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H000000FF&
      Height          =   7215
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   9615
      Begin VB.Frame FramePass 
         BackColor       =   &H00808000&
         Caption         =   "Case 7:"
         Height          =   2535
         Left            =   3240
         TabIndex        =   26
         Top             =   2520
         Width           =   5415
         Begin VB.Frame Frame1 
            Height          =   1215
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   3735
            Begin VB.OptionButton Optionempresas 
               Caption         =   "Administracion por Grupo de empresas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   435
               Left            =   0
               TabIndex        =   34
               Top             =   600
               Width           =   3615
            End
            Begin VB.OptionButton Optionaplicaciones 
               Caption         =   "Administracion Aplicaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   795
            Left            =   4200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmImgreso.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1440
            Width           =   900
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   795
            Left            =   4200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmImgreso.frx":0CD2
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox Text2 
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   8
            PasswordChar    =   "*"
            TabIndex        =   27
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808000&
            Caption         =   "Contraseña"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2775
         Left            =   2040
         Picture         =   "FrmImgreso.frx":1114
         ScaleHeight     =   2715
         ScaleWidth      =   3555
         TabIndex        =   4
         Top             =   2280
         Width           =   3615
         Begin VB.Line Line7 
            X1              =   360
            X2              =   240
            Y1              =   2640
            Y2              =   2760
         End
      End
      Begin Crystal.CrystalReport cryRpt 
         Left            =   120
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Line Line12 
         X1              =   1440
         X2              =   2400
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pymes"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   11
         Left            =   240
         MouseIcon       =   "FrmImgreso.frx":39D5
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   8040
         MouseIcon       =   "FrmImgreso.frx":3CDF
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   19
         Top             =   6240
         Width           =   135
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tesoreria"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   6720
         MouseIcon       =   "FrmImgreso.frx":3FE9
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3810
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Por Cobrar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   7
         Left            =   6720
         MouseIcon       =   "FrmImgreso.frx":42F3
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2970
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " Almacen"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   6600
         MouseIcon       =   "FrmImgreso.frx":45FD
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Por Pagar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   240
         MouseIcon       =   "FrmImgreso.frx":4907
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   5370
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Provisiones"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   4920
         MouseIcon       =   "FrmImgreso.frx":4C11
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilidad"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   5
         Left            =   3000
         MouseIcon       =   "FrmImgreso.frx":4F1B
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " Facturacion"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   6
         Left            =   1080
         MouseIcon       =   "FrmImgreso.frx":5225
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INTEGRASYSTEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -120
         MouseIcon       =   "FrmImgreso.frx":552F
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   6450
         Width           =   2895
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO ACUAPESCA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -240
         MouseIcon       =   "FrmImgreso.frx":5839
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2010.09"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -840
         MouseIcon       =   "FrmImgreso.frx":5B43
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   6720
         Width           =   4095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Activos Fiijos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   8
         Left            =   2880
         MouseIcon       =   "FrmImgreso.frx":5E4D
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Planillas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   10
         Left            =   5160
         MouseIcon       =   "FrmImgreso.frx":6157
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   3600
         X2              =   3600
         Y1              =   1440
         Y2              =   2280
      End
      Begin VB.Line Line3 
         X1              =   5040
         X2              =   5040
         Y1              =   1440
         Y2              =   2160
      End
      Begin VB.Line Line5 
         X1              =   6600
         X2              =   5640
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line6 
         X1              =   6600
         X2              =   5640
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line4 
         X1              =   6480
         X2              =   5640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2160
         Y1              =   1440
         Y2              =   2160
      End
      Begin VB.Line Line8 
         X1              =   2400
         X2              =   2400
         Y1              =   5040
         Y2              =   5400
      End
      Begin VB.Line Line9 
         X1              =   3720
         X2              =   3720
         Y1              =   5040
         Y2              =   5520
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Costos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   9
         Left            =   360
         MouseIcon       =   "FrmImgreso.frx":6461
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Line Line11 
         X1              =   1440
         X2              =   2040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line10 
         X1              =   5520
         X2              =   5535
         Y1              =   5040
         Y2              =   5415
      End
      Begin VB.Label LabelAdm 
         BackStyle       =   0  'Transparent
         Caption         =   "Administracion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4200
         MouseIcon       =   "FrmImgreso.frx":676B
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   6480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   9615
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "JCK Consultores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Index           =   7
         Left            =   120
         MouseIcon       =   "FrmImgreso.frx":6A75
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Telef.RPC 993900810/974989647-RPM *6906374 / Nextel : 51*115*5466 - 41*156*5229 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Index           =   8
         Left            =   1800
         MouseIcon       =   "FrmImgreso.frx":6D7F
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   210
         Width           =   7965
      End
   End
End
Attribute VB_Name = "FrmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uno As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If VERIFICAUSUARIO Then
 Frameopciones.Visible = True
 Frameinicio.Visible = False
 VGComputer = UCase$(ComputerName(1))
 Call enablelista
Else
  Text1.SetFocus
End If

End Sub

Private Sub Command1_Click()
Frameopciones.Visible = True
FramePass.Visible = False
End Sub

Private Sub Command2_Click()
Dim dato As Integer
dato = Day(Date) + Month(Date) * 2 + Year(Date) - 2000
If Text2.Text <> dato Then
   MsgBox " Contrasena no concuerda con validacion "
   Text2.SetFocus
 Else
   If Optionaplicaciones.Value = True Then
      FrmUsusuariosxsistema.Show
    Else
      FrmUsuariosxBasedatos.Show
   End If
   FramePass.Visible = False
   Frameopciones.Visible = True
   Call enablelista
End If
End Sub

Private Sub Form_Load()
uno = 0
Call adoconecta
Frameinicio.Visible = True
Frameopciones.Visible = False
FramePass.Visible = False
' Picture1.Picture = "ziyaz.JPG"


End Sub

Private Sub adoconecta()
'Conexion de Cofiguracion

Set VGConfig = New ADODB.Connection
VGConfig.CursorLocation = adUseClient
VGConfig.CommandTimeout = 0
VGConfig.ConnectionTimeout = 0
VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
VGConfig.Open

End Sub



Private Sub Label10_Click(Index As Integer)
If ExisteElem(0, VGConfig, VGComputer) Then
   VGConfig.Execute ("DELETE " & VGComputer & " where tipodesistema=" & Index & "")
Else
   SQL = "create table " & VGComputer & " ( tipodesistema int , usuariocodigo varchar(8))"
   VGConfig.Execute (SQL)
End If
SQL = "insert " & VGComputer & " ( tipodesistema , usuariocodigo )"
SQL = SQL & " values (" & Index & ",'" & Text1.Text & "')"
VGConfig.Execute (SQL)

On Error GoTo Errores
Select Case Index
Case 1:
    S = Shell(App.Path & "\" & VGalma, vbNormalFocus)
Case 2:
    S = Shell(App.Path & "\" & VGprovi, vbNormalFocus)
Case 3:
    S = Shell(App.Path & "\" & VGpaga, vbNormalFocus)
Case 4:
    S = Shell(App.Path & "\" & VGTeso, vbNormalFocus)
Case 5:
    S = Shell(App.Path & "\" & VGconta, vbNormalFocus)
Case 6:
    S = Shell(App.Path & "\" & VGfactu, vbNormalFocus)
Case 7:
    S = Shell(App.Path & "\" & VGcte, vbNormalFocus)
End Select

's = Shell(App.Path & Ejecutable, vbNormalFocus)


Exit Sub
Errores:
MsgBox "Error Nro: " & Err.Number & Chr(13) & Err.Description, vbCritical, "Error Sistemas"

End Sub

Private Sub Label9_Click()
Unload Me
End Sub
Private Function VERIFICAUSUARIO() As Boolean
    Dim RSPASS As New ADODB.Recordset
      
    'cuando no existe usuarios
    VERIFICAUSUARIO = False
   'VALIDANDO SI EXISTE EL USUARIO
    Set RSPASS = New ADODB.Recordset
    Set RSPASS = VGConfig.Execute("SELECT * FROM si_usuario")
    If RSPASS.RecordCount = 0 Then
       VERIFICAUSUARIO = True
       Exit Function
    End If
    Set RSPASS = New ADODB.Recordset
    SQL = "SELECT * FROM si_usuario wHERE USUarioCODIGO='" & UCase$(Text1.Text) & "'"
    Set RSPASS = VGConfig.Execute(SQL)
    If RSPASS.RecordCount = 0 Then
        MsgBox "NO SE ENCUENTRA EL USUARIO ", vbExclamation
        Text1.SetFocus
        Exit Function
    End If
    VERIFICAUSUARIO = True
End Function
Private Sub enablelista()
Dim rsql As New ADODB.Recordset
SQL = " select * from si_sistemaxusuario where usuariocodigo='" & Text1.Text & "'"
Set rsql = VGConfig.Execute(SQL)
If rsql.RecordCount > 0 Then rsql.MoveFirst
Do While Not rsql.EOF
   Label10(rsql!tipodesistema).Enabled = True
   rsql.MoveNext
Loop
End Sub
Private Sub LabelAdm_Click()
FramePass.Visible = True
Optionaplicaciones.Value = True
Text2.Text = ""
End Sub


Private Sub Text1_LostFocus()
Text1.Text = RTrim(UCase$(Text1.Text))
End Sub
