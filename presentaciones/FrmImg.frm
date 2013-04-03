VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmImg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "SISTEMA INTEGRADO DE GESTION ADMINISTRATIVA"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   9840
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmImg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   7470
   ScaleLeft       =   2000
   ScaleMode       =   0  'User
   ScaleTop        =   2000
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2175
      Left            =   1920
      TabIndex        =   21
      Top             =   2040
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2640
         Picture         =   "FrmImg.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   775
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   675
         Left            =   720
         Picture         =   "FrmImg.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1200
         Width           =   775
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H000000FF&
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9615
      Begin VB.PictureBox Picture1 
         Height          =   2775
         Left            =   2040
         Picture         =   "FrmImg.frx":0890
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
         MouseIcon       =   "FrmImg.frx":3151
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
         Left            =   6720
         MouseIcon       =   "FrmImg.frx":345B
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3810
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Por Cobrar"
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
         Left            =   6720
         MouseIcon       =   "FrmImg.frx":3765
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2970
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " Almacen"
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
         Left            =   6600
         MouseIcon       =   "FrmImg.frx":3A6F
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Por Pagar"
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
         MouseIcon       =   "FrmImg.frx":3D79
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   5370
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Provisiones"
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
         MouseIcon       =   "FrmImg.frx":4083
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilidad"
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
         Left            =   3000
         MouseIcon       =   "FrmImg.frx":438D
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " Facturacion"
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
         Index           =   0
         Left            =   1080
         MouseIcon       =   "FrmImg.frx":4697
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label Label10 
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
         Index           =   9
         Left            =   -120
         MouseIcon       =   "FrmImg.frx":49A1
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   6450
         Width           =   2895
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   -240
         MouseIcon       =   "FrmImg.frx":4CAB
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label10 
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
         Index           =   11
         Left            =   -840
         MouseIcon       =   "FrmImg.frx":4FB5
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   6720
         Width           =   4095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Activos Fiijos"
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
         Index           =   12
         Left            =   2880
         MouseIcon       =   "FrmImg.frx":52BF
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Planillas"
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
         Index           =   13
         Left            =   5160
         MouseIcon       =   "FrmImg.frx":55C9
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
         Index           =   14
         Left            =   360
         MouseIcon       =   "FrmImg.frx":58D3
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
      Begin VB.Label Label10 
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
         Index           =   15
         Left            =   4200
         MouseIcon       =   "FrmImg.frx":5BDD
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
      Top             =   6960
      Width           =   9615
      Begin VB.Label Label10 
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
         MouseIcon       =   "FrmImg.frx":5EE7
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label10 
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
         MouseIcon       =   "FrmImg.frx":61F1
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   210
         Width           =   7965
      End
   End
End
Attribute VB_Name = "FrmImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Ejecutable As String

Private Sub Form_Load()
On Error GoTo Err
Call ADOConectar

' Picture1.Picture = "ziyaz.JPG"

Exit Sub
Err:
MsgBox "" & Err.Number & Chr(13) & Err.Description

End Sub




Private Sub Label10_Click(Index As Integer)
On Error GoTo Errores
Select Case Index
Case 0:
    s = Shell(App.Path & "\" & VGfactu, vbNormalFocus)
Case 1:
    s = Shell(App.Path & "\" & VGconta, vbNormalFocus)
Case 2:
    s = Shell(App.Path & "\" & VGprovi, vbNormalFocus)
Case 3:
    s = Shell(App.Path & "\" & VGpaga, vbNormalFocus)
Case 4:
    s = Shell(App.Path & "\" & VGalma, vbNormalFocus)
Case 5:
    s = Shell(App.Path & "\" & VGcte, vbNormalFocus)
Case 6:
    s = Shell(App.Path & "\" & VGTeso, vbNormalFocus)
End Select

's = Shell(App.Path & Ejecutable, vbNormalFocus)


Exit Sub
Errores:
MsgBox "Error Nro: " & Err.Number & Chr(13) & Err.Description, vbCritical, "Error Sistemas"

End Sub

Private Sub Label9_Click()
End
End Sub

