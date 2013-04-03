VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   Icon            =   "Frmsedacusco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   7275
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   480
      Picture         =   "Frmsedacusco.frx":324A
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   11
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1050
      MouseIcon       =   "Frmsedacusco.frx":1009F
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5790
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00B95017&
      BorderColor     =   &H80000009&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   600
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   315
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "SEDACUSCO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   14
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "EMPRESA PRESTADORA DE SERVICIOS DE SANEAMIENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   8415
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8655
      Y1              =   960
      Y2              =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Telef. 6277657 / 7855381 / 995304767 / NEXTEL 115*5466"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Index           =   8
      Left            =   3240
      MouseIcon       =   "Frmsedacusco.frx":103A9
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   6480
      Width           =   5205
   End
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
      ForeColor       =   &H00008000&
      Height          =   360
      Index           =   7
      Left            =   360
      MouseIcon       =   "Frmsedacusco.frx":106B3
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6480
      Width           =   1965
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   8655
      Y1              =   6360
      Y2              =   6375
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   8655
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   960
      Y2              =   7080
   End
   Begin VB.Line Line2 
      X1              =   8640
      X2              =   8655
      Y1              =   960
      Y2              =   6975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Recaudacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   6000
      MouseIcon       =   "Frmsedacusco.frx":109BD
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":10CC7
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspectoria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":10FD1
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Principales clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":112DB
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":115E5
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comercializacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":118EF
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Reclamos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6120
      MouseIcon       =   "Frmsedacusco.frx":11BF9
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "      Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      MouseIcon       =   "Frmsedacusco.frx":11F03
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   750
      TabIndex        =   0
      Top             =   5100
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   585
      Left            =   5190
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":1220D
      Stretch         =   -1  'True
      Top             =   990
      Width           =   675
   End
   Begin VB.Image Image6 
      Height          =   585
      Left            =   5190
      MouseIcon       =   "Frmsedacusco.frx":1D687
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":1D991
      Stretch         =   -1  'True
      Top             =   1650
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   585
      Left            =   5190
      MouseIcon       =   "Frmsedacusco.frx":35DC9
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":360D3
      Stretch         =   -1  'True
      Top             =   2370
      Width           =   675
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   5070
      MouseIcon       =   "Frmsedacusco.frx":45BCC
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":45ED6
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   5190
      MouseIcon       =   "Frmsedacusco.frx":5531F
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":55629
      Stretch         =   -1  'True
      Top             =   5430
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   5190
      MouseIcon       =   "Frmsedacusco.frx":68AED
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":68DF7
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   5190
      MouseIcon       =   "Frmsedacusco.frx":7EBB1
      MousePointer    =   99  'Custom
      Picture         =   "Frmsedacusco.frx":7EEBB
      Stretch         =   -1  'True
      Top             =   4590
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B95017&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   660
      Shape           =   3  'Circle
      Top             =   5070
      Width           =   315
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Ejecutable As String

Private Sub Form_Load()
On Error GoTo Err

' Sw.Movie = App.Path & "\gremco.swf"

Exit Sub
Err:
MsgBox "" & Err.Number & Chr(13) & Err.Description

End Sub

Private Sub Label10_Click(Index As Integer)
On Error GoTo Errores
Select Case Index


Case 0:
    s = Shell(App.Path & "\" & VGReclamos, vbNormalFocus)
Case 1:
    s = Shell(App.Path & "\" & VGComercializacion, vbNormalFocus)
Case 2:
    s = Shell(App.Path & "\" & VGMedicion, vbNormalFocus)
Case 3:
    s = Shell(App.Path & "\" & VGprincipales, vbNormalFocus)
Case 4:
    s = Shell(App.Path & "\" & VGInspectoria, vbNormalFocus)
Case 5:
    s = Shell(App.Path & "\" & VGfacturacion, vbNormalFocus)
Case 6:
    s = Shell(App.Path & "\" & VGrecaudacion, vbNormalFocus)
End Select

's = Shell(App.Path & Ejecutable, vbNormalFocus)
Unload Me

Exit Sub
Errores:
MsgBox "Error Nro: " & Err.Number & Chr(13) & Err.Description, vbCritical, "Error Sistemas"

End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label9_Click()
End
End Sub
