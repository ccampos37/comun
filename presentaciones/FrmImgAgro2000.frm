VERSION 5.00
Begin VB.Form FrmImg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SISTEMA INTEGRADO DE GESTION ADMINISTRATIVA"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   11025
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmImgAgro2000.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   5910
   ScaleLeft       =   2000
   ScaleMode       =   0  'User
   ScaleTop        =   2000
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   10575
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "MOVISTAR 990381193/974989647-*6906374"
         DataSource      =   "R "
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
         Height          =   255
         Index           =   19
         Left            =   6600
         MouseIcon       =   "FrmImgAgro2000.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   4095
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
         ForeColor       =   &H00004000&
         Height          =   255
         Index           =   17
         Left            =   120
         MouseIcon       =   "FrmImgAgro2000.frx":0316
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "RPC 993900810/ NEXTEL 41*156*5229"
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
         Height          =   255
         Index           =   18
         Left            =   2640
         MouseIcon       =   "FrmImgAgro2000.frx":0620
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2013.01"
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
         Left            =   2280
         MouseIcon       =   "FrmImgAgro2000.frx":092A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGRO 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
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
         MouseIcon       =   "FrmImgAgro2000.frx":0C34
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PYMESYSTEM"
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
         Left            =   0
         MouseIcon       =   "FrmImgAgro2000.frx":0F3E
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   6870
         MousePointer    =   99  'Custom
         Picture         =   "FrmImgAgro2000.frx":1248
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema Integrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   7800
         MouseIcon       =   "FrmImgAgro2000.frx":C6C2
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1410
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00B95017&
         BorderColor     =   &H80000009&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7830
         Shape           =   3  'Circle
         Top             =   3570
         Width           =   315
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
         Left            =   7920
         TabIndex        =   2
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8640
         MouseIcon       =   "FrmImgAgro2000.frx":C9CC
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   3600
         Width           =   495
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
Case 7:
    s = Shell(App.Path & "\" & VGPyme, vbNormalFocus)
End Select

's = Shell(App.Path & Ejecutable, vbNormalFocus)


Exit Sub
Errores:
MsgBox "Error Nro: " & Err.Number & Chr(13) & Err.Description, vbCritical, "Error Sistemas"

End Sub

Private Sub Label9_Click()
End
End Sub

