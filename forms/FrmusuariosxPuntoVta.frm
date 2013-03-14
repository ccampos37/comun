VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmusuariosxpuntovta 
   Caption         =   "Punto de Venta por usuario"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   4695
      Begin VB.CommandButton cmdsalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   735
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4695
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3855
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6800
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "usuariocodigo"
            Caption         =   "Cod. Usuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "usuarionombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2399.811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puntos de ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Frmusuariosxpuntovta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Check1_Click()
Dim rs2 As New ADODB.Recordset
SQL = "select a.puntovtacodigo,a.puntovtadescripcion,valor=0 from vt_puntoventa a"
Set rs2 = VGCNx.Execute(SQL)
   Call LlenarLista(rs2, 1)
End Sub

Private Sub CmdGrabar_Click()
Call grabar
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub



Private Sub Form_Load()
SQL = "select usuariocodigo, usuarionombre from si_usuario order by usuarionombre "
Set rs = Nothing
Set rs = VGConfig.Execute(SQL)
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
Private Sub LlenarLista(rss As ADODB.Recordset, Optional todo As Integer)
 Dim i As Integer
 Dim itmX As ListItem
 Dim rs2 As New ADODB.Recordset
   ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Punto de venta", ListView1.Width / 1
   ListView1.View = lvwReport
   rss.MoveFirst
   i = 1
   Do While Not rss.EOF
      Set itmX = ListView1.ListItems.Add(, , Str(i + 0) + "  " + rss!puntovtacodigo + "  " + rss!puntovtadescripcion)
      If todo = 0 Then
         Set rs2 = VGCNx.Execute(" select * from vt_usuarioxPuntoVta where usuariocodigo+ puntovtacodigo='" & rs!usuariocodigo & rss!puntovtacodigo & "'")
         If rs2.RecordCount = 0 Then
            ListView1.ListItems.item(i + 0).Checked = 0
         Else
            ListView1.ListItems.item(i + 0).Checked = 1
         End If
       Else
         ListView1.ListItems.item(i + 0).Checked = 1
       End If
         i = i + 1
      rss.MoveNext
   Loop
  End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim rs1 As New ADODB.Recordset
SQL = "select a.puntovtacodigo,a.puntovtadescripcion,valor=0 from vt_puntoventa a"
Set rs1 = VGCNx.Execute(SQL)
   Call LlenarLista(rs1, 0)
End Sub

Private Sub grabar()
Dim rs1 As New ADODB.Recordset
SQL = "select a.puntovtacodigo,a.puntovtadescripcion,valor=0 from vt_puntoventa a"
Set rs1 = VGCNx.Execute(SQL)
Dim i As Integer
Dim rs2 As New ADODB.Recordset
i = 1
Do While Not rs1.EOF
   Set rs2 = VGCNx.Execute(" select * from vt_usuarioxPuntoVta where usuariocodigo+ puntovtacodigo='" & rs!usuariocodigo & rs1!puntovtacodigo & "'")
   If ListView1.ListItems.item(i + 0).Checked = 0 Then
      If rs2.RecordCount > 0 Then
         SQL = "delete vt_usuarioxPuntoVta where usuariocodigo+puntovtacodigo="
         SQL = SQL & "'" & rs!usuariocodigo & rs1!puntovtacodigo & "'"
         Set rs2 = VGCNx.Execute(SQL)
      End If
    Else
      If rs2.RecordCount = 0 Then
         SQL = "Insert vt_usuarioxPuntoVta ( usuariocodigo,  puntovtacodigo)"
         SQL = SQL & "values('" & rs!usuariocodigo & "','" & rs1!puntovtacodigo & "')"
         Set rs2 = VGCNx.Execute(SQL)
      End If
    End If
    i = i + 1
    rs1.MoveNext
   Loop
SQL = "select a.puntovtacodigo,a.puntovtadescripcion,valor=0 from vt_puntoventa a"
Set rs1 = VGCNx.Execute(SQL)
Call LlenarLista(rs1, 0)

  End Sub
  
