Attribute VB_Name = "Modulo"
Option Explicit

Public VGdllApi As dll_apisgen.dll_apis
Public VGReclamos As String
Public VGComercializacion As String
Public VGMedicion As String
Public VGprincipales As String
Public VGInspectoria As String
Public VGfacturacion As String
Public VGrecaudacion As String


Public Sub Main()
    'Base de Datos General
      
On Error GoTo Xmain
    'VGusuario = "03"
    'Leer Ini
    Set VGdllApi = New dll_apisgen.dll_apis
    
    
    VGReclamos = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Reclamos", "?")
    VGComercializacion = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Comercializacion", "?")
    VGMedicion = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "Medicion", "?")
    VGprincipales = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "principales", "?")
    VGInspectoria = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "inspectoria", "?")
    VGfacturacion = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "facturacion", "?")
    VGrecaudacion = VGdllApi.LeerIni(App.Path & "\integra.ini", "E01", "recaudacion", "?")
        

    FrmMain.Show
    Exit Sub
Xmain:
    MsgBox Err.Description, vbExclamation, "Error Main"
End Sub

