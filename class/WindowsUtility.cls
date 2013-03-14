VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowsUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
Private Const CSIDL_NETWORK As Long = &H11
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
' Const FLASHW_STOP as Long = 0 'deja de parpadear
Public Enum oppar
    FLASHW_CAPTION = 1
    FLASHW_TRAY = 2
    FLASHW_TIMER = 4 'parpadea hasta que le mandemos parar
    FLASHW_TIMERNOFG = 12
End Enum
'Public Const FLASHW_CAPTION = 1 'parpadea el título
'Public Const FLASHW_TRAY = 2 'parpadea el botón de la barra de tareas
'Public Const FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY 'parpadean los dos
'Public Const FLASHW_TIMER = 4 'parpadea hasta que le mandemos parar
'Public Const FLASHW_TIMERNOFG = 12 'parpadea hasta que se active la ventana
'
Private Type FLASHWINFO
    cbSize As Long 'tamaño de la estructura
    hwnd As Long 'handle de la ventana que debe parpadear
    dwFlags As Long 'opciones
    uCount As Long 'número de parpadeos
    dwTimeOut As Long 'tiempo entre los parpadeos en milisegundos, si es cero será el del parpadeo del cursor
End Type

Private Declare Function FlashWindowEx Lib "user32" (pFWI As FLASHWINFO) As Boolean

Private Type BROWSEINFO 'BI
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

'Emplearemos esta función auxiliar para tratar las cadenas devueltas por el API :
Private Function AsciiZ(texto As String) As String
'esta función se utiliza para devolver la parte del string
'hasta el primer chr(0). Las funciones del API suelen emplear
'un chr(0) para marcar el final de las cadenas de texto
Dim I As Long
I = InStr(texto, Chr(0))
If I = 0 Then
    AsciiZ = texto
Else
    AsciiZ = Left(texto, I - 1)
End If
End Function

'Y luego esta función nos devolverá un nombre de máquina o de recurso compartido en notación UNC (\\NombreOrdenador\NombreRecurso) :
Public Function BrowseForShares(hwnd As Long, msg As String) As String
  'mostrar sólo ordenadores y recursos compartidos
   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer

  'obtener el pidl de la carpeta 'Entorno de red'
   If SHGetSpecialFolderLocation(hwnd, CSIDL_NETWORK, pidl) = ERROR_SUCCESS Then
     'rellenar la estructura limitando a ordenadores especificando
     'el pidl devuelto como pidl raíz
      With BI
         .hOwner = hwnd
         .pidlRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = msg
         .ulFlags = BIF_RETURNONLYFSDIRS
      End With
     'mostrar el diálogo de búsqueda
      pidl = SHBrowseForFolder(BI)
      If pidl <> 0 Then
        'comprobar que es un pidl válido
         sPath = Space$(MAX_PATH)
         If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
           'es válido, se trata de un recurso compartido
            BrowseForShares = AsciiZ(sPath)
         Else
            'se ha seleccionado un ordenador
             BrowseForShares = "\\" & AsciiZ(BI.pszDisplayName)
         End If
      End If
   End If
End Function
Public Function ParpadeaWindow(Size As Long, hwnd As Long, flags As oppar, count As Long, timeout As Long)
Dim para As FLASHWINFO
    With para
        .cbSize = Size
        .dwFlags = flags
        .dwTimeOut = timeout
        .hwnd = hwnd
        .uCount = count
    End With
    ParpadeaWindow = FlashWindowEx(para)
End Function

