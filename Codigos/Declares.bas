Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public indexs As Integer

Public font_list() As D3DXFont
Public mode        As Boolean
Public temp_rgb(3) As Long

'drag&drop cosas
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public DataChanged As Boolean

Public grh_index   As Byte

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public SurfaceDB  As clsTexManager

Public FileManager As clsIniManager

Public NumEscudosAnims            As Integer

Public ArmasHerrero(0 To 100)     As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100)    As Integer

Public CnTd                       As Byte

Public SecuenciaMacroHechizos     As Byte

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40

Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

Public Tips()                                            As String * 255

Public Actual       As Byte

'RGB Type
Public Type RGB

    r As Long
    g As Long
    B As Long

End Type

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

'Inventario
Type Inventory

    OBJIndex As Integer
    Name As String
    grhindex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer

End Type

Public Nombres            As Boolean

Public MixedKey           As Long

'User status vars
Public UserPassword       As String

Public UserName           As String

Public Musica             As Boolean

Public Sound              As Boolean

Public SkillPoints        As Integer

Public Alocados           As Integer

Public flags()            As Integer

Public Oscuridad          As Integer

Public logged             As Boolean

Public NoPuedeUsar        As Boolean

'Server stuff
Public RequestPosTimer   As Integer 'Used in main loop

Public stxtbuffer        As String 'Holds temp raw data from server

Public stxtbuffercmsg    As String 'Holds temp raw data from server

Public SendNewChar       As Boolean 'Used during login

Public Connected         As Boolean 'True when connected to server

Public DownloadingMap    As Boolean 'Currently downloading a map from server

Public UserMap           As Integer

'Control
Public prgRun            As Boolean 'When true the program ends

Public IPdelServidor     As String

Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nsize As Long, _
                                                 ByVal lpfilename As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Old fashion BitBlt function
Public Declare Function BitBlt _
               Lib "gdi32" (ByVal hDestDC As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, _
                            ByVal xSrc As Long, _
                            ByVal ySrc As Long, _
                            ByVal dwRop As Long) As Long

Public Declare Function SelectObject _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SetPixel _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long

Public Declare Function GetPixel _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal x As Long, _
                            ByVal Y As Long) As Long

