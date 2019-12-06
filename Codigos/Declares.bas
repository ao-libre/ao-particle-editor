Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public indexs As Integer

'Drag & Drop cosas
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

Public SurfaceDB  As clsTexManager

Public FileManager As clsIniManager

'RGB Type
Public Type RGB
    r As Long
    g As Long
    B As Long
End Type

'Control
Public prgRun            As Boolean 'When true the program ends

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

