Attribute VB_Name = "modDX8Requires"
Option Explicit

Public vertList(3) As TLVERTEX

Public Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public DirectX        As DirectX8
Public D3D       As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX      As D3DX8

Public Type TLVERTEX
    x As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Const PI   As Single = 3.14159265358979

'To get free bytes in RAM
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function General_Get_Free_Ram_Bytes() As Long
    
    Call GlobalMemoryStatus(pUdtMemStatus)
    
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
    
End Function
