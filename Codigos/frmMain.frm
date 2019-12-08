VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Editor de Particulas ORE "
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   11490
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOpenStreamFile 
      Caption         =   "&Open Stream File"
      Height          =   255
      Left            =   240
      TabIndex        =   99
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lista de particulas"
      Height          =   4455
      Left            =   0
      TabIndex        =   97
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List2 
         BackColor       =   &H00C0C0C0&
         Height          =   4155
         ItemData        =   "frmMain.frx":000C
         Left            =   45
         List            =   "frmMain.frx":000E
         TabIndex        =   98
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Visor de grh"
      Height          =   2775
      Left            =   7920
      TabIndex        =   96
      Top             =   6600
      Width           =   3495
      Begin VB.PictureBox invpic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   3255
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frameGrhs 
      Caption         =   "Grh Parameters"
      Height          =   6075
      Left            =   9480
      TabIndex        =   88
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   600
         TabIndex        =   93
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   0
         TabIndex        =   92
         Top             =   3840
         Width           =   615
      End
      Begin VB.ListBox lstSelGrhs 
         BackColor       =   &H00C0C0C0&
         Height          =   1620
         Left            =   120
         TabIndex        =   91
         Top             =   4320
         Width           =   1770
      End
      Begin VB.ListBox lstGrhs 
         BackColor       =   &H00C0C0C0&
         Height          =   3180
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   1740
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   1200
         TabIndex        =   89
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grhs Seleccionados"
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Grh"
         Height          =   195
         Left            =   60
         TabIndex        =   94
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame frameGravity 
      BorderStyle     =   0  'None
      Caption         =   "Gravity Settings"
      Height          =   1095
      Left            =   330
      TabIndex        =   81
      Top             =   7110
      Width           =   1935
      Begin VB.TextBox txtGravStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   84
         Text            =   "5"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtBounceStrength 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   83
         Text            =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkGravity 
         Caption         =   "Gravity Influence"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bounce Strength:"
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   705
         Width           =   1245
      End
   End
   Begin VB.Frame frameMovement 
      BorderStyle     =   0  'None
      Caption         =   "Movement Settings"
      Height          =   1935
      Left            =   315
      TabIndex        =   70
      Top             =   7095
      Width           =   1935
      Begin VB.TextBox move_x1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   76
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox move_x2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   75
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox move_y1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   74
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox move_y2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   73
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox chkYMove 
         Caption         =   "Y Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkXMove 
         Caption         =   "X Movement"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   525
         Width           =   1035
      End
   End
   Begin VB.Frame frameSpinSettings 
      BorderStyle     =   0  'None
      Caption         =   "Spin Settings"
      Height          =   1095
      Left            =   345
      TabIndex        =   64
      Top             =   7095
      Width           =   1935
      Begin VB.TextBox spin_speedL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   67
         Text            =   "1"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox spin_speedH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   66
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkSpin 
         Caption         =   "Spin"
         Height          =   255
         Left            =   105
         TabIndex        =   65
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (L):"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Speed (H):"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Particle Duration"
      Height          =   855
      Left            =   330
      TabIndex        =   60
      Top             =   7125
      Width           =   1935
      Begin VB.CheckBox chkNeverDies 
         Caption         =   "Never Dies"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox life 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   61
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Life:"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame frmSettings 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   960
      TabIndex        =   27
      Top             =   7080
      Width           =   6600
      Begin VB.TextBox txRad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         TabIndex        =   102
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtry 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "0"
         Top             =   1635
         Width           =   495
      End
      Begin VB.CheckBox chkresize 
         Caption         =   "Resize"
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   1920
         Width           =   1245
      End
      Begin VB.CheckBox chkAlphaBlend 
         Caption         =   "Alpha Blend"
         Height          =   255
         Left            =   3930
         TabIndex        =   42
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox fric 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "5"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox life2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "50"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox life1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox vecy2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox vecy1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "-50"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox vecx2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox vecx1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "-10"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "0"
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox txtY2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtY1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtX2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtX1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   30
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtrx 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3150
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "0"
         Top             =   1395
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Radio:"
         Height          =   255
         Left            =   3915
         TabIndex        =   103
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize X:"
         Height          =   195
         Left            =   1950
         TabIndex        =   59
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Y:"
         Height          =   195
         Left            =   1950
         TabIndex        =   58
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X2:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X1:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   525
         Width           =   240
      End
      Begin VB.Label lblPCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of Particles:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y1:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y2:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Friction:"
         Height          =   195
         Left            =   3915
         TabIndex        =   52
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (H):"
         Height          =   195
         Left            =   3915
         TabIndex        =   51
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Range (L):"
         Height          =   195
         Left            =   3915
         TabIndex        =   50
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y2"
         Height          =   195
         Left            =   1950
         TabIndex        =   49
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector Y1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   48
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X2:"
         Height          =   195
         Left            =   1950
         TabIndex        =   47
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vector X1:"
         Height          =   195
         Left            =   1950
         TabIndex        =   46
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1650
         Width           =   450
      End
   End
   Begin VB.Frame frameColorSettings 
      BorderStyle     =   0  'None
      Caption         =   "Color Tint Settings"
      Height          =   2175
      Left            =   375
      TabIndex        =   15
      Top             =   7035
      Width           =   3975
      Begin VB.HScrollBar RScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   23
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar GScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   22
         Top             =   1500
         Width           =   3015
      End
      Begin VB.HScrollBar BScroll 
         Height          =   255
         Left            =   360
         Max             =   255
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ListBox lstColorSets 
         Height          =   840
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         ScaleHeight     =   795
         ScaleWidth      =   2355
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtG 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "0"
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox txtB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1500
         Width           =   165
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   150
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Particle Speed"
      Height          =   855
      Left            =   435
      TabIndex        =   12
      Top             =   7170
      Width           =   1935
      Begin VB.TextBox speed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "0.5"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Render Delay:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame frmfade 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2235
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   7680
      Begin VB.TextBox txtfout 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Text            =   "0"
         Top             =   405
         Width           =   645
      End
      Begin VB.TextBox txtfin 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "0"
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label36 
         Caption         =   "Note: The time a particle remains alive is set in the Duration Tab"
         Height          =   585
         Left            =   90
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label37 
         Caption         =   "Fade out time"
         Height          =   300
         Left            =   60
         TabIndex        =   10
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "Fade in time"
         Height          =   180
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   2400
      ScaleHeight     =   435
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   5
      Top             =   0
      Width           =   7035
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Desaparecer"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nueva Particula"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar Particula"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Vista Previa"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2670
      Left            =   0
      TabIndex        =   87
      Top             =   6720
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   4710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuracion de Particula"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravedad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimiento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Vueltas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Velocidad"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Duracion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Color "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fade"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Adaptado y Traducido por Lorwik www.RincondelAO.com.ar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   101
      Top             =   9360
      Width           =   11415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      Height          =   375
      Left            =   9480
      TabIndex        =   100
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnunueva 
         Caption         =   "&Nueva particula"
      End
      Begin VB.Menu mnuabrir 
         Caption         =   "&Abrir archivo.."
      End
      Begin VB.Menu mnuguardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Mas 
      Caption         =   "Más"
      Begin VB.Menu sobre 
         Caption         =   "Sobre..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'--> Current Stream File <--
Public CurStreamFile As String

Private Sub Command4_Click()
    
    Dim StreamFile As String
    Dim Bypass     As Boolean
    Dim RetVal     As Byte

    CurStreamFile = App.Path & "\INIT\Particles.ini"

    If FileExists(CurStreamFile, vbNormal) = True Then
        
        RetVal = MsgBox("¡El archivo " & CurStreamFile & " ya existe!" & vbCrLf & "¿Deseas sobreescribirlo?", vbYesNoCancel Or vbQuestion)
        
        Select Case RetVal
        
            Case vbNo
                Bypass = False
            
            Case vbCancel
                Exit Sub
                
            Case vbYes
                StreamFile = CurStreamFile
                Bypass = True
                
        End Select

    End If

    If Bypass = False Then

        With ComDlg
            .Filter = "*.ini (Stream Data Files)|*.ini"
            .ShowSave
            StreamFile = .FileName
        End With
    
        If FileExists(StreamFile, vbNormal) = True Then
            
            RetVal = MsgBox("¡El archivo " & StreamFile & " ya existe!" & vbCrLf & "¿Desea sobreescribirlo?", vbYesNo Or vbQuestion)

            If RetVal = vbNo Then Exit Sub

        End If

    End If

    Dim GrhListing As String
    Dim i          As Long
    Dim j          As Long
    Dim LoopC      As Long

    'Check for existing data file and kill it
    If FileExists(StreamFile, vbNormal) Then Kill StreamFile

    'Write particle data to Particles.ini
    Call WriteVar(StreamFile, "INIT", "Total", Val(TotalStreams))

    For LoopC = 1 To TotalStreams
        Call WriteVar(StreamFile, Val(LoopC), "Name", StreamData(LoopC).Name)
        Call WriteVar(StreamFile, Val(LoopC), "NumOfParticles", Val(StreamData(LoopC).NumOfParticles))
        Call WriteVar(StreamFile, Val(LoopC), "X1", Val(StreamData(LoopC).x1))
        Call WriteVar(StreamFile, Val(LoopC), "Y1", Val(StreamData(LoopC).y1))
        Call WriteVar(StreamFile, Val(LoopC), "X2", Val(StreamData(LoopC).x2))
        Call WriteVar(StreamFile, Val(LoopC), "Y2", Val(StreamData(LoopC).y2))
        Call WriteVar(StreamFile, Val(LoopC), "Angle", Val(StreamData(LoopC).angle))
        Call WriteVar(StreamFile, Val(LoopC), "VecX1", Val(StreamData(LoopC).vecx1))
        Call WriteVar(StreamFile, Val(LoopC), "VecX2", Val(StreamData(LoopC).vecx2))
        Call WriteVar(StreamFile, Val(LoopC), "VecY1", Val(StreamData(LoopC).vecy1))
        Call WriteVar(StreamFile, Val(LoopC), "VecY2", Val(StreamData(LoopC).vecy2))
        Call WriteVar(StreamFile, Val(LoopC), "Life1", Val(StreamData(LoopC).life1))
        Call WriteVar(StreamFile, Val(LoopC), "Life2", Val(StreamData(LoopC).life2))
        Call WriteVar(StreamFile, Val(LoopC), "Friction", Val(StreamData(LoopC).friction))
        Call WriteVar(StreamFile, Val(LoopC), "Spin", Val(StreamData(LoopC).spin))
        Call WriteVar(StreamFile, Val(LoopC), "Spin_SpeedL", Val(StreamData(LoopC).spin_speedL))
        Call WriteVar(StreamFile, Val(LoopC), "Spin_SpeedH", Val(StreamData(LoopC).spin_speedH))
        Call WriteVar(StreamFile, Val(LoopC), "Grav_Strength", Val(StreamData(LoopC).grav_strength))
        Call WriteVar(StreamFile, Val(LoopC), "Bounce_Strength", Val(StreamData(LoopC).bounce_strength))
    
        Call WriteVar(StreamFile, Val(LoopC), "AlphaBlend", Val(StreamData(LoopC).AlphaBlend))
        Call WriteVar(StreamFile, Val(LoopC), "Gravity", Val(StreamData(LoopC).gravity))
    
        Call WriteVar(StreamFile, Val(LoopC), "XMove", Val(StreamData(LoopC).XMove))
        Call WriteVar(StreamFile, Val(LoopC), "YMove", Val(StreamData(LoopC).YMove))
        Call WriteVar(StreamFile, Val(LoopC), "move_x1", Val(StreamData(LoopC).move_x1))
        Call WriteVar(StreamFile, Val(LoopC), "move_x2", Val(StreamData(LoopC).move_x2))
        Call WriteVar(StreamFile, Val(LoopC), "move_y1", Val(StreamData(LoopC).move_y1))
        Call WriteVar(StreamFile, Val(LoopC), "move_y2", Val(StreamData(LoopC).move_y2))
        Call WriteVar(StreamFile, Val(LoopC), "Radio", Val(StreamData(LoopC).Radio))
        Call WriteVar(StreamFile, Val(LoopC), "life_counter", Val(StreamData(LoopC).life_counter))
        Call WriteVar(StreamFile, Val(LoopC), "Speed", Str(StreamData(LoopC).speed))
    
        Call WriteVar(StreamFile, Val(LoopC), "resize", CInt(StreamData(LoopC).grh_resize))
        Call WriteVar(StreamFile, Val(LoopC), "rx", StreamData(LoopC).grh_resizex)
        Call WriteVar(StreamFile, Val(LoopC), "ry", StreamData(LoopC).grh_resizey)
    
        Call WriteVar(StreamFile, Val(LoopC), "NumGrhs", Val(StreamData(LoopC).NumGrhs))
    
        GrhListing = vbNullString

        For i = 1 To StreamData(LoopC).NumGrhs
            GrhListing = GrhListing & StreamData(LoopC).grh_list(i) & ","
        Next i
    
        Call WriteVar(StreamFile, Val(LoopC), "Grh_List", GrhListing)
    
        Call WriteVar(StreamFile, Val(LoopC), "ColorSet1", StreamData(LoopC).colortint(0).r & "," & StreamData(LoopC).colortint(0).g & "," & StreamData(LoopC).colortint(0).B)
        Call WriteVar(StreamFile, Val(LoopC), "ColorSet2", StreamData(LoopC).colortint(1).r & "," & StreamData(LoopC).colortint(1).g & "," & StreamData(LoopC).colortint(1).B)
        Call WriteVar(StreamFile, Val(LoopC), "ColorSet3", StreamData(LoopC).colortint(2).r & "," & StreamData(LoopC).colortint(2).g & "," & StreamData(LoopC).colortint(2).B)
        Call WriteVar(StreamFile, Val(LoopC), "ColorSet4", StreamData(LoopC).colortint(3).r & "," & StreamData(LoopC).colortint(3).g & "," & StreamData(LoopC).colortint(3).B)
    
    Next LoopC
        
    'Report the results
    If TotalStreams > 1 Then
        Call MsgBox(TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation)
    Else
        Call MsgBox(TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation)

    End If
    
    'Set DataChanged variable to false
    DataChanged = False
    CurStreamFile = StreamFile

End Sub

Private Sub Command5_Click()

    Dim Nombre          As String
    Dim NewStreamNumber As Integer
    Dim grhlist(0)      As Long

    'Get name for new stream
    Nombre = InputBox("Por favor inserte un nombre a la particula", "New Stream")

    If LenB(Nombre) = 0 Then Exit Sub

    'Set new stream
    NewStreamNumber = List2.ListCount + 1

    'Add stream to combo box
    Call List2.AddItem(Nombre)

    'Add 1 to TotalStreams
    TotalStreams = TotalStreams + 1

    grhlist(0) = 19751
    
    'Add stream data to StreamData array
    With StreamData(NewStreamNumber)
        .Name = Nombre
        .NumOfParticles = 20
        .x1 = 0
        .y1 = 0
        .x2 = 0
        .y2 = 0
        .angle = 0
        .vecx1 = -20
        .vecx2 = 20
        .vecy1 = -20
        .vecy2 = 20
        .life1 = 10
        .life2 = 50
        .friction = 8
        .spin_speedL = 0.1
        .spin_speedH = 0.1
        .grav_strength = 2
        .bounce_strength = -5
        .speed = 0.5
        .AlphaBlend = 1
        .gravity = 0
        .XMove = 0
        .YMove = 0
        .move_x1 = 0
        .move_x2 = 0
        .move_y1 = 0
        .move_y2 = 0
        .life_counter = -1
        .NumGrhs = 1
        .grh_list = grhlist()
    End With
    

    'Select the new stream type in the combo box
    List2.ListIndex = NewStreamNumber - 1

End Sub

Private Sub Command6_Click()

    If List2.ListIndex < 0 Then Exit Sub
    Call CargarParticulasLista

End Sub

Private Sub Command8_Click()
    Particle_Group_Remove_All

End Sub

Private Sub Form_Load()

    With lstColorSets
        Call .AddItem("Bottom Left")
        Call .AddItem("Top Left")
        Call .AddItem("Bottom Right")
        Call .AddItem("Top Right")
    End With
    
    frmSettings.Visible = True
    frmfade.Visible = False
    frameColorSettings.Visible = False
    Frame2.Visible = False
    Frame1.Visible = False
    frameSpinSettings.Visible = False
    frameMovement.Visible = False
    frameGravity.Visible = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    HookSurfaceHwnd Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label2_Click()
    End
End Sub

Private Sub List2_Click()
    Call CargarParticulasLista
End Sub

Sub CargarParticulasLista()

    Dim LoopC    As Long
    Dim DataTemp As Boolean

    DataTemp = DataChanged

    'Set the values
    With StreamData(List2.ListIndex + 1)
        txtPCount.Text = .NumOfParticles
        txtX1.Text = .x1
        txtY1.Text = .y1
        txtX2.Text = .x2
        txtY2.Text = .y2
        txtAngle.Text = .angle
        vecx1.Text = .vecx1
        vecx2.Text = .vecx2
        vecy1.Text = .vecy1
        vecy2.Text = .vecy2
        life1.Text = .life1
        life2.Text = .life2
        fric.Text = .friction
        chkSpin.Value = .spin
        spin_speedL.Text = .spin_speedL
        spin_speedH.Text = .spin_speedH
        txtGravStrength.Text = .grav_strength
        txtBounceStrength.Text = .bounce_strength
        chkAlphaBlend.Value = .AlphaBlend
        chkGravity.Value = .gravity
        txtrx.Text = .grh_resizex
        txtry.Text = .grh_resizey
        chkXMove.Value = .XMove
        chkYMove.Value = .YMove
        move_x1.Text = .move_x1
        move_x2.Text = .move_x2
        move_y1.Text = .move_y1
        move_y2.Text = .move_y2
        txRad.Text = .Radio

        If .grh_resize = True Then
            chkresize = vbChecked
        Else
            chkresize = vbUnchecked
        End If

        If .life_counter = -1 Then
            life.Enabled = False
            chkNeverDies.Value = vbChecked
        Else
            life.Enabled = True
            life.Text = .life_counter
            chkNeverDies.Value = vbUnchecked
        End If

        speed.Text = .speed

        lstSelGrhs.Clear

        For LoopC = 1 To .NumGrhs
            Call lstSelGrhs.AddItem(.grh_list(LoopC))
        Next LoopC
    
    End With

    DataChanged = DataTemp

    indexs = frmMain.List2.ListIndex + 1

    Call General_Particle_Create(indexs, 50, 50)

End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim LoopC    As Long
    Dim DataTemp As Boolean

    DataTemp = DataChanged

    'Set the values
    With StreamData(List2.ListIndex + 1)
    
        txtPCount.Text = .NumOfParticles
        txtX1.Text = .x1
        txtY1.Text = .y1
        txtX2.Text = .x2
        txtY2.Text = .y2
        txtAngle.Text = .angle
        vecx1.Text = .vecx1
        vecx2.Text = .vecx2
        vecy1.Text = .vecy1
        vecy2.Text = .vecy2
        life1.Text = .life1
        life2.Text = .life2
        fric.Text = .friction
        chkSpin.Value = .spin
        spin_speedL.Text = .spin_speedL
        spin_speedH.Text = .spin_speedH
        txtGravStrength.Text = .grav_strength
        txtBounceStrength.Text = .bounce_strength

        chkAlphaBlend.Value = .AlphaBlend
        chkGravity.Value = .gravity

        chkXMove.Value = .XMove
        chkYMove.Value = .YMove
        move_x1.Text = .move_x1
        move_x2.Text = .move_x2
        move_y1.Text = .move_y1
        move_y2.Text = .move_y2
        txRad.Text = .Radio

        lstSelGrhs.Clear

        For LoopC = 1 To .NumGrhs
            Call lstSelGrhs.AddItem(.grh_list(LoopC))
        Next LoopC

    End With

End Sub

Private Sub lstGrhs_Click()
On Error Resume Next

    Call invpic.Cls
    Call GrhRenderToHdc(lstGrhs.List(lstGrhs.ListIndex), invpic.hdc, 2, 2, True)

End Sub

Private Sub mnuabrir_Click()
    Call cmdOpenStreamFile_Click
End Sub

Private Sub cmdOpenStreamFile_Click()

    Dim sFile As String

    With ComDlg
        .Filter = "*.ini (Stream Data Files)|*.ini"
        .ShowOpen
        sFile = .FileName
    End With
    
    If LenB(sFile) Then
        Call LoadStreamFile(sFile)
        CurStreamFile = sFile
    End If
    
End Sub

Private Sub mnuguardar_Click()
    Call Command4_Click
End Sub

Private Sub mnunueva_Click()
    Call Command5_Click
End Sub

Private Sub mnusalir_Click()
    End
End Sub

Private Sub mnusobre_Click()
    frmCreditos.Show vbModal, Me

End Sub

Private Sub sobre_Click()
    frmCreditos.Show
End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.index

        Case 1:
            frmSettings.Visible = True
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 2:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = True
            frmfade.Visible = False

        Case 3:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = True
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 4:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = True
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 5:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = True
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 6:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = True
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 7:
            frmSettings.Visible = False
            frameColorSettings.Visible = True
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = False

        Case 8:
            frmSettings.Visible = False
            frameColorSettings.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
            frameSpinSettings.Visible = False
            frameMovement.Visible = False
            frameGravity.Visible = False
            frmfade.Visible = True

    End Select

End Sub

Private Sub txRad_Change()

    On Error Resume Next

    StreamData(frmMain.List2.ListIndex + 1).Radio = Val(txRad.Text)

End Sub

Private Sub txtrx_Change()

    On Error Resume Next

    StreamData(frmMain.List2.ListIndex + 1).grh_resizex = txtrx.Text

End Sub

Private Sub txtry_Change()

    On Error Resume Next

    StreamData(frmMain.List2.ListIndex + 1).grh_resizey = txtry.Text

End Sub

Private Sub vecx1_GotFocus()
    
    With vecx1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub vecx1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).vecx1 = vecx1.Text

End Sub

Private Sub vecx2_GotFocus()
    
    With vecx2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub vecx2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).vecx2 = vecx2.Text

End Sub

Private Sub vecy1_GotFocus()
    
    With vecy1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub vecy1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).vecy1 = vecy1.Text

End Sub

Private Sub vecy2_GotFocus()
    
    With vecy2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub vecy2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).vecy2 = vecy2.Text

End Sub

Private Sub life1_GotFocus()
    
    With life1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub life1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).life1 = life1.Text

End Sub

Private Sub life2_GotFocus()

    With life2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub life2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).life2 = life2.Text

End Sub

Private Sub fric_GotFocus()
    
    With fric
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub fric_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).friction = fric.Text

End Sub

Private Sub spin_speedL_GotFocus()
    
    With spin_speedL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub spin_speedL_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).spin_speedL = spin_speedL.Text

End Sub

Private Sub spin_speedH_GotFocus()
    
    With spin_speedH
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub spin_speedH_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).spin_speedH = spin_speedH.Text

End Sub

Private Sub txtPCount_GotFocus()
    
    With txtPCount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtPCount_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).NumOfParticles = txtPCount.Text

End Sub

Private Sub txtX1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).x1 = txtX1.Text

End Sub

Private Sub txtX1_GotFocus()
    
    With txtX1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtY1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).y1 = txtY1.Text

End Sub

Private Sub txtY1_GotFocus()

    With txtY1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtX2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).x2 = txtX2.Text

End Sub

Private Sub txtX2_GotFocus()

    txtX2.SelStart = 0
    txtX2.SelLength = Len(txtX2.Text)

End Sub

Private Sub txtY2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).y2 = txtY2.Text

End Sub

Private Sub txtY2_GotFocus()

    With txtY2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtAngle_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).angle = txtAngle.Text

End Sub

Private Sub txtAngle_GotFocus()

    With txtAngle
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtGravStrength_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).grav_strength = txtGravStrength.Text

End Sub

Private Sub txtGravStrength_GotFocus()
    
    With txtGravStrength
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtBounceStrength_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).bounce_strength = txtBounceStrength.Text

End Sub

Private Sub txtBounceStrength_GotFocus()
    
    With txtBounceStrength
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub move_x1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).move_x1 = move_x1.Text

End Sub

Private Sub move_x1_GotFocus()
    
    With move_x1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub move_x2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).move_x2 = move_x2.Text

End Sub

Private Sub move_x2_GotFocus()
    
    With move_x2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub move_y1_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).move_y1 = move_y1.Text

End Sub

Private Sub move_y1_GotFocus()
    
    With move_y1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub move_y2_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).move_y2 = move_y2.Text

End Sub

Private Sub move_y2_GotFocus()

    With move_y2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub chkAlphaBlend_Click()

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).AlphaBlend = chkAlphaBlend.Value

End Sub

Private Sub chkGravity_Click()

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).gravity = chkGravity.Value

    If chkGravity.Value = vbChecked Then
        txtGravStrength.Enabled = True
        txtBounceStrength.Enabled = True
    Else
        txtGravStrength.Enabled = False
        txtBounceStrength.Enabled = False
    End If

End Sub

Private Sub chkXMove_Click()

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).XMove = chkXMove.Value

    If chkXMove.Value = vbChecked Then
        move_x1.Enabled = True
        move_x2.Enabled = True
    Else
        move_x1.Enabled = False
        move_x2.Enabled = False
    End If

End Sub

Private Sub chkYMove_Click()

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).YMove = chkYMove.Value

    If chkYMove.Value = vbChecked Then
        move_y1.Enabled = True
        move_y2.Enabled = True
    Else
        move_y1.Enabled = False
        move_y2.Enabled = False
    End If

End Sub

Private Sub BScroll_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).B = BScroll.Value
    txtB.Text = BScroll.Value

    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub chkNeverDies_Click()

    DataChanged = True
    
    With StreamData(frmMain.List2.ListIndex + 1)

        If chkNeverDies.Value = vbChecked Then
            life.Enabled = False
            .life_counter = -1
        Else
            life.Enabled = True
            .life_counter = life.Text
        End If
    
    End With

End Sub

Private Sub chkSpin_Click()

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).spin = chkSpin.Value

    If chkSpin.Value = vbChecked Then
        spin_speedL.Enabled = True
        spin_speedH.Enabled = True
    Else
        spin_speedL.Enabled = False
        spin_speedH.Enabled = False
    End If

End Sub

Private Sub GScroll_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).g = GScroll.Value
    txtG.Text = GScroll.Value

    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub life_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).life_counter = life.Text

End Sub

Private Sub life_GotFocus()

    With life
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub lstColorSets_Click()

    Dim DataTemp As Boolean
        DataTemp = DataChanged
    
    With StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex)
        RScroll.Value = .r
        GScroll.Value = .g
        BScroll.Value = .B
    End With

    DataChanged = DataTemp

End Sub

Private Sub RScroll_Change()

    On Error Resume Next

    DataChanged = True

    StreamData(frmMain.List2.ListIndex + 1).colortint(lstColorSets.ListIndex).r = RScroll.Value
    
    txtR.Text = RScroll.Value

    picColor.BackColor = RGB(txtB.Text, txtG.Text, txtR.Text)

End Sub

Private Sub speed_Change()

    On Error Resume Next

    DataChanged = True

    'Arrange decimal separator
    Dim temp As String
        temp = ReadField(1, speed.Text, 44)

    If LenB(temp) Then
    
        With speed
            .Text = temp & "." & Right$(.Text, Len(.Text) - Len(temp) - 1)
            .SelStart = Len(.Text)
            .SelLength = 0
        End With
        
    End If

    StreamData(frmMain.List2.ListIndex + 1).speed = Val(speed.Text)

End Sub

Private Sub speed_GotFocus()

    speed.SelStart = 0
    speed.SelLength = Len(speed.Text)

End Sub

Private Sub lstSelGrhs_DblClick()

    Call cmdDelete_Click

End Sub

Private Sub cmdDelete_Click()
    
    ' [PARCHE - By Jopi] - Al clickear en un Grh animado tira Error 9.
    If InStrB(1, "(animacion)", "(animacion)", vbTextCompare) Then Exit Sub
    
    Dim LoopC As Long

    If lstSelGrhs.ListIndex >= 0 Then
        Call lstSelGrhs.RemoveItem(lstSelGrhs.ListIndex)
    End If
    
    With StreamData(List2.ListIndex + 1)
    
        .NumGrhs = lstSelGrhs.ListCount

        If .NumGrhs = 0 Then
            Erase .grh_list
        Else
            ReDim .grh_list(1 To lstSelGrhs.ListCount)
        End If

        For LoopC = 1 To .NumGrhs
            .grh_list(LoopC) = lstSelGrhs.List(LoopC - 1)
        Next LoopC
    
    End With

End Sub

Private Sub lstSelGrhs_Click()

    Dim framecount As Long
    If framecount <= 0 Then Exit Sub

    Call invpic.Cls
    Call GrhRenderToHdc(lstSelGrhs.List(lstSelGrhs.ListIndex), invpic.hdc, 2, 2, True)

End Sub

Private Sub lstGrhs_DblClick()

    Call cmdAdd_Click

End Sub

Private Sub cmdAdd_Click()
    
    ' [PARCHE - By Jopi] - Al clickear en un Grh animado tira Error 9.
    If InStrB(1, "(animacion)", "(animacion)", vbTextCompare) Then Exit Sub
    
    Dim LoopC As Long

    If lstGrhs.ListIndex >= 0 Then
        Call lstSelGrhs.AddItem(lstGrhs.List(lstGrhs.ListIndex))

    End If
    
    With StreamData(List2.ListIndex + 1)
    
        .NumGrhs = lstSelGrhs.ListCount

        ReDim .grh_list(1 To lstSelGrhs.ListCount)

        For LoopC = 1 To .NumGrhs
            .grh_list(LoopC) = lstSelGrhs.List(LoopC - 1)
        Next LoopC
    
    End With
    
End Sub

Private Sub chkresize_Click()
    
    With StreamData(frmMain.List2.ListIndex + 1)
    
        If chkresize.Value = vbChecked Then
            .grh_resize = True
        Else
            .grh_resize = False
        End If

    End With

End Sub
