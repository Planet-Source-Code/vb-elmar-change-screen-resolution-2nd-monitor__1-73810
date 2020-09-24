VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Change the screen resolution"
      Height          =   735
      Left            =   4560
      TabIndex        =   62
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C000&
      Caption         =   "Screen Rotation Listing : "
      Height          =   1455
      Index           =   26
      Left            =   7920
      TabIndex        =   56
      Top             =   4440
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "0° only"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   225
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "90° only"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   465
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "180° only"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   705
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "270° only"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   58
         Top             =   945
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   " list  all  rotations"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   1185
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " List1 "
      Height          =   1095
      Left            =   0
      TabIndex        =   54
      Top             =   120
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " List2 "
      Height          =   2055
      Left            =   0
      TabIndex        =   52
      Top             =   1560
      Width           =   4455
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDisplayFrequency"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   6240
      TabIndex        =   25
      Top             =   4440
      Width           =   1575
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDisplayFlags"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   4560
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPelsHeight"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   9240
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPelsWidth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   7680
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmBitsPerPel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   6240
      TabIndex        =   21
      Top             =   3720
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmUnusedPadding"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   4560
      TabIndex        =   20
      Top             =   3720
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "dmFormName"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   9240
      TabIndex        =   19
      Top             =   3000
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmCollate"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   7680
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmTTOption"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmYResolution"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   4560
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDuplex"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   9240
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmColor"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   7680
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPrintQuality"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   6240
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDefaultSource"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   4560
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmCopies"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   9240
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmScale"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   7680
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPaperWidth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPaperLength"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   4560
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmPaperSize"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   9240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmOrientation"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7680
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmFields"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   6240
      TabIndex        =   5
      Top             =   840
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDriverExtra"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmSize"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmDriverVersion"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "dmSpecVersion"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "dmDeviceName"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (DeviceName As Any, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "User32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long

Private Declare Function ChangeDisplaySettingsEx Lib "user32" Alias _
                          "ChangeDisplaySettingsExA" (lpszDeviceName As Any, lpDevMode As Any, _
                           ByVal hWnd As Long, ByVal dwFlags As Long, lParam As Any) As Long

Private Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1 'store screen resolution for this user
Const CDS_TEST = &H2 'return to prev resolution when vb app is closed
Const CDS_FULLSCREEN = &H4 'change resolution
Private Const CDS_GLOBAL = &H8 'store screen resolution for all users
Const CDS_RESET = &H40000000 'change scr resolution even it's the same


Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Private Sub Command1_Click()
MsgBox "To change the monitor resolution please Double Click on ListBox 2.": Command1.Visible = False
End Sub

Private Sub List1_Click()
    Call Populate2
End Sub

Private Sub List2_Click()
Call EnumDisplaySettings(ByVal List1.Text, List2.ListIndex, DevM)
Label1(0).Caption = DevM.dmDeviceName
Label1(1).Caption = DevM.dmSpecVersion
Label1(2).Caption = DevM.dmDriverVersion
Label1(3).Caption = DevM.dmSize
Label1(4).Caption = DevM.dmDriverExtra
Label1(5).Caption = DevM.dmFields

Label1(6).Caption = DevM.dmOrientation
Label1(7).Caption = DevM.dmPaperSize
Label1(8).Caption = DevM.dmPaperLength
Label1(9).Caption = DevM.dmPaperWidth
Label1(10).Caption = DevM.dmScale
Label1(11).Caption = DevM.dmCopies

Label1(12).Caption = DevM.dmDefaultSource
Label1(13).Caption = DevM.dmPrintQuality
Label1(14).Caption = DevM.dmColor
Label1(15).Caption = DevM.dmDuplex
Label1(16).Caption = DevM.dmYResolution
Label1(17).Caption = DevM.dmTTOption

Label1(18).Caption = DevM.dmCollate
Label1(19).Caption = DevM.dmFormName
Label1(20).Caption = DevM.dmUnusedPadding
Label1(21).Caption = DevM.dmBitsPerPel
Label1(22).Caption = DevM.dmPelsWidth
Label1(23).Caption = DevM.dmPelsHeight

Label1(24).Caption = DevM.dmDisplayFlags
Label1(25).Caption = DevM.dmDisplayFrequency

End Sub

Private Sub List2_DblClick()
Dim m
m = MsgBox("Change Resolution ?", vbYesNo)
If m = vbYes Then
Call EnumDisplaySettings(ByVal List1.Text, List2.ListIndex, DevM)
Call ChangeDisplaySettingsEx(ByVal List1.Text, DevM, ByVal 0&, CDS_UPDATEREGISTRY, ByVal 0&)
End If
End Sub

Private Sub Form_Load()
    Frame2.Height = 5500: List2.Height = 5100
    Call Populate1
End Sub

Private Sub Populate1()
Dim cnt As Long, DispDev As DISPLAY_DEVICE

DispDev.cb = Len(DispDev)

Do While EnumDisplayDevices(ByVal 0&, cnt, DispDev, 0&) <> 0
If InStr(1, DispDev.DeviceName, "displayv", vbTextCompare) = 0 Then
    List1.AddItem Left$(DispDev.DeviceName, InStr(DispDev.DeviceName, Chr$(0)) - 1)
End If
    cnt = cnt + 1
Loop

 If List1.ListCount > 0 Then List1.Selected(0) = True
End Sub

Private Sub Populate2()
List2.Clear
Dim a As Boolean: Dim i&, yesterday As String
i = 0

Do
    a = EnumDisplaySettings(List1.Text, i, DevM)
    
    yesterday = i
    yesterday = yesterday & "   "
    yesterday = yesterday & DevM.dmPelsWidth
    yesterday = yesterday & "x"
    yesterday = yesterday & DevM.dmPelsHeight
    yesterday = yesterday & " @"
    yesterday = yesterday & DevM.dmDisplayFrequency
    yesterday = yesterday & "  Bits:"
    yesterday = yesterday & DevM.dmBitsPerPel
    yesterday = yesterday & "  "
    yesterday = yesterday & "dmScale"
    yesterday = yesterday & DevM.dmScale
    
    If DevM.dmScale = 1 Then yesterday = yesterday & " (=90°rotate)"
    If DevM.dmScale = 2 Then yesterday = yesterday & " (=180°rotate)"
    If DevM.dmScale = 3 Then yesterday = yesterday & " (=270°rotate)"
If Option1(0).Value And DevM.dmScale <> 0 Then yesterday = i '0° only
If Option1(1).Value And DevM.dmScale <> 1 Then yesterday = i '90° only
If Option1(2).Value And DevM.dmScale <> 2 Then yesterday = i '180° only
If Option1(3).Value And DevM.dmScale <> 3 Then yesterday = i '270° only
    
    List2.AddItem yesterday
    i = i + 1

Loop Until (a = False)
Frame2.Caption = " List2 (" & i & " items)"
End Sub

Private Sub Option1_Click(Index As Integer)
MsgBox "Now filling Listbox 2": Call Populate2
End Sub
