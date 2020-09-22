VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3ds Files Loader (andreafontana@mail.com)"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   9000
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rotations' Keys:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9000
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "C,V = yZ"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "F,G = XZ"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "R,T = XY"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Translations' Keys:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9000
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "E = Down"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   11
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Q = In"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "D = Right"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   9
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A = Left"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "S = Down"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "W = Up"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.ComboBox TxtFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6960
      Width           =   8295
   End
   Begin VB.ListBox Lst 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1800
      TabIndex        =   2
      Top             =   7440
      Width           =   7095
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H80000007&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load 3ds!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6990
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim bClick As Integer
Dim C As POINTAPI





Private Sub Command2_Click()
Pic.Refresh
End Sub

Private Sub CmdLoad_Click()
Dim Lista() As String
ReDim Lista(0)
'Pic.Cls
Lst.Clear
AzzeraVar
ReadFile TxtFile, True, Lista
TraslaX Pic.ScaleWidth / 2
Traslay Pic.ScaleHeight / 2
Dim I As Integer
For I = 1 To UBound(Lista)
Lst.AddItem Lista(I)
Next
Render tHdc
BitBlt Pic.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, tHdc, 0, 0, vbSrcCopy
Pic.SetFocus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
GetKey Pic, KeyCode
End Sub

Private Sub Form_Load()
CreaBit Pic
TxtFile.AddItem App.Path + "\pyramid.3ds"
TxtFile.AddItem App.Path + "\cube.3ds"
TxtFile.AddItem App.Path + "\toydog.3ds"
bDone = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteAll
End
End Sub

Private Sub Pic_KeyDown(KeyCode As Integer, Shift As Integer)
GetKey Pic, KeyCode
End Sub


Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bClick = True
C.X = X
C.Y = Y
End Sub



Private Sub Pic_Paint()
'Render Pic.hdc
End Sub
