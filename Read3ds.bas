Attribute VB_Name = "Read3ds"
Option Explicit
Public bDone As Boolean
Public bRunning As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Dim tPen As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim R As RECT
Dim SizeX As Long
Dim SizeY As Long
Public Const pi = 3.14159265358979
Public tHdc As Long
Public tBit As Long
Dim s3dStudioV As String
Dim sMeshV As String
Dim nPunti As Integer
Dim nFaccie As Integer
Dim xScarto As Double
Dim yScarto As Double
Public Type POINTAPI
    X As Single
    Y As Single
End Type
Private Type tVertici
    nX As Single
    nY As Single
    nZ As Single
    tu As Single
    tv As Single
    X As Single
    Y As Single
    Z As Single
End Type
Private Type tVettore
    X As Single
    Y As Single
    Z As Single
End Type
Public Type tChunk
    Header As Integer
    Length As Long
End Type
Public nFile As Integer
Dim Vertsi() As Integer
Dim nVertsi As Long
Public Type tSeg
    mtrl As Integer
    Vertsi() As Integer
    nVertsi As Long
    CullMode As Long
    Shade As Long
End Type
Public Type tSolido
    Verts() As tVertici
    nVerts As Integer
    nSegs As Integer
    Segs() As tSeg
End Type
Dim lSegs(100) As tSeg
Public nSolidi As Integer
Public Solidi() As tSolido
Dim lScale As Single
Dim nSegs As Long
'
'
'
'

Private Sub gFill(hdc As Long, StartX As Single, StartY As Single, StartZ As String, EndX As Single, Endy As Single, EndZ As Single)
  '
End Sub


Public Function ReadStr() As String
    Dim TmpChar As String
    Dim Ris As String
    TmpChar = StrConv(InputB(1, nFile), vbUnicode)
    While TmpChar <> Chr(0)
        Ris = Ris & TmpChar
        TmpChar = StrConv(InputB(1, nFile), vbUnicode)
    Wend
    ReadStr = Ris
End Function

Public Sub ReadFile(FileName As String, UseLst As Boolean, ByRef Lst() As String)
    Dim idx As Long
    Dim tch As tChunk
    Dim I As Long
    On Error GoTo err0
    nFile = FreeFile
    idx = 1
    Open FileName For Binary As nFile
    AzzeraVar
    While Not EOF(nFile)
        ReadChunk idx, UseLst, tch, Lst
    Wend
    Close nFile
    If UseLst Then
     Add2List Lst, String(100, "-")
     Add2List Lst, "Number of Points: " + Str(nPunti)
     Add2List Lst, "Number of Polys: " + Str(nFaccie)
     Add2List Lst, "Number of Solids: " + Str(nSolidi)
    End If
    If UseLst Then
    If nSolidi > 1 Then
        For I = 1 To nSolidi
             Add2List Lst, "Solid " + CStr(I) + ": " + CStr(Solidi(I).nVerts) + " Vertex."
        Next
    End If
    End If
err0:
End Sub
Public Sub ReadChunk(ByRef fPos As Long, UseLst As Boolean, ptch As tChunk, ByRef Lst() As String)
    Dim tch As tChunk
    Dim I As Long
    Dim iV1 As Integer, iV2 As Integer, iV3 As Integer, iV4 As Integer
    Dim TmpInteger As Integer
    Dim TmpLong As Long
    Dim TmpSingle As Single
    Dim TmpStr As String
    Dim lstItm As Integer
    Get #nFile, fPos, tch
    TmpStr = ""
    Select Case tch.Header
    'Versione di 3dStudio
    Case 2
        Get #nFile, , TmpLong
        TmpStr = CStr(TmpLong)
        s3dStudioV = TmpStr
        TmpStr = "3dStudio Version: " + TmpStr
        'Versione delle Mesh
    Case &H3D3E
        Get #nFile, , TmpLong
        TmpStr = CStr(TmpLong)
        sMeshV = TmpStr
        TmpStr = "Mesh Versione: " + TmpStr
    ' Scala
    Case &H100
        Get #nFile, , TmpSingle
        TmpStr = "Scale: " + CStr(TmpSingle)
        lScale = TmpSingle
    ' Oggetti
    Case &H4000
        TmpStr = ReadStr()
        TmpLong = fPos + Len(TmpStr) + 7
        TmpStr = tch.Length & "* " & TmpStr
        nSegs = 0
        While TmpLong < (fPos + tch.Length)
            ReadChunk TmpLong, UseLst, tch, Lst
        Wend
        If nSolidi > 0 Then
            If nSegs = 0 Then
                Solidi(nSolidi).nSegs = 1
                ReDim Solidi(nSolidi).Segs(1)
                ReDim Solidi(nSolidi).Segs(0).Vertsi(nVertsi)
                Solidi(nSolidi).Segs(0).nVertsi = nVertsi
                For I = 0 To nVertsi - 1
                    Solidi(nSolidi).Segs(0).Vertsi(I) = Vertsi(I)
                Next
            Else
                Solidi(nSolidi).nSegs = nSegs
                ReDim Solidi(nSolidi).Segs(nSegs)
                For I = 0 To nSegs - 1
                    ReDim Solidi(nSolidi).Segs(I).Vertsi(lSegs(I).nVertsi)
                    LSet Solidi(nSolidi).Segs(I) = lSegs(I)
                Next
            End If
        End If
        TmpStr = ""
    Case &H4110
        nSolidi = nSolidi + 1
        ReDim Preserve Solidi(nSolidi)
        Get #nFile, , TmpInteger
        nPunti = nPunti + TmpInteger
        Solidi(nSolidi).nVerts = TmpInteger
        ReDim Solidi(nSolidi).Verts(TmpInteger)
        For I = 0 To TmpInteger - 1
            Get #nFile, , TmpSingle
            Solidi(nSolidi).Verts(I).X = TmpSingle
            Get #nFile, , TmpSingle
            Solidi(nSolidi).Verts(I).Y = TmpSingle
            Get #nFile, , TmpSingle
            Solidi(nSolidi).Verts(I).Z = TmpSingle
        Next
        
    ' Triangoli
    Case &H4120
        Get #nFile, , TmpInteger
        nFaccie = nFaccie + CStr(TmpInteger)
        nVertsi = CLng(TmpInteger) * 3
        ReDim Vertsi(CLng(TmpInteger) * 3)
        For I = 0 To CLng(TmpInteger) * 3 - 1 Step 3
            Get #nFile, , Vertsi(I)
            Get #nFile, , Vertsi(I + 1)
            Get #nFile, , Vertsi(I + 2)
            Get #nFile, , iV4
            If (iV4 And 4) = 0 Then
                iV1 = Vertsi(I)
                Vertsi(I) = Vertsi(I + 1)
                Vertsi(I + 1) = iV1
            End If
            If (iV4 And 2) = 0 Then
                iV1 = Vertsi(I + 1)
                Vertsi(I + 1) = Vertsi(I + 2)
                Vertsi(I + 2) = iV1
            End If
            If (iV4 And 1) = 0 Then
                iV1 = Vertsi(I)
                Vertsi(I) = Vertsi(I + 2)
                Vertsi(I + 2) = iV1
            End If
        Next
        
        TmpLong = fPos + 8 + 8 * CLng(TmpInteger)
        Call ComputeNormals
        While TmpLong < (fPos + tch.Length)
            ReadChunk TmpLong, UseLst, tch, Lst
        Wend
    Case &H4160
        '
    Case &H4170
        Get #nFile, , TmpInteger
        TmpStr = CStr(TmpInteger)
        For I = 1 To 21
            Get #nFile, , TmpSingle
            TmpStr = TmpStr & " " & CStr(TmpSingle)
        Next
        TmpStr = ""
    ' Varie sezioni
    Case &HA200, &HA204, &HA210, &HA220, &HA230, &H4100, &H1200, &H3000
        TmpLong = fPos + 6
        While TmpLong < (fPos + tch.Length)
            ReadChunk TmpLong, UseLst, tch, Lst
        Wend
    Case &H4D4D, &HAFFF, &H3D3D
        tch.Length = 6
    Case Else
        'Colori , effetti visivi, sezione Keyframer
        If tch.Header > &HA000 And tch.Header < &HA080 Then
            TmpLong = fPos + 6
        Else
            TmpStr = "* " & (tch.Length - 6)
            TmpStr = ""
        End If
    End Select
    fPos = fPos + tch.Length
    If TmpStr <> "" Then
     If UseLst Then
        Add2List Lst, TmpStr
     End If
    End If
End Sub
Public Sub ComputeNormals()
'hmm
End Sub


Public Sub AzzeraVar()
ReDim Solidi(0)
ReDim Vertsi(0)
xScarto = 0
yScarto = 0
s3dStudioV = ""
sMeshV = ""
nSolidi = 0
nVertsi = 0
nPunti = 0
nFaccie = 0
nSegs = 0
lScale = 0
End Sub

Public Sub RuotaXY(Radi As Double)
Dim Tmp1 As Double
Dim Tmp2 As Double
Dim I As Long, J As Long, K As Long
For I = 1 To nSolidi
For K = 0 To Solidi(I).nVerts - 1
    Tmp1 = Solidi(I).Verts(K).X
    Tmp2 = Solidi(I).Verts(K).Y
    Solidi(I).Verts(K).X = Tmp1 * Cos(Radi) - Tmp2 * Sin(Radi)
    Solidi(I).Verts(K).Y = Tmp1 * Sin(Radi) + Tmp2 * Cos(Radi)
Next
Next
End Sub
Public Sub RuotaXZ(Radi As Double)
Dim Tmp1 As Double
Dim Tmp2 As Double
Dim I As Long, J As Long, K As Long
For I = 1 To nSolidi
For K = 0 To Solidi(I).nVerts - 1
    Tmp1 = Solidi(I).Verts(K).X
    Tmp2 = Solidi(I).Verts(K).Z
    Solidi(I).Verts(K).X = Tmp1 * Cos(Radi) - Tmp2 * Sin(Radi)
    Solidi(I).Verts(K).Z = Tmp1 * Sin(Radi) + Tmp2 * Cos(Radi)
Next
Next
End Sub
Public Sub RuotaYZ(Radi As Double)
Dim Tmp1 As Double
Dim Tmp2 As Double
Dim I As Long, J As Long, K As Long
For I = 1 To nSolidi
For K = 0 To Solidi(I).nVerts - 1
    Tmp1 = Solidi(I).Verts(K).Y
    Tmp2 = Solidi(I).Verts(K).Z
    Solidi(I).Verts(K).Y = Tmp1 * Cos(Radi) - Tmp2 * Sin(Radi)
    Solidi(I).Verts(K).Z = Tmp1 * Sin(Radi) + Tmp2 * Cos(Radi)
Next
Next
End Sub
Public Sub TraslaX(qt As Double)
xScarto = xScarto + qt
End Sub
Public Sub Traslay(qt As Double)
yScarto = yScarto + qt
End Sub
Public Sub Traslaz(qt As Double)
Dim I As Long, J As Long, K As Long
For I = 1 To nSolidi
For K = 0 To Solidi(I).nVerts - 1
    Solidi(I).Verts(K).X = Solidi(I).Verts(K).X * qt
    Solidi(I).Verts(K).Y = Solidi(I).Verts(K).Y * qt
    Solidi(I).Verts(K).Z = Solidi(I).Verts(K).Z * qt
Next
Next
End Sub

Public Sub Render(hdc As Long)
Dim Old As POINTAPI
Dim ppoly() As POINTAPI
Dim I As Long, J As Long, K As Long
FillRect tHdc, R, 255 'GetSysColorBrush(5)
For I = 1 To nSolidi
For K = 0 To Solidi(I).nSegs - 1
ReDim ppoly(Solidi(I).Segs(K).nVertsi - 1)
For J = 0 To Solidi(I).Segs(K).nVertsi - 1
    ppoly(J).X = Solidi(I).Verts(Solidi(I).Segs(K).Vertsi(J)).X + xScarto
    ppoly(J).Y = Solidi(I).Verts(Solidi(I).Segs(K).Vertsi(J)).Y + yScarto
    'If J = 2 Then Exit For
Next
Prova:
MoveToEx hdc, ppoly(0).X, ppoly(0).Y, Old
For J = 0 To UBound(ppoly) - 1
    LineTo hdc, ppoly(J + 1).X, ppoly(J + 1).Y
Next
Next
Next

'On Error Resume Next

End Sub
Public Sub GetKey(Pic As PictureBox, KeyCode As Integer)
Select Case KeyCode
    Case Asc("A")
        TraslaX -5
    Case Asc("D")
        TraslaX 5
    Case Asc("W")
        Traslay -5
    Case Asc("S")
        Traslay 5
    Case Asc("R")
        RuotaXY -pi / 24
    Case Asc("T")
        RuotaXY pi / 24
    Case Asc("F")
        RuotaXZ -pi / 24
    Case Asc("G")
        RuotaXZ pi / 24
    Case Asc("V")
        RuotaYZ -pi / 24
    Case Asc("B")
        RuotaYZ pi / 24
    Case Asc("Q")
        Traslaz 1.05
    Case Asc("E")
        Traslaz 1 / 1.05
    Case vbKeyEscape
        bRunning = False
    End Select
'Pic.Refresh
Render tHdc
BitBlt Pic.hdc, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight, tHdc, 0, 0, vbSrcCopy
End Sub
Public Sub Add2List(Lista() As String, Stringa As String)
ReDim Preserve Lista(UBound(Lista) + 1)
Lista(UBound(Lista)) = Stringa
End Sub
Public Sub CreaBit(Pic As PictureBox)
tPen = CreatePen(0, 1, vbRed)
tHdc = CreateCompatibleDC(Pic.hdc)
tBit = CreateCompatibleBitmap(tHdc, Pic.ScaleWidth, Pic.ScaleHeight)
SizeX = Pic.ScaleWidth
SizeY = Pic.ScaleHeight
DeleteObject SelectObject(tHdc, tPen)
DeleteObject SelectObject(tHdc, tBit)
SetRect R, 0, 0, SizeX, SizeY

End Sub

Public Sub DeleteAll()
DeleteDC tHdc
DeleteObject tBit
DeleteObject tPen
End Sub
