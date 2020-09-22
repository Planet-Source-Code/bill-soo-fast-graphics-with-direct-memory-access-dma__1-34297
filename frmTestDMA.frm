VERSION 5.00
Begin VB.Form frmTestDMA 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
      Begin VB.OptionButton OpFade 
         Caption         =   "Fade"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton opFill 
         Caption         =   "Fill"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opRotate 
         Caption         =   "Rotate"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.OptionButton opFast 
      Caption         =   "Fast DMA"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.OptionButton opDMA 
      Caption         =   "Use DMA"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.OptionButton opAPI 
      Caption         =   "Use API"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   4320
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Timer tmrFPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3960
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox picSrc 
      AutoSize        =   -1  'True
      Height          =   3300
      Left            =   0
      ScaleHeight     =   3240
      ScaleWidth      =   2700
      TabIndex        =   1
      Top             =   240
      Width           =   2760
   End
   Begin VB.PictureBox picDest 
      AutoSize        =   -1  'True
      Height          =   3300
      Left            =   2880
      ScaleHeight     =   3240
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   240
      Width           =   2760
   End
   Begin VB.Label lblMsg 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTestDMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const PI As Double = 3.14159265358979

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim dmaSrc As New bchDirectMemoryAccess.clsDMA   'source DMA
Dim dmaDest As New bchDirectMemoryAccess.clsDMA  'destination DMA
Dim dmaLogo As New bchDirectMemoryAccess.clsDMA  'Logo DMA
Dim MyDir As String                       'app.dir
Dim FPS As Long                           'Frames Per Second counter
Dim bWorking As Boolean                   'flag to indicate we are in loop
Dim theta As Double                       'rotation Angle
Dim aSrc() As Byte                        'DMA byte array of source
Dim aDest() As Byte                       'DMA byte array of Dest
Dim aLogo() As Byte                       'DMA byte array of Logo

Private Sub btnTest_Click()
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, cx As Long, cy As Long
Dim tcos As Double, tsin As Double, tcy As Double, tsy As Double
Dim r As Long, g As Long, b As Long, c As Long, c2 As Long
Dim aTcx() As Double
Dim aTsx() As Double

ReDim aTcx(dmaSrc.UBoundX)
ReDim aTsx(dmaSrc.UBoundX)
cx = dmaSrc.UBoundX \ 2
cy = dmaSrc.UBoundY \ 2
bWorking = Not bWorking
tmrFPS.Enabled = True
btnTest.Caption = "Cancel Test"
Do While bWorking
    FPS = FPS + 1
    If OpFade Then
        theta = (theta + 1) Mod 256   'alpha value
        For y1 = 0 To dmaLogo.UBoundY
            y2 = y1 + (dmaDest.UBoundY - dmaLogo.UBoundY) \ 2
            If (y2 > 0) And (y2 < dmaDest.UBoundY) Then
                For x1 = 0 To dmaLogo.UBoundX
                    x2 = x1 + (dmaDest.UBoundX - dmaLogo.UBoundX) \ 2
                    If (x2 > 0) And (x2 < dmaDest.UBoundX) Then
                        If opAPI Then
                            c = GetPixel(picLogo.hdc, x1, y1)
                            c2 = GetPixel(picSrc.hdc, x2, y2)
                            r = (c And &HFF) * theta / 255
                            g = (c \ 256 And &HFF) * theta / 255
                            b = (c \ &H10000 And &HFF) * theta / 255
                            r = r + (c2 And &HFF): If r > 255 Then r = 255
                            g = g + (c2 \ 256 And &HFF): If g > 255 Then g = 255
                            b = b + (c2 \ &H10000 And &HFF): If b > 255 Then b = 255
                            SetPixel picDest.hdc, x2, y2, RGB(r, g, b)
                        ElseIf opDMA Then
                            c = dmaLogo.ReadPixel(x1, y1)
                            c2 = dmaSrc.ReadPixel(x2, y2)
                            r = (c And &HFF) * theta / 255
                            g = (c \ 256 And &HFF) * theta / 255
                            b = (c \ &H10000 And &HFF) * theta / 255
                            r = r + (c2 And &HFF): If r > 255 Then r = 255
                            g = g + (c2 \ 256 And &HFF): If g > 255 Then g = 255
                            b = b + (c2 \ &H10000 And &HFF): If b > 255 Then b = 255
                            dmaDest.DrawPixel x2, y2, RGB(r, g, b)
                        Else
                            For c = 0 To 2
                                c2 = aSrc(x2 * 3 + c, y2) + aLogo(x1 * 3 + c, y1) * theta / 255
                                If c2 > 255 Then c2 = 255
                                aDest(x2 * 3 + c, y2) = c2
                            Next c
                        End If
                    End If
                Next x1
            End If
        Next y1
    Else
        theta = theta + 5 * PI / 180 'rotate another 5 degrees
        tcos = Cos(theta): tsin = Sin(theta)
        For x1 = 0 To dmaSrc.UBoundX
            aTcx(x1) = tcos * (x1 - cx)
            aTsx(x1) = tsin * (x1 - cx)
        Next x1
        For y1 = 0 To dmaSrc.UBoundY
            tcy = cy - tcos * (y1 - cy): tsy = cx + tsin * (y1 - cy)
            For x1 = 0 To dmaSrc.UBoundX
                If opFill Then
                    If opAPI Then
                        SetPixel picDest.hdc, x1, y1, vbRed
                    ElseIf opDMA Then
                        dmaDest.DrawPixel x1, y1, vbRed
                    Else
                        x2 = x1 * 3
                        aDest(x2, y1) = 0
                        aDest(x2 + 1, y1) = 0
                        aDest(x2 + 2, y1) = 255
                    End If
                ElseIf opRotate Then
                    x2 = aTcx(x1) + tsy
                    y2 = aTsx(x1) + tcy
                    If (x2 > -1) And (x2 < dmaDest.UBoundX) And (y2 > -1) And (y2 < dmaDest.UBoundY) Then
                        If opAPI Then
                            SetPixel picDest.hdc, x2, y2, GetPixel(picSrc.hdc, x1, y1)
                        ElseIf opDMA Then 'dma
                            dmaDest.DrawPixel x2, y2, dmaSrc.ReadPixel(x1, y1)
                        Else
                            CopyMemory aDest(x2 * 3, y2), aSrc(x1 * 3, y1), 3
                        End If
                    End If
                End If
            Next x1
        Next y1
    End If
    picDest.Refresh
    DoEvents
Loop
tmrFPS.Enabled = False
btnTest.Caption = "Test"
End Sub

Private Sub Form_Load()
MyDir = App.Path
If Right$(MyDir, 1) <> "\" Then MyDir = MyDir & "\"
picSrc.Picture = LoadPicture(MyDir & "unknown.jpg")
picDest.Picture = LoadPicture(MyDir & "unknown.jpg")
picLogo.Picture = LoadPicture(MyDir & "logo.bmp")
If dmaSrc.LoadPicArray(picSrc.Picture) Then
    If dmaDest.LoadPicArray(picDest.Picture) Then
        If dmaLogo.LoadPicArray(picLogo.Picture) Then
            dmaSrc.GetData aSrc
            dmaDest.GetData aDest
            dmaLogo.GetData aLogo
        Else
            lblMsg = dmaLogo.ErrorMsg
        End If
    Else
        lblMsg = dmaDest.ErrorMsg
    End If
Else
    lblMsg = dmaSrc.ErrorMsg
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bWorking Then MsgBox "Cancel test first!": Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
dmaSrc.ReleaseData aSrc              'MUST RELEASE ARRAYS FROM DMA
dmaDest.ReleaseData aDest
dmaLogo.ReleaseData aLogo

Set dmaSrc = Nothing
Set dmaDest = Nothing
Set dmaLogo = Nothing
End Sub

Private Sub tmrFPS_Timer()
lblMsg = CStr(FPS) & " fps."
FPS = 0
End Sub
