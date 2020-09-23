VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl CircProg 
   AutoRedraw      =   -1  'True
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   44
   ToolboxBitmap   =   "CircProg.ctx":0000
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483640
      MaskColor       =   -2147483638
      _Version        =   393216
   End
   Begin VB.PictureBox pIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   480
   End
End
Attribute VB_Name = "CircProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CircRadius As Double
Private CenterX As Double
Private CenterY As Double
Private PI As Double
Private ICol(3) As Long
Private OCol(3) As Long
Private mShowIcon As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Function GetX(ByVal Ang As Double, CircRadius As Double) As Double
 GetX = CircRadius * Sin(Radians(180 - Ang)) + CenterX
End Function
Private Function GetY(ByVal Ang As Double, CircRadius As Double) As Double
 GetY = CircRadius * Cos(Radians(180 - Ang)) + CenterY
End Function
Private Function Radians(ByVal Deg As Double) As Double
 Radians = (Deg * PI) / 180
End Function

Private Sub UserControl_Initialize()
 ICol(0) = &HC0C0C0: ICol(1) = &H808080
 ICol(2) = &H404040: ICol(3) = 0
 OCol(0) = &HC8D0C8: OCol(1) = &HC8D0CC
 OCol(2) = &HC8D0D0: OCol(3) = &HC8D0D4
 FillStyle = vbFSSolid
 PI = Atn(1) * 4
 CircRadius = 14
 CenterX = ScaleWidth / 2: CenterY = ScaleHeight / 2
End Sub
Private Sub Timer1_Timer()
 Static StartAngle As Long, ko As Long
 Dim i As Long, k As Long
 Cls
 FillColor = OCol(ko)
 For i = 0 To 359 Step 30 'outer circle 12
  Circle (GetX(i, CircRadius), GetY(i, CircRadius)), 4, OCol(ko)
 Next
 ko = (ko + 1) Mod 4
 For i = StartAngle To StartAngle + 90 Step 30
  FillColor = ICol(k)
  Circle (GetX(i, CircRadius), GetY(i, CircRadius)), 2, ICol(k)
  k = k + 1
 Next
 StartAngle = (StartAngle + 30) Mod 360
 If mShowIcon Then
  With UserControl
   StretchBlt pIcon.hdc, 0, 0, pIcon.ScaleWidth, pIcon.ScaleHeight, _
     .hdc, 0, 0, .ScaleWidth, .ScaleHeight, vbSrcCopy
   pIcon.Refresh
   imlIcon.ListImages.Clear
   imlIcon.ListImages.Add , , pIcon.Image
   Set UserControl.Parent.Icon = imlIcon.ListImages(1).ExtractIcon
  End With
 End If
End Sub

Private Sub UserControl_InitProperties()
 mShowIcon = True
End Sub

Private Sub UserControl_Resize()
 Static Busy As Boolean
' Restrict size to desired dimensions
 If Not Busy Then
  Busy = True
  UserControl.Width = 660
  UserControl.Height = 660
  Busy = False
 End If
 Timer1_Timer
End Sub
Public Sub StartProgress()
 Timer1.Enabled = True
End Sub
Public Sub StopProgress()
 Timer1.Enabled = False
End Sub
Public Sub SetSpeed(ByVal NewV1To100 As Long)
 If NewV1To100 >= 1 And NewV1To100 <= 100 Then
  Timer1.Interval = 100 - NewV1To100 + 1
 End If
End Sub
Public Property Get ShowIcon() As Boolean
 ShowIcon = mShowIcon
End Property
Public Property Let ShowIcon(ByVal NewV As Boolean)
 mShowIcon = NewV
 If NewV = False Then
  Set UserControl.Parent.Icon = Nothing
 End If
 PropertyChanged "Icon"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 mShowIcon = PropBag.ReadProperty("Icon", False)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 PropBag.WriteProperty "Icon", mShowIcon, False
End Sub
Private Sub UserControl_Terminate()
 Timer1.Enabled = False
End Sub

