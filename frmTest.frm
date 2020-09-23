VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Test"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2700
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSt 
      Caption         =   "Stop"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSt 
      Caption         =   "Start"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1455
   End
   Begin TestCP.CircProg CircProg1 
      Height          =   660
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      Icon            =   -1  'True
   End
   Begin VB.CheckBox chkIcon 
      Caption         =   "Show Icon"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 Dim i As Long
 For i = 1 To 100
  cboSpeed.AddItem i
 Next
 cboSpeed.ListIndex = 49
End Sub

Private Sub cboSpeed_Click()
 CircProg1.SetSpeed cboSpeed.Text
End Sub

Private Sub chkIcon_Click()
 CircProg1.ShowIcon = CBool(chkIcon.Value)
End Sub

Private Sub cmdSt_Click(Index As Integer)
 If Index = 0 Then
  CircProg1.StartProgress
 Else
  CircProg1.StopProgress
 End If
End Sub


