VERSION 5.00
Object = "{A089CF78-B21A-49C2-B741-15E390C259A6}#1.0#0"; "acSR.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   6255
      LargeChange     =   5
      Left            =   9660
      Max             =   20
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   5
      Left            =   540
      Max             =   20
      TabIndex        =   2
      Top             =   6840
      Width           =   9135
   End
   Begin acSR.SuperRuler SuperRuler2 
      Height          =   6300
      Left            =   0
      Top             =   555
      Width           =   540
      _ExtentX        =   0
      _ExtentY        =   0
      Orientation     =   1
      ScaleMode       =   2
      ForeColor       =   -2147483636
      BackColor       =   -2147483624
   End
   Begin acSR.SuperRuler SuperRuler1 
      Height          =   555
      Left            =   555
      Top             =   0
      Width           =   9150
      _ExtentX        =   0
      _ExtentY        =   0
      ScaleMode       =   2
      ForeColor       =   -2147483636
      BackColor       =   -2147483624
   End
   Begin VB.PictureBox Picture1 
      Height          =   6315
      Left            =   540
      ScaleHeight     =   6255
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   540
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "SuperRuler features : ScaleMode with RightClick menu, Fore and Back Color, StartValue, Mouse Tracking, HooverValue, ClickValue"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   7140
      Width           =   9915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change()
    SuperRuler1.StartValue = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    SuperRuler1.StartValue = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    SuperRuler2.StartValue = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    SuperRuler2.StartValue = VScroll1.Value
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SuperRuler1.RenderTrackLine X, Y
    SuperRuler2.RenderTrackLine X, Y
End Sub

Private Sub SuperRuler1_Click(Button As Integer, Shift As Integer, Value As Long)
    If Button = vbLeftButton Then MsgBox Value
End Sub

Private Sub SuperRuler2_Click(Button As Integer, Shift As Integer, Value As Long)
    If Button = vbLeftButton Then MsgBox Value
End Sub

