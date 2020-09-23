VERSION 5.00
Begin VB.UserControl SuperRuler 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SuperRuler.ctx":0000
   Begin VB.Menu mnuScaleModeMenu 
      Caption         =   "ScaleModeMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Twips"
         Index           =   0
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Pixels"
         Index           =   1
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Milimeters"
         Index           =   2
      End
      Begin VB.Menu mnuScaleMode 
         Caption         =   "Inches"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SuperRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***[Enumerations]***************************************************************************************************
Public Enum enuOrientation
    orHorizontal = 0
    orVertical = 1
End Enum

Public Enum enuScaleMode
    smTwips = 0
    smPixels = 1
    smMilimeters = 2
    smInches = 3
End Enum

Public Enum enuBorderStyle
    bsNoBorder = 0
    bsSingle = 1
End Enum

'***[Default Constants]******************************************************************************************************
Private Const mvar_def_Orientation As Long = orHorizontal
Private Const mvar_def_BorderStyle As Long = bsNoBorder
Private Const mvar_def_ScaleMode As Long = smTwips
Private Const mvar_def_MouseTrackingOn As Boolean = False
Private Const mvar_def_StartValue As Long = 0


'***[Shared Variables]******************************************************************************************************
Private mvarOrientation As Long
Private mvarBorderStyle As Long
Private mvarScaleMode As Long
Private mvarMouseTrackingOn As Boolean
Private mvarStartValue As Long


'***[Storage Variables]******************************************************************************************************
Private mvarScale As Long

'***[Events]*********************************************************************************************************
Public Event ScaleModeChanged(Mode As enuScaleMode)
Public Event HooverValue(Value As Long)
Public Event Click(Button As Integer, Shift As Integer, Value As Long)
Public Event Resize()

'***[Properties]*****************************************************************************************************
Public Property Get Orientation() As enuOrientation
    Orientation = mvarOrientation
End Property

Public Property Let Orientation(ByVal Value As enuOrientation)
    mvarOrientation = Value
    RenderControl
    PropertyChanged "Orientation"
End Property

Public Property Get StartValue() As Long
    StartValue = mvarStartValue
End Property

Public Property Let StartValue(ByVal Value As Long)
    mvarStartValue = Value
    RenderControl
    PropertyChanged "StartValue"
End Property

Public Property Get BorderStyle() As enuBorderStyle
    BorderStyle = mvarBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As enuBorderStyle)
    mvarBorderStyle = Value
    UserControl.BorderStyle = mvarBorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get ScaleMode() As enuScaleMode
    ScaleMode = mvarScaleMode
End Property

Public Property Let ScaleMode(ByVal Value As enuScaleMode)
    Dim i As Long
    
    mvarScaleMode = Value
    
    'Set scaling
    Select Case mvarScaleMode
        Case smTwips
            mvarScale = 1000
        Case smPixels
            mvarScale = Screen.TwipsPerPixelX * 100
        Case smMilimeters
            mvarScale = 570
        Case smInches
            mvarScale = 1440
    End Select
    
    For i = 0 To 3
        mnuScaleMode(i).Checked = False
    Next i
    mnuScaleMode(Value).Checked = True
    
    RenderControl
    PropertyChanged "ScaleMode"
    RaiseEvent ScaleModeChanged(Value)
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    RenderControl
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    RenderControl
    PropertyChanged "BackColor"
End Property

Public Property Get MouseTrackingOn() As Boolean
    MouseTrackingOn = mvarMouseTrackingOn
End Property

Public Property Let MouseTrackingOn(ByVal Value As Boolean)
    mvarMouseTrackingOn = Value
    PropertyChanged "MouseTrackingOn"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mvarOrientation = PropBag.ReadProperty("Orientation", mvar_def_Orientation)
    mvarStartValue = PropBag.ReadProperty("StartValue", mvar_def_StartValue)
    BorderStyle = PropBag.ReadProperty("BorderStyle", mvar_def_BorderStyle)
    mvarMouseTrackingOn = PropBag.ReadProperty("MouseTrackingOn", mvar_def_MouseTrackingOn)
    ScaleMode = PropBag.ReadProperty("ScaleMode", mvar_def_ScaleMode)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    RenderControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Orientation", mvarOrientation, mvar_def_Orientation)
    Call PropBag.WriteProperty("StartValue", mvarStartValue, mvar_def_StartValue)
    Call PropBag.WriteProperty("BorderStyle", mvarBorderStyle, mvar_def_BorderStyle)
    Call PropBag.WriteProperty("MouseTrackingOn", mvarMouseTrackingOn, mvar_def_MouseTrackingOn)
    Call PropBag.WriteProperty("ScaleMode", mvarScaleMode, mvar_def_ScaleMode)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
End Sub


Private Sub UserControl_Initialize()
    ScaleMode = smTwips
End Sub

Private Sub UserControl_Resize()
    RenderControl
    RaiseEvent Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuScaleModeMenu
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RenderTrackLine X, Y
    RaiseEvent HooverValue(CalculateValue(X, Y))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Click(Button, Shift, CalculateValue(X, Y))
End Sub

Private Function CalculateValue(X As Single, Y As Single) As Long
    Dim myValue As Long
    Select Case mvarOrientation
    Case orHorizontal
        myValue = Int(X / (mvarScale / 10))
    Case orVertical
        myValue = Int(Y / (mvarScale / 10))
    End Select
    myValue = myValue + mvarStartValue * 10
    Select Case mvarScaleMode
        Case smTwips
            myValue = myValue * 100
        Case smPixels
            myValue = myValue * 10
        Case smMilimeters
            myValue = myValue
        Case smInches
            myValue = Int(myValue / 10)
    End Select
    
    CalculateValue = myValue
End Function

Public Sub RenderTrackLine(X As Single, Y As Single)
    If mvarMouseTrackingOn = True Then
        RenderControl
        'Optionaly render Mouse tracking line
        Select Case Orientation
        Case orHorizontal
            Line (X, 0)-(X, ScaleHeight)
        Case orVertical
            Line (0, Y)-(ScaleWidth, Y)
        End Select
    End If
End Sub

Private Sub mnuScaleMode_Click(Index As Integer)
    ScaleMode = Index
    RenderControl
End Sub

Public Sub Refresh()
    RenderControl
End Sub

Private Sub RenderControl()
    Dim mySmallScale As Long
    Dim myValue As String
    Dim i As Long
    Dim j As Long
    mySmallScale = mvarScale / 10
    
    Cls
    Select Case mvarOrientation
    Case orHorizontal
        For j = 0 To Width Step mvarScale
            'Draw big line
            Line (j, 0)-(j, ScaleHeight)
            'Print Value
            myValue = j / mvarScale
            CurrentY = 0
            CurrentX = CurrentX + 30
            Print myValue + StartValue
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (i, ScaleHeight / 2)-(i, ScaleHeight)
                Else
                    Line (i, ScaleHeight - ScaleHeight / 3)-(i, ScaleHeight)
                End If
            Next i
        Next j
        
    Case orVertical
        For j = 0 To Height Step mvarScale
            'Draw big line
            Line (0, j)-(ScaleWidth, j)
            'Print Value
            myValue = j / mvarScale
            CurrentY = CurrentY + 30
            CurrentX = 0
            Print myValue + StartValue
            'Draw small lines
            For i = j + mySmallScale To j + mvarScale - mySmallScale Step mySmallScale
                If i = j + mvarScale / 2 Then
                    Line (ScaleWidth / 2, i)-(ScaleWidth, i)
                Else
                    Line (ScaleWidth - ScaleWidth / 3, i)-(ScaleWidth, i)
                End If
            Next i
        Next j
        
    End Select
    
    
End Sub

