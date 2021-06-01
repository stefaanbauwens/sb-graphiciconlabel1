VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseHoverClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public hwnd As Long
Public IsHovered As Boolean
Public HoverTime As Long
Public prevProc As Long

Public Event HoverChange(IsHovered As Boolean)
Public Event MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

Public Sub RaiseHoverChange(IsHovered As Boolean)

RaiseEvent HoverChange(IsHovered)

End Sub

Public Sub RaiseMouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

RaiseEvent MouseMove(MouseState, MouseX, MouseY)

End Sub

Private Sub Class_Initialize()

HoverTime = 400
IsHovered = False

End Sub
