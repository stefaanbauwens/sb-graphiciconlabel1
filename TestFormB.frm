VERSION 5.00
Object = "{01882636-63CF-44FD-AA95-407759F90746}#1.0#0"; "sbgil1.ocx"
Begin VB.Form TestFormB 
   Caption         =   "TestFormB"
   ClientHeight    =   9375
   ClientLeft      =   8055
   ClientTop       =   1905
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   5895
   Begin VB.CommandButton TestCaption 
      Caption         =   "Lines"
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Tag             =   "I"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.HScrollBar TestCaptionValue 
      Height          =   255
      Left            =   1440
      Max             =   8
      TabIndex        =   29
      Top             =   4560
      Value           =   1
      Width           =   1215
   End
   Begin VB.HScrollBar TestPaddingValue 
      Height          =   255
      Left            =   1440
      Max             =   150
      TabIndex        =   27
      Top             =   3960
      Value           =   1
      Width           =   1215
   End
   Begin VB.CommandButton TestIcon 
      Caption         =   "Icon"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Tag             =   "N"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.HScrollBar TestInsideValue 
      Height          =   255
      Left            =   1440
      Max             =   30
      TabIndex        =   23
      Top             =   2160
      Value           =   1
      Width           =   1215
   End
   Begin VB.HScrollBar TestOutsideValue 
      Height          =   255
      Left            =   1440
      Max             =   30
      TabIndex        =   22
      Top             =   960
      Value           =   1
      Width           =   1215
   End
   Begin sbgil1.GraphicIconLabel GraphicIconLabel 
      Height          =   3615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6376
      BlackInside     =   5
      BlackOutside    =   5
      FillNormalColor =   16777088
      FillDisabledColor=   8388736
      FillHoverColor  =   16744576
      FillPressColor  =   65280
      BorderNormalColor=   16744703
      BorderHoverColor=   12640511
      BorderPressColor=   8421631
      BorderRadius    =   60
      BorderSize      =   20
      Caption         =   "TestFormB.frx":0000
      CaptionAlignHorizontal=   1
      CaptionAlignVertical=   1
      Cursor          =   1
      ForeNormalColor =   255
      ForeHoverColor  =   32768
      ForePressColor  =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconPadding     =   20
   End
   Begin VB.CheckBox TestOptions 
      Caption         =   "Debug MouseMove"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CheckBox TestOptions 
      Caption         =   "Debug CaptionChange"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   8760
      Width           =   2535
   End
   Begin VB.ListBox TestLogging 
      Height          =   3180
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton TestDefault 
      Caption         =   "Default"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox TestEnabled 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox TestAutoRedraw 
      Caption         =   "Auto Redraw"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1295
   End
   Begin VB.CheckBox TestMultiLines 
      Caption         =   "MultiLines"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox TestWordWrap 
      Caption         =   "Word Wrap"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.HScrollBar TestVerticalValue 
      Height          =   255
      Left            =   1440
      Max             =   150
      Min             =   1
      TabIndex        =   10
      Top             =   3360
      Value           =   1
      Width           =   1215
   End
   Begin VB.HScrollBar TestHorizontalValue 
      Height          =   255
      Left            =   1440
      Max             =   150
      Min             =   1
      TabIndex        =   9
      Top             =   2760
      Value           =   1
      Width           =   1215
   End
   Begin VB.CommandButton TestVertical 
      Caption         =   "Vertical"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton TestHorizontal 
      Caption         =   "Horizontal"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.HScrollBar TestRadiusValue 
      Height          =   255
      Left            =   1440
      Max             =   150
      TabIndex        =   4
      Top             =   360
      Value           =   40
      Width           =   1215
   End
   Begin VB.HScrollBar TestScrollValue 
      Height          =   255
      Left            =   1440
      Max             =   150
      TabIndex        =   3
      Top             =   1560
      Value           =   20
      Width           =   1215
   End
   Begin VB.CommandButton TestFill 
      Caption         =   "Fill"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton TestBorder 
      Caption         =   "Border"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton TestBackground 
      Caption         =   "Background"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer TestTimer 
      Interval        =   100
      Left            =   10680
      Top             =   720
   End
   Begin VB.Label TestCaptionLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label TestPaddingLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label TestInsideLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label TestOutsideLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label TestVerticalLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label TestHorizontalLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label TestRadiusLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label TestScrollLabel 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "TestFormB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RotateCaption As Integer

Private Sub GraphicIconLabel_CaptionChange(ByVal CaptionText As String)

If TestOptions(0).Value = 0 Then Exit Sub

Call AddLogging("CaptionChange")

End Sub

Private Sub GraphicIconLabel_Click()

Call AddLogging("Click")

End Sub

Private Sub GraphicIconLabel_DblClick()

Call AddLogging("DblClick")

End Sub

Private Sub GraphicIconLabel_HoverChange(IsHovered As Boolean)

Call AddLogging("HoverChange " & IsHovered)

End Sub

Private Sub GraphicIconLabel_MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

If TestOptions(1).Value = 0 Then Exit Sub

Call AddLogging("MouseMove " & MouseState & " " & MouseX & " " & MouseY)

End Sub

Private Sub TestAutoRedraw_Click()

If TestAutoRedraw.Value = 0 Then
    GraphicIconLabel.AutoRedraw = False
Else
    GraphicIconLabel.AutoRedraw = True
End If

End Sub

Private Sub TestCaption_Click()

If TestCaption.Tag = "I" Then
    TestCaption.Tag = "A"
    TestCaptionValue.Value = GraphicIconLabel.CaptionLinesMaximum
Else
    TestCaption.Tag = "I"
    TestCaptionValue.Value = GraphicIconLabel.CaptionLinesMinimum
End If

End Sub

Private Sub TestCaptionValue_Change()

If TestCaption.Tag = "I" Then
    GraphicIconLabel.CaptionLinesMinimum = TestCaptionValue.Value
Else
    GraphicIconLabel.CaptionLinesMaximum = TestCaptionValue.Value
End If

TestCaptionLabel = "Min/Max " & GraphicIconLabel.CaptionLinesMinimum & "/" & GraphicIconLabel.CaptionLinesMaximum

End Sub

Private Sub TestCaptionValue_Scroll()

Call TestCaptionValue_Change

End Sub

Private Sub TestDefault_Click()

Call GraphicIconLabel.SetDefaults

TestAutoRedraw.Value = 1
TestWordWrap.Value = 1
TestEnabled.Value = 1

TestScrollValue.Value = GraphicIconLabel.BorderSize
TestRadiusValue.Value = GraphicIconLabel.BorderRadius
TestHorizontalValue.Value = GraphicIconLabel.CaptionPaddingHorizontal
TestVerticalValue.Value = GraphicIconLabel.CaptionPaddingVertical
TestOutsideValue.Value = GraphicIconLabel.BlackSizeOutside
TestInsideValue.Value = GraphicIconLabel.BlackSizeInside
TestPaddingValue.Value = GraphicIconLabel.IconPadding

If TestCaption.Tag = "I" Then
    TestCaptionValue.Value = GraphicIconLabel.CaptionLinesMinimum
Else
    TestCaptionValue.Value = GraphicIconLabel.CaptionLinesMaximum
End If

End Sub

Private Sub TestWordWrap_Click()

If TestWordWrap.Value = 0 Then
    GraphicIconLabel.WordWrap = False
Else
    GraphicIconLabel.WordWrap = True
End If

End Sub

Private Sub Form_Load()

RotateCaption = 0

If GraphicIconLabel.IsValidIcon Then
    TestIcon.Tag = "Y"
Else
    TestIcon.Tag = "N"
End If

Call TestRadiusValue_Change
Call TestScrollValue_Change
Call TestVerticalValue_Change
Call TestHorizontalValue_Change
Call TestOutsideValue_Change
Call TestInsideValue_Change
Call TestPaddingValue_Change
Call TestCaptionValue_Change
Call TestWordWrap_Click

TestFormB.Show

End Sub

Private Sub Form_Resize()

If TestFormB.Width > 3120 Then GraphicIconLabel.Width = TestFormB.Width - 3120
If TestFormB.Height > 780 Then GraphicIconLabel.Height = TestFormB.Height - 780

End Sub

Private Sub TestBackground_Click()

If GraphicIconLabel.BackColor = &H8000000F Then
    GraphicIconLabel.BackColor = &HFFFFFF
    TestBackground.BackColor = &HFFFFFF
Else
    GraphicIconLabel.BackColor = &H8000000F
    TestBackground.BackColor = &H8000000F
End If

End Sub

Private Sub TestBorder_Click()

If GraphicIconLabel.BorderNormalColor = &H80FFFF Then
    GraphicIconLabel.BorderNormalColor = &HFF80FF
Else
    GraphicIconLabel.BorderNormalColor = &H80FFFF
End If

End Sub

Private Sub TestEnabled_Click()

If TestEnabled.Value = 0 Then
    GraphicIconLabel.Enabled = False
Else
    GraphicIconLabel.Enabled = True
End If

End Sub

Private Sub TestFill_Click()

If GraphicIconLabel.FillNormalColor = &HFFFF80 Then
    GraphicIconLabel.FillNormalColor = &HFF
Else
    GraphicIconLabel.FillNormalColor = &HFFFF80
End If

End Sub

Private Sub TestHorizontal_Click()

If GraphicIconLabel.CaptionAlignHorizontal = 2 Then
    GraphicIconLabel.CaptionAlignHorizontal = 0
Else
    GraphicIconLabel.CaptionAlignHorizontal = GraphicIconLabel.CaptionAlignHorizontal + 1
End If

End Sub

Private Sub TestVertical_Click()

If GraphicIconLabel.CaptionAlignVertical = 2 Then
    GraphicIconLabel.CaptionAlignVertical = 0
Else
    GraphicIconLabel.CaptionAlignVertical = GraphicIconLabel.CaptionAlignVertical + 1
End If

End Sub

Private Sub TestRadiusValue_Change()

GraphicIconLabel.BorderRadius = TestRadiusValue.Value
TestRadiusLabel.Caption = "Radius " & Trim$(Str$(TestRadiusValue.Value))

End Sub

Private Sub TestRadiusValue_Scroll()

Call TestRadiusValue_Change

End Sub

Private Sub TestScrollValue_Change()

GraphicIconLabel.BorderSize = TestScrollValue.Value
TestScrollLabel.Caption = "Border " & Trim$(Str$(TestScrollValue.Value))

End Sub

Private Sub TestScrollValue_Scroll()

Call TestScrollValue_Change

End Sub

Private Sub TestInsideValue_Change()

GraphicIconLabel.BlackSizeInside = TestInsideValue.Value
TestInsideLabel.Caption = "Inside " & Trim$(Str$(TestInsideValue.Value))

End Sub

Private Sub TestInsideValue_Scroll()

Call TestInsideValue_Change

End Sub

Private Sub TestOutsideValue_Change()

GraphicIconLabel.BlackSizeOutside = TestOutsideValue.Value
TestOutsideLabel.Caption = "Outside " & Trim$(Str$(TestOutsideValue.Value))

End Sub

Private Sub TestOutsideValue_Scroll()

Call TestOutsideValue_Change

End Sub

Private Sub TestIcon_Click()

If TestIcon.Tag = "Y" Then
    TestIcon.Tag = "N"
    Set GraphicIconLabel.IconPicture = Nothing
Else
    TestIcon.Tag = "Y"
    Set GraphicIconLabel.IconPicture = LoadPicture("TestIcon.bmp")
End If

End Sub

Private Sub TestPaddingValue_Change()

GraphicIconLabel.IconPadding = TestPaddingValue.Value
TestPaddingLabel.Caption = "Padding " & Trim$(Str$(TestPaddingValue.Value))

End Sub

Private Sub TestPaddingValue_Scroll()

Call TestPaddingValue_Change

End Sub

Private Sub TestTimer_Timer()

Dim CaptionText As String

If RotateCaption < 9 Then
    RotateCaption = RotateCaption + 1
Else
    RotateCaption = 0
End If

CaptionText = Trim$(Str$(Timer))
CaptionText = Left$(CaptionText, InStr(CaptionText + ".", ".") - 1) & " " & RotateCaption

TestFormB.Caption = CaptionText & " " & GraphicIconLabel.hWnd

If TestMultiLines.Value <> 0 Then
    CaptionText = "A " & CaptionText & vbCrLf & "B " & CaptionText & vbCrLf & "C " & CaptionText
Else
    CaptionText = "A " & CaptionText & " B " & CaptionText & " C " & CaptionText
End If

GraphicIconLabel.Caption = CaptionText
GraphicIconLabel.RedrawGraphics

End Sub

Private Sub TestVerticalValue_Change()

GraphicIconLabel.CaptionPaddingVertical = TestVerticalValue.Value
TestVerticalLabel.Caption = "Vertical " & Trim$(Str$(TestVerticalValue.Value))

End Sub

Private Sub TestVerticalValue_Scroll()

TestVerticalValue_Change

End Sub

Private Sub TestHorizontalValue_Change()

GraphicIconLabel.CaptionPaddingHorizontal = TestHorizontalValue.Value
TestHorizontalLabel.Caption = "Horizontal " & Trim$(Str$(TestHorizontalValue.Value))

End Sub

Private Sub TestHorizontalValue_Scroll()

TestHorizontalValue_Change

End Sub

Private Sub AddLogging(LoggingText As String)

If TestLogging.ListCount > 15 Then TestLogging.RemoveItem 0

TestLogging.AddItem LoggingText & " (" & Timer & ")"

End Sub
