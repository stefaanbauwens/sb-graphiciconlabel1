VERSION 5.00
Begin VB.UserControl GraphicIconLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   FillStyle       =   4  'Upward Diagonal
   ScaleHeight     =   855
   ScaleWidth      =   1650
   ToolboxBitmap   =   "GraphicIconLabel.ctx":0000
   Begin VB.PictureBox UserIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox UserPicture 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "GraphicIconLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Enum GILCaptionAlignHorizontal
    cLeft = 0
    cCenter = 1
    cRight = 2
End Enum
Public Enum GILCaptionAlignVertical
    cTop = 0
    cMiddle = 1
    cBottom = 2
End Enum
Public Enum GILCursor
    cArrow = 0
    cHand = 1
End Enum

Dim GILControls As GILControls

Dim WithEvents GILFont As StdFont
Attribute GILFont.VB_VarHelpID = -1
Dim WithEvents GILMouseTrack As MouseHoverClass
Attribute GILMouseTrack.VB_VarHelpID = -1

Public Event Click()
Public Event DblClick()
Public Event CaptionChange(ByVal CaptionText As String)
Public Event MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)
Public Event HoverChange(IsHovered As Boolean)

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

GILControls.ControlBackColor = PropBag.ReadProperty("BackColor", &H8000000F)

GILControls.ControlBlackInside = PropBag.ReadProperty("BlackInside", 1)
GILControls.ControlBlackOutside = PropBag.ReadProperty("BlackOutside", 1)

GILControls.ControlFillNormalColor = PropBag.ReadProperty("FillNormalColor", &H8000000F)
GILControls.ControlFillDisabledColor = PropBag.ReadProperty("FillDisabledColor", &H8000000F)
GILControls.ControlFillHoverColor = PropBag.ReadProperty("FillHoverColor", &H8000000F)
GILControls.ControlFillPressColor = PropBag.ReadProperty("FillPressColor", &H8000000F)

GILControls.ControlBorderNormalColor = PropBag.ReadProperty("BorderNormalColor", &H646464)
GILControls.ControlBorderDisabledColor = PropBag.ReadProperty("BorderDisabledColor", &H646464)
GILControls.ControlBorderHoverColor = PropBag.ReadProperty("BorderHoverColor", &H646464)
GILControls.ControlBorderPressColor = PropBag.ReadProperty("BorderPressColor", &H646464)
GILControls.ControlBorderRadius = PropBag.ReadProperty("BorderRadius", 0)
GILControls.ControlBorderSize = PropBag.ReadProperty("BorderSize", 0)

GILControls.ControlButtonSize = PropBag.ReadProperty("ButtonSize", 0)

GILControls.ControlCaptionBytes = PropBag.ReadProperty("Caption", GILDefaultCaption)
GILControls.ControlCaptionAlignHorizontal = PropBag.ReadProperty("CaptionAlignHorizontal", 0)
GILControls.ControlCaptionAlignVertical = PropBag.ReadProperty("CaptionAlignVertical", 0)
GILControls.ControlCaptionPaddingHorizontal = PropBag.ReadProperty("CaptionPaddingHorizontal", 1)
GILControls.ControlCaptionPaddingVertical = PropBag.ReadProperty("CaptionPaddingVertical", 1)
GILControls.ControlCaptionLinesMinimum = PropBag.ReadProperty("CaptionLinesMinimum", 0)
GILControls.ControlCaptionLinesMaximum = PropBag.ReadProperty("CaptionLinesMaximum", 8)
GILControls.ControlCaptionChanged = True

GILControls.ControlCursorNumber = PropBag.ReadProperty("Cursor", 0)

GILControls.ControlForeNormalColor = PropBag.ReadProperty("ForeNormalColor", &H80000012)
GILControls.ControlForeDisabledColor = PropBag.ReadProperty("ForeDisabledColor", &H80000012)
GILControls.ControlForeHoverColor = PropBag.ReadProperty("ForeHoverColor", &H80000012)
GILControls.ControlForePressColor = PropBag.ReadProperty("ForePressColor", &H80000012)

GILControls.ControlWordWrap = PropBag.ReadProperty("WordWrap", True)

GILControls.ControlAutoRedraw = PropBag.ReadProperty("AutoRedraw", True)

Set GILFont = PropBag.ReadProperty("Font", Ambient.Font)

GILControls.ControlIconPadding = PropBag.ReadProperty("IconPadding", 1)
GILControls.ControlIconValid = PropBag.ReadProperty("IconValid", False)

If GILControls.ControlIconValid = True Then Set UserIcon = PropBag.ReadProperty("IconPicture", Nothing)

GILControls.ControlEnabled = PropBag.ReadProperty("Enabled", True)

Set GILMouseTrack = New MouseHoverClass

GILMouseTrack.hwnd = UserPicture.hwnd

GILMouseTrack.HoverTime = PropBag.ReadProperty("HoverTime", 400)

If Ambient.UserMode Then
    StartTrack GILMouseTrack
End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("BackColor", GILControls.ControlBackColor, &H8000000F)

Call PropBag.WriteProperty("BlackInside", GILControls.ControlBlackInside, 1)
Call PropBag.WriteProperty("BlackOutside", GILControls.ControlBlackOutside, 1)

Call PropBag.WriteProperty("FillNormalColor", GILControls.ControlFillNormalColor, &H8000000F)
Call PropBag.WriteProperty("FillDisabledColor", GILControls.ControlFillDisabledColor, &H8000000F)
Call PropBag.WriteProperty("FillHoverColor", GILControls.ControlFillHoverColor, &H8000000F)
Call PropBag.WriteProperty("FillPressColor", GILControls.ControlFillPressColor, &H8000000F)

Call PropBag.WriteProperty("BorderNormalColor", GILControls.ControlBorderNormalColor, &H646464)
Call PropBag.WriteProperty("BorderDisabledColor", GILControls.ControlBorderDisabledColor, &H646464)
Call PropBag.WriteProperty("BorderHoverColor", GILControls.ControlBorderHoverColor, &H646464)
Call PropBag.WriteProperty("BorderPressColor", GILControls.ControlBorderPressColor, &H646464)
Call PropBag.WriteProperty("BorderRadius", GILControls.ControlBorderRadius, 0)
Call PropBag.WriteProperty("BorderSize", GILControls.ControlBorderSize, 0)

Call PropBag.WriteProperty("ButtonSize", GILControls.ControlButtonSize, 0)

Call PropBag.WriteProperty("Caption", GILControls.ControlCaptionBytes, GILDefaultCaption)
Call PropBag.WriteProperty("CaptionAlignHorizontal", GILControls.ControlCaptionAlignHorizontal, 0)
Call PropBag.WriteProperty("CaptionAlignVertical", GILControls.ControlCaptionAlignVertical, 0)
Call PropBag.WriteProperty("CaptionPaddingHorizontal", GILControls.ControlCaptionPaddingHorizontal, 1)
Call PropBag.WriteProperty("CaptionPaddingVertical", GILControls.ControlCaptionPaddingVertical, 1)
Call PropBag.WriteProperty("CaptionLinesMinimum", GILControls.ControlCaptionLinesMinimum, 0)
Call PropBag.WriteProperty("CaptionLinesMaximum", GILControls.ControlCaptionLinesMaximum, 8)

Call PropBag.WriteProperty("Cursor", GILControls.ControlCursorNumber, 0)

Call PropBag.WriteProperty("ForeNormalColor", GILControls.ControlForeNormalColor, &H80000012)
Call PropBag.WriteProperty("ForeDisabledColor", GILControls.ControlForeDisabledColor, &H80000012)
Call PropBag.WriteProperty("ForeHoverColor", GILControls.ControlForeHoverColor, &H80000012)
Call PropBag.WriteProperty("ForePressColor", GILControls.ControlForePressColor, &H80000012)

Call PropBag.WriteProperty("WordWrap", GILControls.ControlWordWrap, True)

Call PropBag.WriteProperty("AutoRedraw", GILControls.ControlAutoRedraw, True)

Call PropBag.WriteProperty("Font", GILFont, UserControl.Ambient.Font)

Call PropBag.WriteProperty("IconPadding", GILControls.ControlIconPadding, 1)
Call PropBag.WriteProperty("IconValid", GILControls.ControlIconValid, False)

If GILControls.ControlIconValid = True Then Call PropBag.WriteProperty("IconPicture", UserIcon, Nothing)

Call PropBag.WriteProperty("Enabled", GILControls.ControlEnabled, True)

Call PropBag.WriteProperty("HoverTime", GILMouseTrack.HoverTime, 400)

End Sub

Private Sub UserControl_InitProperties()

Call SetDefaults

Set GILMouseTrack = New MouseHoverClass

GILMouseTrack.HoverTime = 400

End Sub

Private Sub UserControl_Show()

RedrawGraphics

End Sub

Private Sub UserControl_Resize()

UserPicture.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

RedrawGraphics

End Sub

Private Sub UserControl_Terminate()

EndTrack GILMouseTrack

Set GILMouseTrack = Nothing

End Sub

Private Sub GILMouseTrack_HoverChange(IsHovered As Boolean)

If IsHovered = True Then
    Call CheckMouseCursor(GILControls.ControlCursorNumber)
Else
    Call CheckMouseCursor(0)
End If

GILControls.ControlCurrentHover = IsHovered

RedrawGraphics

RaiseEvent HoverChange(IsHovered)

End Sub

Private Sub GILMouseTrack_MouseMove(MouseState As Long, MouseX As Long, MouseY As Long)

RaiseEvent MouseMove(MouseState, MouseX, MouseY)

End Sub

Private Sub UserPicture_DblClick()

If GILControls.ControlEnabled = False Then Exit Sub
If GILControls.ControlPreviousButtons <> 1 Then Exit Sub

GILControls.ControlDblClickTiming = Timer + 0.1

Call SimulateMouseDown(1, 0)

RaiseEvent DblClick

End Sub

Private Sub UserPicture_MouseDown(MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

Call SimulateMouseDown(MouseButton, MouseShift)

End Sub

Private Sub UserPicture_MouseUp(MouseButton As Integer, MouseShift As Integer, MouseX As Single, MouseY As Single)

GILControls.ControlCurrentButtons = (GILControls.ControlCurrentButtons Or MouseButton) - MouseButton
GILControls.ControlPreviousButtons = MouseButton

If GILControls.ControlEnabled = False Then Exit Sub
If MouseButton <> 1 Then Exit Sub

If Timer > GILControls.ControlDblClickTiming Then RaiseEvent Click

If GILControls.ControlCurrentPress = False Then Exit Sub

GILControls.ControlCurrentPress = False

RedrawGraphics

End Sub

Public Property Get hwnd() As Long

hwnd = UserControl.hwnd

End Property

Public Property Get HoverTime() As Long

HoverTime = GILMouseTrack.HoverTime

End Property

Public Property Let HoverTime(newHoverTime As Long)

GILMouseTrack.HoverTime = newHoverTime

End Property

Public Property Get ButtonSize() As Integer

ButtonSize = GILControls.ControlButtonSize

End Property

Public Property Let ButtonSize(ByVal NewButtonSize As Integer)

If NewButtonSize < 0 Then NewButtonSize = 0

GILControls.ControlButtonSize = NewButtonSize

RedrawGraphics

End Property

Public Property Get BorderSize() As Integer

BorderSize = GILControls.ControlBorderSize

End Property

Public Property Let BorderSize(ByVal NewBorderSize As Integer)

If NewBorderSize < 0 Then NewBorderSize = 0

GILControls.ControlBorderSize = NewBorderSize

RedrawGraphics

End Property

Public Property Get HexCaption() As String

HexCaption = DoGetHexCaption(GILControls)

End Property

Public Property Let HexCaption(NewHexCaption As String)

Call DoLetHexCaption(GILControls, NewHexCaption)

RedrawGraphics

RaiseEvent CaptionChange(GILControls.ControlCaptionBytes)

End Property

Public Property Get Caption() As String

Caption = GILControls.ControlCaptionBytes

End Property

Public Property Let Caption(ByVal NewCaption As String)

GILControls.ControlCaptionBytes = NewCaption
GILControls.ControlCaptionChanged = True

RedrawGraphics

RaiseEvent CaptionChange(GILControls.ControlCaptionBytes)

End Property

Public Property Get CaptionAlignHorizontal() As GILCaptionAlignHorizontal

CaptionAlignHorizontal = GILControls.ControlCaptionAlignHorizontal

End Property

Public Property Let CaptionAlignHorizontal(ByVal NewCaptionAlignHorizontal As GILCaptionAlignHorizontal)

GILControls.ControlCaptionAlignHorizontal = NewCaptionAlignHorizontal

RedrawGraphics

End Property

Public Property Get CaptionAlignVertical() As GILCaptionAlignVertical

CaptionAlignVertical = GILControls.ControlCaptionAlignVertical

End Property

Public Property Let CaptionAlignVertical(ByVal NewCaptionAlignVertical As GILCaptionAlignVertical)

GILControls.ControlCaptionAlignVertical = NewCaptionAlignVertical

RedrawGraphics

End Property

Public Property Get CaptionPaddingHorizontal() As Integer

CaptionPaddingHorizontal = GILControls.ControlCaptionPaddingHorizontal

End Property

Public Property Let CaptionPaddingHorizontal(ByVal NewCaptionPadding As Integer)

If NewCaptionPadding < 1 Then NewCaptionPadding = 1

GILControls.ControlCaptionPaddingHorizontal = NewCaptionPadding

RedrawGraphics

End Property

Public Property Get CaptionPaddingVertical() As Integer

CaptionPaddingVertical = GILControls.ControlCaptionPaddingVertical

End Property

Public Property Let CaptionPaddingVertical(ByVal NewCaptionPadding As Integer)

If NewCaptionPadding < 1 Then NewCaptionPadding = 1

GILControls.ControlCaptionPaddingVertical = NewCaptionPadding

RedrawGraphics

End Property

Public Property Get CaptionLinesMinimum() As Integer

CaptionLinesMinimum = GILControls.ControlCaptionLinesMinimum

End Property

Public Property Let CaptionLinesMinimum(ByVal NewLines As Integer)

If NewLines < 0 Then NewLines = 0
If NewLines > 8 Then NewLines = 8

GILControls.ControlCaptionLinesMinimum = NewLines

If GILControls.ControlCaptionLinesMaximum < NewLines Then GILControls.ControlCaptionLinesMaximum = NewLines

RedrawGraphics

End Property

Public Property Get CaptionLinesMaximum() As Integer

CaptionLinesMaximum = GILControls.ControlCaptionLinesMaximum

End Property

Public Property Let CaptionLinesMaximum(ByVal NewLines As Integer)

If NewLines < 0 Then NewLines = 0
If NewLines > 8 Then NewLines = 8

GILControls.ControlCaptionLinesMaximum = NewLines

If GILControls.ControlCaptionLinesMinimum > NewLines Then GILControls.ControlCaptionLinesMinimum = NewLines

RedrawGraphics

End Property

Public Property Get AutoRedraw() As Boolean

AutoRedraw = GILControls.ControlAutoRedraw

End Property

Public Property Let AutoRedraw(ByVal NewAutoRedraw As Boolean)

GILControls.ControlAutoRedraw = NewAutoRedraw

RedrawGraphics

End Property

Public Property Get WordWrap() As Boolean

WordWrap = GILControls.ControlWordWrap

End Property

Public Property Let WordWrap(ByVal NewWordWrap As Boolean)

GILControls.ControlWordWrap = NewWordWrap

RedrawGraphics

End Property

Public Property Get BorderRadius() As Integer

BorderRadius = GILControls.ControlBorderRadius

End Property

Public Property Let BorderRadius(ByVal NewBorderRadius As Integer)

If NewBorderRadius < 0 Then NewBorderRadius = 0

GILControls.ControlBorderRadius = NewBorderRadius

RedrawGraphics

End Property

Public Property Get Cursor() As GILCursor

Cursor = GILControls.ControlCursorNumber

End Property

Public Property Let Cursor(ByVal NewCursor As GILCursor)

GILControls.ControlCursorNumber = NewCursor

RedrawGraphics

End Property

Public Property Get FillNormalColor() As OLE_COLOR

FillNormalColor = GILControls.ControlFillNormalColor

End Property

Public Property Let FillNormalColor(ByVal NewFillNormalColor As OLE_COLOR)

GILControls.ControlFillNormalColor = NewFillNormalColor

RedrawGraphics

End Property

Public Property Get FillDisabledColor() As OLE_COLOR

FillDisabledColor = GILControls.ControlFillDisabledColor

End Property

Public Property Let FillDisabledColor(ByVal NewFillDisabledColor As OLE_COLOR)

GILControls.ControlFillDisabledColor = NewFillDisabledColor

RedrawGraphics

End Property

Public Property Get FillHoverColor() As OLE_COLOR

FillHoverColor = GILControls.ControlFillHoverColor

End Property

Public Property Let FillHoverColor(ByVal NewFillHoverColor As OLE_COLOR)

GILControls.ControlFillHoverColor = NewFillHoverColor

RedrawGraphics

End Property

Public Property Get FillPressColor() As OLE_COLOR

FillPressColor = GILControls.ControlFillPressColor

End Property

Public Property Let FillPressColor(ByVal NewFillPressColor As OLE_COLOR)

GILControls.ControlFillPressColor = NewFillPressColor

RedrawGraphics

End Property

Public Property Get ForeNormalColor() As OLE_COLOR

ForeNormalColor = GILControls.ControlForeNormalColor

End Property

Public Property Let ForeNormalColor(ByVal NewForeNormalColor As OLE_COLOR)

GILControls.ControlForeNormalColor = NewForeNormalColor

RedrawGraphics

End Property

Public Property Get ForeDisabledColor() As OLE_COLOR

ForeDisabledColor = GILControls.ControlForeDisabledColor

End Property

Public Property Let ForeDisabledColor(ByVal NewForeDisabledColor As OLE_COLOR)

GILControls.ControlForeDisabledColor = NewForeDisabledColor

RedrawGraphics

End Property

Public Property Get ForeHoverColor() As OLE_COLOR

ForeHoverColor = GILControls.ControlForeHoverColor

End Property

Public Property Let ForeHoverColor(ByVal NewForeHoverColor As OLE_COLOR)

GILControls.ControlForeHoverColor = NewForeHoverColor

RedrawGraphics

End Property

Public Property Get ForePressColor() As OLE_COLOR

ForePressColor = GILControls.ControlForePressColor

End Property

Public Property Let ForePressColor(ByVal NewForePressColor As OLE_COLOR)

GILControls.ControlForePressColor = NewForePressColor

RedrawGraphics

End Property

Public Property Get BackColor() As OLE_COLOR

BackColor = GILControls.ControlBackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

GILControls.ControlBackColor = NewBackColor

RedrawGraphics

End Property

Public Property Get BorderNormalColor() As OLE_COLOR

BorderNormalColor = GILControls.ControlBorderNormalColor

End Property

Public Property Let BorderNormalColor(ByVal NewBorderNormalColor As OLE_COLOR)

GILControls.ControlBorderNormalColor = NewBorderNormalColor

RedrawGraphics

End Property

Public Property Get BorderDisabledColor() As OLE_COLOR

BorderDisabledColor = GILControls.ControlBorderDisabledColor

End Property

Public Property Let BorderDisabledColor(ByVal NewBorderDisabledColor As OLE_COLOR)

GILControls.ControlBorderDisabledColor = NewBorderDisabledColor

RedrawGraphics

End Property

Public Property Get BorderHoverColor() As OLE_COLOR

BorderHoverColor = GILControls.ControlBorderHoverColor

End Property

Public Property Let BorderHoverColor(ByVal NewBorderHoverColor As OLE_COLOR)

GILControls.ControlBorderHoverColor = NewBorderHoverColor

RedrawGraphics

End Property

Public Property Get BorderPressColor() As OLE_COLOR

BorderPressColor = GILControls.ControlBorderPressColor

End Property

Public Property Let BorderPressColor(ByVal NewBorderPressColor As OLE_COLOR)

GILControls.ControlBorderPressColor = NewBorderPressColor

RedrawGraphics

End Property

Public Property Get Enabled() As Boolean

Enabled = GILControls.ControlEnabled

End Property

Public Property Let Enabled(ByVal NewValue As Boolean)

GILControls.ControlEnabled = NewValue

RedrawGraphics

End Property

Public Property Get Font() As StdFont

Set Font = GILFont

End Property

Public Property Set Font(ByVal NewFont As StdFont)

Set GILFont = NewFont

RedrawGraphics

End Property

Public Property Get IconPicture() As Picture
    
Set IconPicture = UserIcon.Picture

End Property

Public Property Set IconPicture(ByVal NewPicture As Picture)

If NewPicture Is Nothing Then

    GILControls.ControlIconValid = False

Else

    GILControls.ControlIconValid = True

    Set UserIcon.Picture = NewPicture

End If

RedrawGraphics

End Property

Public Property Get IconPadding() As Integer
    
IconPadding = GILControls.ControlIconPadding

End Property

Public Property Let IconPadding(ByVal NewGap As Integer)

GILControls.ControlIconPadding = NewGap

RedrawGraphics

End Property

Public Property Get BlackSizeInside() As Integer
    
BlackSizeInside = GILControls.ControlBlackInside

End Property

Public Property Let BlackSizeInside(ByVal NewSize As Integer)

GILControls.ControlBlackInside = NewSize

RedrawGraphics

End Property

Public Property Get BlackSizeOutside() As Integer
    
BlackSizeOutside = GILControls.ControlBlackOutside

End Property

Public Property Let BlackSizeOutside(ByVal NewSize As Integer)

GILControls.ControlBlackOutside = NewSize

RedrawGraphics

End Property

Public Function IsValidIcon() As Boolean

IsValidIcon = GILControls.ControlIconValid

End Function

Public Sub SetDefaults()

Call DoSetDefaults(GILControls)

Set GILFont = Ambient.Font

RedrawGraphics

End Sub

Public Sub RedrawGraphics()

Set UserPicture.Font = GILFont

Call DoRedrawGraphics(GILControls, UserPicture, UserIcon)

UserControl.AutoRedraw = True
UserControl.Refresh

End Sub

Private Sub SimulateMouseDown(MouseButton As Integer, MouseShift As Integer)

Dim MouseFlag As Boolean

GILControls.ControlCurrentButtons = (GILControls.ControlCurrentButtons Or MouseButton)

If GILControls.ControlEnabled = False Then Exit Sub

If (GILControls.ControlCurrentButtons And 1) = 1 Then
    MouseFlag = True
Else
    MouseFlag = False
End If

If GILControls.ControlCurrentPress = MouseFlag Then Exit Sub

GILControls.ControlCurrentPress = MouseFlag

RedrawGraphics

End Sub

Private Sub CheckMouseCursor(CheckMode As Integer)

If CheckMode = 0 Or GILControls.ControlEnabled = False Then

    UserPicture.MousePointer = 0

Else

    Call LoadHandCursor(GILControls)
    
    If Not GILControls.ControlHandCursor Is Nothing Then
        UserPicture.MouseIcon = GILControls.ControlHandCursor
        UserPicture.MousePointer = vbCustom
    End If

End If

End Sub
