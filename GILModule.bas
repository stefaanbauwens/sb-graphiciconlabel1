Attribute VB_Name = "GILModule"

Public Const GILDefaultCaption = "GraphicButtonPictureLabel"
Public Const GILDefaultBlack = &H646464

Public Type GILControls
    ControlHandhandle As Long
    ControlHandCursor As StdPicture
    ControlCursorNumber As Integer
    ControlCaptionBytes() As Byte
    ControlCaptionChanged As Boolean
    ControlCaptionRows() As String
    ControlCaptionAlignHorizontal As Integer
    ControlCaptionAlignVertical As Integer
    ControlCaptionPaddingHorizontal As Integer
    ControlCaptionPaddingVertical As Integer
    ControlCaptionLinesMinimum As Integer
    ControlCaptionLinesMaximum As Integer
    ControlBorderSize As Integer
    ControlBorderRadius As Integer
    ControlBorderNormalColor As OLE_COLOR
    ControlBorderDisabledColor As OLE_COLOR
    ControlBorderHoverColor As OLE_COLOR
    ControlBorderPressColor As OLE_COLOR
    ControlForeNormalColor As OLE_COLOR
    ControlForeDisabledColor As OLE_COLOR
    ControlForeHoverColor As OLE_COLOR
    ControlForePressColor As OLE_COLOR
    ControlFillNormalColor As OLE_COLOR
    ControlFillDisabledColor As OLE_COLOR
    ControlFillHoverColor As OLE_COLOR
    ControlFillPressColor As OLE_COLOR
    ControlBackColor As OLE_COLOR
    ControlButtonSize As Integer
    ControlBlackInside As Integer
    ControlBlackOutside As Integer
    ControlPreviousButtons As Integer
    ControlCurrentButtons As Integer
    ControlCurrentHover As Boolean
    ControlCurrentPress As Boolean
    ControlDblClickTiming As Double
    ControlAutoRedraw As Boolean
    ControlWordWrap As Boolean
    ControlIconPadding As Integer
    ControlIconValid As Boolean
    ControlEnabled As Boolean
End Type

Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Function DoGetHexCaption(GetControls As GILControls) As String

Dim GetBytevalue As Byte
Dim GetIndex As Long
Dim GetBytes() As Byte

If UBound(GetControls.ControlCaptionBytes) >= 1 Then
    ReDim GetBytes(UBound(GetControls.ControlCaptionBytes) * 4 + 3)
    For GetIndex = 0 To UBound(GetControls.ControlCaptionBytes)
        GetBytevalue = (GetControls.ControlCaptionBytes(GetIndex) And &HF0) \ &H10
        Select Case GetBytevalue
            Case 0 To 9:   GetBytes(GetIndex * 4) = GetBytevalue Or 48
            Case 10 To 15: GetBytes(GetIndex * 4) = GetBytevalue + 55
        End Select
        GetBytevalue = GetControls.ControlCaptionBytes(GetIndex) And &HF
        Select Case GetBytevalue
            Case 0 To 9:   GetBytes(GetIndex * 4 + 2) = GetBytevalue Or 48
            Case 10 To 15: GetBytes(GetIndex * 4 + 2) = GetBytevalue + 55
        End Select
    Next GetIndex
    DoGetHexCaption = GetBytes
Else
    DoGetHexCaption = vbNullString
End If

End Function

Public Sub DoLetHexCaption(GetControls As GILControls, NewHexCaption As String)

Dim LetIndex As Long
Dim LetBytevalue As Byte
Dim LetBytes() As Byte
Dim LetAscii As Byte

LetBytes = UCase$(NewHexCaption)

If UBound(LetBytes) >= 3 Then
    ReDim GetControls.ControlCaptionBytes(0 To (UBound(LetBytes) - 3) \ 4)
    For LetIndex = 0 To UBound(GetControls.ControlCaptionBytes)
        LetBytevalue = LetBytes(LetIndex * 4)
        Select Case LetBytevalue
            Case 48 To 57: LetAscii = (LetBytevalue - 48) * &H10
            Case 65 To 70: LetAscii = (LetBytevalue - 55) * &H10
            Case Else:     LetAscii = 0
        End Select
        LetBytevalue = LetBytes(LetIndex * 4 + 2)
        Select Case LetBytevalue
            Case 48 To 57: LetAscii = LetAscii Or (LetBytevalue - 48)
            Case 65 To 70: LetAscii = LetAscii Or (LetBytevalue - 55)
        End Select
        GetControls.ControlCaptionBytes(LetIndex) = LetAscii
    Next LetIndex
Else
    GetControls.ControlCaptionBytes = vbNullString
End If

End Sub

Public Sub DoSetDefaults(SetControls As GILControls)

SetControls.ControlBackColor = &H8000000F
SetControls.ControlBorderNormalColor = &H8000000F
SetControls.ControlBorderDisabledColor = &H8000000F
SetControls.ControlBorderHoverColor = &HFFFFFF
SetControls.ControlBorderPressColor = &HFFFFFF
SetControls.ControlBorderRadius = 0
SetControls.ControlBorderSize = 0
SetControls.ControlButtonSize = 2
SetControls.ControlCaptionBytes = GILDefaultCaption
SetControls.ControlCaptionAlignHorizontal = 1
SetControls.ControlCaptionAlignVertical = 1
SetControls.ControlCaptionPaddingHorizontal = 1
SetControls.ControlCaptionPaddingVertical = 1
SetControls.ControlCaptionChanged = True
SetControls.ControlCaptionLinesMinimum = 0
SetControls.ControlCaptionLinesMaximum = 8
SetControls.ControlCursorNumber = 0
SetControls.ControlForeNormalColor = &H80000012
SetControls.ControlForeDisabledColor = &H80000012
SetControls.ControlForeHoverColor = &H80000012
SetControls.ControlForePressColor = &H80000012
SetControls.ControlFillNormalColor = &H8000000F
SetControls.ControlFillDisabledColor = &HC0C0C0
SetControls.ControlFillHoverColor = &H8000000F
SetControls.ControlFillPressColor = &HE0E0E0
SetControls.ControlDblClickTiming = Timer - 1
SetControls.ControlPreviousButtons = 0
SetControls.ControlCurrentHover = False
SetControls.ControlCurrentPress = False
SetControls.ControlCurrentButtons = 0
SetControls.ControlAutoRedraw = True
SetControls.ControlWordWrap = True
SetControls.ControlEnabled = True

End Sub

Public Sub LoadHandCursor(LoadControls As GILControls)

LoadControls.ControlHandhandle = LoadCursor(0, 32649)

If LoadControls.ControlHandhandle <> 0 Then Set LoadControls.ControlHandCursor = HandleToPicture(LoadControls.ControlHandhandle, False)

End Sub

Public Sub DoRedrawGraphics(RedrawControls As GILControls, RedrawPicture As PictureBox, RedrawIcon As PictureBox)

Dim RedrawFill As OLE_COLOR
Dim RedrawFore As OLE_COLOR
Dim RedrawRowIndex As Integer
Dim RedrawTotalRows As Integer
Dim RedrawTotalLimit As Integer
Dim RedrawButtonColor As OLE_COLOR
Dim RedrawBorderColor As OLE_COLOR
Dim RedrawBorderSize As Integer
Dim RedrawBlackInside As Integer
Dim RedrawBlackOutside As Integer
Dim RedrawOutBottom As Integer
Dim RedrawOutRight As Integer
Dim RedrawOutLeft As Integer
Dim RedrawOutTop As Integer
Dim RedrawInBottom As Integer
Dim RedrawInRight As Integer
Dim RedrawInLeft As Integer
Dim RedrawInTop As Integer
Dim RedrawRadius As Integer
Dim RedrawHeight As Integer

If RedrawControls.ControlAutoRedraw = False Then Exit Sub

If RedrawControls.ControlEnabled = False Then
    RedrawBorderColor = RedrawControls.ControlBorderDisabledColor
    RedrawFore = RedrawControls.ControlForeDisabledColor
    RedrawFill = RedrawControls.ControlFillDisabledColor
ElseIf RedrawControls.ControlCurrentPress = True Then
    RedrawBorderColor = RedrawControls.ControlBorderPressColor
    RedrawFore = RedrawControls.ControlForePressColor
    RedrawFill = RedrawControls.ControlFillPressColor
ElseIf RedrawControls.ControlCurrentHover = True Then
    RedrawBorderColor = RedrawControls.ControlBorderHoverColor
    RedrawFore = RedrawControls.ControlForeHoverColor
    RedrawFill = RedrawControls.ControlFillHoverColor
Else
    RedrawBorderColor = RedrawControls.ControlBorderNormalColor
    RedrawFore = RedrawControls.ControlForeNormalColor
    RedrawFill = RedrawControls.ControlFillNormalColor
End If

RedrawRadius = RedrawControls.ControlBorderRadius
RedrawBorderSize = RedrawControls.ControlBorderSize
RedrawBlackOutside = RedrawControls.ControlBlackOutside
RedrawBlackInside = RedrawControls.ControlBlackInside

If RedrawRadius < 0 Then RedrawRadius = 0
If RedrawBorderSize < 0 Then RedrawBorderSize = 0
If RedrawBlackOutside < 0 Then RedrawBlackOutside = 0
If RedrawBlackInside < 0 Then RedrawBlackInside = 0

RedrawOutRight = RedrawPicture.ScaleWidth - 1
RedrawOutBottom = RedrawPicture.ScaleHeight - 1
RedrawOutLeft = 0
RedrawOutTop = 0

If RedrawControls.ControlButtonSize > 1 Then
    
    RedrawOutRight = RedrawOutRight - RedrawControls.ControlButtonSize
    RedrawOutBottom = RedrawOutBottom - RedrawControls.ControlButtonSize
    RedrawOutLeft = RedrawControls.ControlButtonSize
    RedrawOutTop = RedrawControls.ControlButtonSize

End If

RedrawInRight = (RedrawOutRight - RedrawOutLeft) / 2
RedrawInBottom = (RedrawOutBottom - RedrawOutTop) / 2

If RedrawInRight < 0 Then RedrawInRight = 0
If RedrawInBottom < 0 Then RedrawInBottom = 0

RedrawHeight = RedrawBlackOutside + RedrawBorderSize + RedrawBlackInside

If RedrawInRight < RedrawHeight Or RedrawInBottom < RedrawHeight Then
    If RedrawInBottom < RedrawBorderSize Then RedrawBorderSize = RedrawInBottom
    If RedrawInRight < RedrawBorderSize Then RedrawBorderSize = RedrawInRight
    RedrawBlackOutside = 0
    RedrawBlackInside = 0
End If

RedrawHeight = RedrawBlackOutside + RedrawBorderSize + RedrawBlackInside

If RedrawRadius < RedrawHeight Then RedrawRadius = RedrawHeight
If RedrawInRight < RedrawRadius Then RedrawRadius = RedrawInRight
If RedrawInBottom < RedrawRadius Then RedrawRadius = RedrawInBottom

RedrawInRight = RedrawOutRight - RedrawRadius
RedrawInBottom = RedrawOutBottom - RedrawRadius
RedrawInLeft = RedrawOutLeft + RedrawRadius
RedrawInTop = RedrawOutTop + RedrawRadius

RedrawPicture.DrawStyle = 0
RedrawPicture.DrawWidth = 1
RedrawPicture.BackColor = RedrawControls.ControlBackColor
RedrawPicture.FillStyle = 0

RedrawIcon.BackColor = RedrawControls.ControlBackColor

If RedrawBlackOutside > 0 Then Call RedrawRounded(RedrawPicture, RedrawRadius, RedrawInLeft, RedrawInRight, RedrawInTop, RedrawInBottom, RedrawOutLeft, RedrawOutRight, RedrawOutTop, RedrawOutBottom, RedrawBlackOutside, (GBLDefaultBlack))
If RedrawBorderSize > 0 Then Call RedrawRounded(RedrawPicture, RedrawRadius, RedrawInLeft, RedrawInRight, RedrawInTop, RedrawInBottom, RedrawOutLeft, RedrawOutRight, RedrawOutTop, RedrawOutBottom, RedrawBorderSize, RedrawBorderColor)
If RedrawBlackInside > 0 Then Call RedrawRounded(RedrawPicture, RedrawRadius, RedrawInLeft, RedrawInRight, RedrawInTop, RedrawInBottom, RedrawOutLeft, RedrawOutRight, RedrawOutTop, RedrawOutBottom, RedrawBlackInside, (GBLDefaultBlack))

Call RedrawRounded(RedrawPicture, RedrawRadius, RedrawInLeft, RedrawInRight, RedrawInTop, RedrawInBottom, RedrawOutLeft, RedrawOutRight, RedrawOutTop, RedrawOutBottom, 0, RedrawFill)

If RedrawControls.ControlButtonSize > 1 Then RedrawHeight = RedrawHeight + RedrawControls.ControlButtonSize

RedrawBorderSize = RedrawHeight
RedrawHeight = RedrawPicture.TextHeight("X")

RedrawPicture.ForeColor = RedrawFore

If RedrawControls.ControlCaptionChanged = True Then

    RedrawTotalRows = RedrawOutRight - RedrawOutLeft - RedrawPicture.TextWidth("XX")
    
    Call SplitCaption(RedrawControls, RedrawPicture, RedrawTotalRows)

End If

RedrawTotalRows = UBound(RedrawControls.ControlCaptionRows) + 1
RedrawTotalLimit = RedrawTotalRows

If RedrawTotalLimit < RedrawControls.ControlCaptionLinesMinimum Then RedrawTotalLimit = RedrawControls.ControlCaptionLinesMinimum
If RedrawTotalRows > RedrawControls.ControlCaptionLinesMaximum Then RedrawTotalRows = RedrawControls.ControlCaptionLinesMaximum

RedrawRowIndex = 0
While RedrawRowIndex < RedrawTotalRows
    
    Call RedrawText(RedrawControls, RedrawPicture, RedrawControls.ControlCaptionRows(RedrawRowIndex), RedrawRowIndex + 1, RedrawTotalLimit, RedrawBorderSize, RedrawHeight, RedrawIcon)

    RedrawRowIndex = RedrawRowIndex + 1
Wend

If RedrawControls.ControlIconValid = True Then

    Select Case RedrawControls.ControlCaptionAlignVertical
        Case 0: RedrawInTop = RedrawBorderSize + RedrawControls.ControlCaptionPaddingVertical + RedrawControls.ControlIconPadding
        Case 1: RedrawInTop = (RedrawPicture.ScaleHeight - RedrawHeight * RedrawTotalLimit - RedrawControls.ControlCaptionPaddingVertical * (RedrawTotalLimit - 1) - RedrawIcon.ScaleHeight - RedrawControls.ControlCaptionPaddingVertical) / 2
        Case 2: RedrawInTop = RedrawPicture.ScaleHeight - RedrawBorderSize - (RedrawControls.ControlCaptionPaddingVertical + RedrawHeight) * RedrawTotalLimit - RedrawControls.ControlIconPadding - RedrawIcon.ScaleHeight - RedrawControls.ControlCaptionPaddingVertical
    End Select

    Select Case RedrawControls.ControlCaptionAlignHorizontal
        Case 0: RedrawInLeft = RedrawBorderSize + RedrawControls.ControlCaptionPaddingHorizontal
        Case 1: RedrawInLeft = (RedrawPicture.ScaleWidth - RedrawIcon.ScaleWidth) / 2
        Case 2: RedrawInLeft = RedrawPicture.ScaleWidth - RedrawBorderSize - RedrawIcon.ScaleWidth - RedrawControls.ControlCaptionPaddingHorizontal
    End Select

    If RedrawControls.ControlCurrentPress = True Then
        RedrawInLeft = RedrawInLeft + 1
        RedrawInTop = RedrawInTop + 1
    End If

    RedrawPicture.PaintPicture RedrawIcon, RedrawInLeft, RedrawInTop, RedrawIcon.ScaleWidth, RedrawIcon.ScaleHeight

End If

If RedrawControls.ControlButtonSize > 1 Then
    
    RedrawOutRight = RedrawPicture.ScaleWidth - 1
    RedrawOutBottom = RedrawPicture.ScaleHeight - 1
    RedrawInBottom = RedrawOutBottom - 1
    RedrawInRight = RedrawOutRight - 1
    
    If RedrawControls.ControlCurrentPress = True Then
        RedrawBorderColor = &H646464
        RedrawButtonColor = &H646464
    Else
        RedrawBorderColor = &HFFFFFF
        RedrawButtonColor = &H696969
    End If
    
    RedrawPicture.Line (0, 0)-(RedrawOutRight, 0), RedrawBorderColor
    RedrawPicture.Line (0, 0)-(0, RedrawOutBottom), RedrawBorderColor
    RedrawPicture.Line (0, RedrawOutBottom)-(RedrawOutRight, RedrawOutBottom), RedrawButtonColor
    RedrawPicture.Line (RedrawOutRight, 0)-(RedrawOutRight, RedrawOutBottom + 1), RedrawButtonColor
        
    If RedrawControls.ControlCurrentPress = True Then
        RedrawBorderColor = &HA0A0A0
        RedrawButtonColor = &HA0A0A0
    Else
        RedrawBorderColor = &HE3E3E3
        RedrawButtonColor = &HA0A0A0
    End If
    
    RedrawPicture.Line (1, 1)-(RedrawInRight, 1), RedrawBorderColor
    RedrawPicture.Line (1, 1)-(1, RedrawInBottom), RedrawBorderColor
    RedrawPicture.Line (1, RedrawInBottom)-(RedrawInRight, RedrawInBottom), RedrawButtonColor
    RedrawPicture.Line (RedrawInRight, 1)-(RedrawInRight, RedrawInBottom + 1), RedrawButtonColor
        
End If

RedrawPicture.FillStyle = 1

RedrawPicture.AutoRedraw = True
RedrawPicture.Refresh

End Sub

Private Sub RedrawRounded(RedrawPicture As PictureBox, RedrawRadius As Integer, RedrawInLeft As Integer, RedrawInRight As Integer, RedrawInTop As Integer, RedrawInBottom As Integer, RedrawOutLeft As Integer, RedrawOutRight As Integer, RedrawOutTop As Integer, RedrawOutBottom As Integer, RedrawSize As Integer, RedrawColor As OLE_COLOR)

RedrawPicture.FillColor = RedrawColor
RedrawPicture.ForeColor = RedrawColor

If RedrawRadius > 0 Then
    RedrawPicture.Circle (RedrawInRight, RedrawInTop), RedrawRadius, , -0.000001, -1.571429
    RedrawPicture.Circle (RedrawInLeft, RedrawInTop), RedrawRadius, , -1.571429, -3.142857
    RedrawPicture.Circle (RedrawInLeft, RedrawInBottom), RedrawRadius, , -3.142857, -4.714286
    RedrawPicture.Circle (RedrawInRight, RedrawInBottom), RedrawRadius, , -4.714286, -0.000001
End If

RedrawPicture.Line (RedrawInLeft, RedrawOutTop)-(RedrawInRight, RedrawOutBottom), RedrawColor, BF
RedrawPicture.Line (RedrawOutLeft, RedrawInTop)-(RedrawOutRight, RedrawInBottom), RedrawColor, BF
        
RedrawRadius = RedrawRadius - RedrawSize
RedrawOutRight = RedrawOutRight - RedrawSize
RedrawOutBottom = RedrawOutBottom - RedrawSize
RedrawOutLeft = RedrawOutLeft + RedrawSize
RedrawOutTop = RedrawOutTop + RedrawSize

End Sub

Private Sub RedrawText(RedrawControls As GILControls, RedrawPicture As PictureBox, RedrawCaption As String, RedrawRowIndex As Integer, RedrawTotalLimit As Integer, RedrawBorderSize As Integer, RedrawHeight As Integer, RedrawIcon As PictureBox)

Dim RedrawOffset As Integer

If RedrawControls.ControlCurrentPress = True Then
    RedrawOffset = 1
Else
    RedrawOffset = 0
End If

Select Case RedrawControls.ControlCaptionAlignHorizontal
    Case 0: RedrawPicture.CurrentX = RedrawBorderSize + RedrawControls.ControlCaptionPaddingHorizontal + RedrawOffset
    Case 1: RedrawPicture.CurrentX = (RedrawPicture.ScaleWidth - RedrawPicture.TextWidth(RedrawCaption)) / 2 + RedrawOffset
    Case 2: RedrawPicture.CurrentX = RedrawPicture.ScaleWidth - RedrawBorderSize - RedrawControls.ControlCaptionPaddingHorizontal - RedrawPicture.TextWidth(RedrawCaption) + RedrawOffset
End Select

If RedrawControls.ControlIconValid = True Then

    Select Case RedrawControls.ControlCaptionAlignVertical
        Case 0: RedrawPicture.CurrentY = RedrawBorderSize + RedrawControls.ControlCaptionPaddingVertical * RedrawRowIndex + RedrawHeight * (RedrawRowIndex - 1) + RedrawOffset + RedrawControls.ControlIconPadding + RedrawIcon.ScaleHeight + RedrawControls.ControlCaptionPaddingVertical
        Case 1: RedrawPicture.CurrentY = (RedrawPicture.ScaleHeight - RedrawHeight * RedrawTotalLimit - RedrawControls.ControlCaptionPaddingVertical * (RedrawTotalLimit - 1) + RedrawIcon.ScaleHeight + RedrawControls.ControlCaptionPaddingVertical) / 2 + (RedrawControls.ControlCaptionPaddingVertical + RedrawHeight) * (RedrawRowIndex - 1) + RedrawOffset
        Case 2: RedrawPicture.CurrentY = RedrawPicture.ScaleHeight - RedrawBorderSize - (RedrawControls.ControlCaptionPaddingVertical + RedrawHeight) * (RedrawTotalLimit - RedrawRowIndex + 1) + RedrawOffset - RedrawControls.ControlIconPadding
    End Select

Else

    Select Case RedrawControls.ControlCaptionAlignVertical
        Case 0: RedrawPicture.CurrentY = RedrawBorderSize + RedrawControls.ControlCaptionPaddingVertical * RedrawRowIndex + RedrawHeight * (RedrawRowIndex - 1) + RedrawOffset + RedrawControls.ControlIconPadding
        Case 1: RedrawPicture.CurrentY = (RedrawPicture.ScaleHeight - RedrawHeight * RedrawTotalLimit - RedrawControls.ControlCaptionPaddingVertical * (RedrawTotalLimit - 1)) / 2 + (RedrawControls.ControlCaptionPaddingVertical + RedrawHeight) * (RedrawRowIndex - 1) + RedrawOffset
        Case 2: RedrawPicture.CurrentY = RedrawPicture.ScaleHeight - RedrawBorderSize - (RedrawControls.ControlCaptionPaddingVertical + RedrawHeight) * (RedrawTotalLimit - RedrawRowIndex + 1) + RedrawOffset - RedrawControls.ControlIconPadding
    End Select

End If

RedrawPicture.Print RedrawCaption

End Sub

Private Sub SplitCaption(SplitControls As GILControls, SplitPicture As PictureBox, SplitWidth As Integer)

Dim SplitLine As String
Dim SplitText As String
Dim SplitCheck As String
Dim SplitRowIndex As Integer
Dim SplitRowTotal As Integer
Dim SplitWords() As String
Dim SplitWordIndex As Integer
Dim SplitWordTotal As Integer
Dim SplitRows() As String

SplitText = SplitControls.ControlCaptionBytes
SplitText = Replace(Trim$(SplitText), "  ", " ")
SplitText = Replace(Trim$(SplitText), vbCrLf + " ", vbCrLf)
SplitText = Replace(SplitText, " " + vbCrLf, vbCrLf)
SplitText = Replace(SplitText, vbCrLf + vbCrLf, vbCrLf)
    
If SplitControls.ControlWordWrap = True Then

    SplitRows() = Split(SplitText, vbCrLf)
    
    SplitRowTotal = UBound(SplitRows) + 1
    SplitText = ""
    
    SplitRowIndex = 0
    While SplitRowIndex < SplitRowTotal
    
        SplitWords() = Split(SplitRows(SplitRowIndex), " ")
    
        SplitWordTotal = UBound(SplitWords) + 1
        SplitLine = ""
        
        SplitWordIndex = 0
        While SplitWordIndex < SplitWordTotal
        
            SplitCheck = Trim$(SplitLine + " " + SplitWords(SplitWordIndex))
            
            If SplitPicture.TextWidth(SplitCheck) > SplitWidth Then
                If SplitLine <> "" Then SplitText = SplitText + vbCrLf + SplitLine
                SplitLine = SplitWords(SplitWordIndex)
            Else
                SplitLine = SplitCheck
            End If
            
            SplitWordIndex = SplitWordIndex + 1
        Wend
        
        SplitText = SplitText + vbCrLf + SplitLine
        
        SplitRowIndex = SplitRowIndex + 1
    Wend

    SplitText = Mid$(SplitText, 1 + Len(vbCrLf))
    
End If

SplitControls.ControlCaptionRows() = Split(SplitText, vbCrLf)

End Sub

Private Function HandleToPicture(ByVal ConvertHandle As Long, ConvertBitmap As Boolean) As IPicture

Dim ConvertPicture As PICTDESC
Dim ConvertGUID(0 To 3) As Long

ConvertPicture.cbSize = Len(ConvertPicture)

If ConvertBitmap Then
    ConvertPicture.pictType = vbPicTypeBitmap
Else
    ConvertPicture.pictType = vbPicTypeIcon
End If

ConvertPicture.hIcon = ConvertHandle

ConvertGUID(0) = &H7BF80980
ConvertGUID(1) = &H101ABF32
ConvertGUID(2) = &HAA00BB8B
ConvertGUID(3) = &HAB0C3000

OleCreatePictureIndirect ConvertPicture, ConvertGUID(0), True, HandleToPicture

End Function

