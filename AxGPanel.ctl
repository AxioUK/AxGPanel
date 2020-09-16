VERSION 5.00
Begin VB.UserControl AxGPanel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ToolboxBitmap   =   "AxGPanel.ctx":0000
End
Attribute VB_Name = "AxGPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As PenAlignment) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
'-
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

Public Enum CrossPos
    cTopRight
    cMiddleRight
    cBottomRight
    cTopLeft
    cMiddleLeft
    cBottomLeft
    cMiddleTop
    cMiddleBottom
End Enum

Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeAntiAlias As Long = 4
Private Const LOGPIXELSX As Long = 88
Private Const CLR_INVALID = -1
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_Appearance = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Color1 = &HD59B5B
Const m_def_Color2 = &H6A5444
Const m_def_Angulo = 0

'Property Variables:
Dim GdipToken As Long
Dim nScale    As Single

Dim m_BorderColor   As OLE_COLOR
Dim m_ForeColor     As OLE_COLOR
Dim m_Color1        As OLE_COLOR
Dim m_Color2        As OLE_COLOR
Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Font          As Font
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim YCrossPos       As Long
Dim XCrossPos       As Long
Dim m_CrossPosition As CrossPos
Dim m_CrossVisible  As Boolean
Dim m_Moveable      As Boolean

'Event Declarations:
Public Event Click()
Public Event CrossClick()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Sub Refresh()
  UserControl.Cls
  Draw
  DrawCross
End Sub

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Sub Draw()
Dim REC As RECTL
Dim hBrush As Long, hGraphics As Long, hPen As Long
Dim Rgn As Long
    
'UserControl.ScaleMode = 1

With REC
    .Left = 0
    .Top = 0
  If m_BorderWidth <> 0 Then
    .Width = (UserControl.ScaleWidth * nScale) - 2 '((UserControl.ScaleWidth / Screen.TwipsPerPixelX) * nScale) - 2
    .Height = (UserControl.ScaleHeight * nScale) - 2  '((UserControl.ScaleHeight / Screen.TwipsPerPixelY) * nScale) - 2
   Else
    .Width = (UserControl.ScaleWidth * nScale)  '((UserControl.ScaleWidth / Screen.TwipsPerPixelX) * nScale)
    .Height = (UserControl.ScaleHeight * nScale)  '((UserControl.ScaleHeight / Screen.TwipsPerPixelY) * nScale)
   End If
End With

gRoundRect UserControl.hdc, REC, ConvertColor(m_Color1, 90), ConvertColor(m_Color2, 90), m_Angulo, ConvertColor(m_BorderColor, 90), m_CornerCurve
'UserControl.ScaleMode = 3
Rgn = CreateRoundRectRgn(0, 0, (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY), m_CornerCurve, m_CornerCurve)
SetWindowRgn UserControl.hwnd, Rgn, True
DeleteObject Rgn
'UserControl.ScaleMode = 1
UserControl.Refresh

End Sub


Private Sub DrawCross()
'Cross
  If m_CrossVisible Then
  
   'UserControl.ScaleMode = 3
   UserControl.DrawWidth = 2
   UserControl.ForeColor = m_BorderColor

    Select Case m_CrossPosition
      Case Is = cTopRight
            YCrossPos = BorderWidth + 10
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
      Case Is = cMiddleRight
            YCrossPos = (UserControl.ScaleHeight / 2) - 6
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
      Case Is = cBottomRight
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = UserControl.ScaleWidth - (BorderWidth + 15)
      Case Is = cTopLeft
            YCrossPos = BorderWidth + 10
            XCrossPos = (BorderWidth + 10)
      Case Is = cMiddleLeft
            YCrossPos = (UserControl.ScaleHeight / 2) - 6
            XCrossPos = (BorderWidth + 10)
      Case Is = cBottomLeft
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = (BorderWidth + 10)
      Case Is = cMiddleTop
            YCrossPos = BorderWidth + 10
            XCrossPos = (UserControl.ScaleWidth / 2)
      Case Is = cMiddleBottom
            YCrossPos = UserControl.ScaleHeight - (BorderWidth + 14)
            XCrossPos = (UserControl.ScaleWidth / 2)
    End Select

   UserControl.Line (XCrossPos, YCrossPos)-(XCrossPos + 6, YCrossPos + 6)
   UserControl.Line (XCrossPos, YCrossPos + 6)-(XCrossPos + 6, YCrossPos)
   'UserControl.ScaleMode = 1
  End If
  
End Sub
  
Private Sub gRoundRect(ByVal hdc As Long, RECT As RECTL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Angulo As Single, ByVal BorderColor As Long, Round As Long)
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    If m_BorderWidth <> 0 Then
      GdipCreatePen1 BorderColor, m_BorderWidth * nScale, &H2, hPen  '&H1 * nScale, &H2, hPen
    End If
    GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        If Round = 0 Then
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
            GdipAddPathLineI mPath, .Left, .Top, .Width, .Top       'Line-Top
            GdipAddPathLineI mPath, .Width, .Top, .Width, .Height   'Line-Left
            GdipAddPathLineI mPath, .Width, .Height, .Left, .Height 'Line-Bottom
            GdipAddPathLineI mPath, .Left, .Height, .Left, .Top     'Line-Right
        Else
        
            GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
            GdipAddPathArcI mPath, .Left + .Width - Round, .Top, Round, Round, 270, 90
            GdipAddPathArcI mPath, .Left + .Width - Round, .Top + .Height - Round, Round, Round, 0, 90
            GdipAddPathArcI mPath, .Left, .Top + .Height - Round, Round, Round, 90, 90
        End If
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics

End Sub

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  m_ForeColor = m_def_ForeColor
  m_Enabled = m_def_Enabled
  Set m_Font = Ambient.Font
  m_Color1 = m_def_Color1
  m_Color2 = m_def_Color2
  m_Angulo = m_def_Angulo
  m_BorderWidth = 1
  m_CornerCurve = 0
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'UserControl.ScaleMode = 3
If Button = vbLeftButton And m_Moveable = True Then
  Dim res As Long
  Call ReleaseCapture
  res = SendMessage(UserControl.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X > XCrossPos And X < XCrossPos + 6 And Y > (YCrossPos) And Y < (YCrossPos) + 6 Then
    If m_CrossVisible Then
      RaiseEvent CrossClick
      Extender.Visible = False
    End If
End If

RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
  m_Color1 = .ReadProperty("Color1", m_def_Color1)
  m_Color2 = .ReadProperty("Color2", m_def_Color2)
  m_Angulo = .ReadProperty("Angulo", m_def_Angulo)
  m_BorderColor = .ReadProperty("BorderColor", &HC0&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 0)
  m_CrossPosition = .ReadProperty("CrossPosition", cTopRight)
  m_CrossVisible = .ReadProperty("CrossVisible", False)
  m_Moveable = .ReadProperty("Moveable", False)
End With
Refresh
End Sub

Private Sub UserControl_Resize()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call .WriteProperty("Color1", m_Color1, m_def_Color1)
  Call .WriteProperty("Color2", m_Color2, m_def_Color2)
  Call .WriteProperty("Angulo", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  Call .WriteProperty("CrossPosition", m_CrossPosition, cTopRight)
  Call .WriteProperty("CrossVisible", m_CrossVisible, False)
  Call .WriteProperty("Moveable", m_Moveable, False)
End With
End Sub

Public Property Get Angulo() As Single
  Angulo = m_Angulo
End Property

Public Property Let Angulo(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "Angulo"
  Refresh
End Property

'Properties-------------------
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get Color1() As OLE_COLOR
  Color1 = m_Color1
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
  m_Color1 = New_Color1
  PropertyChanged "Color1"
  Refresh
End Property

Public Property Get Color2() As OLE_COLOR
  Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
  m_Color2 = New_Color2
  PropertyChanged "Color2"
  Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

'm_CrossVisible
Public Property Get CrossVisible() As Boolean
    CrossVisible = m_CrossVisible
End Property

Public Property Let CrossVisible(ByVal New_Value As Boolean)
    m_CrossVisible = New_Value
    PropertyChanged "CrossVisible"
    Refresh
End Property

'm_CrossPosition
Public Property Get CrossPosition() As CrossPos
    CrossPosition = m_CrossPosition
End Property

Public Property Let CrossPosition(ByVal New_Value As CrossPos)
    m_CrossPosition = New_Value
    PropertyChanged "CrossPosition"
    Refresh
End Property

Public Property Get Moveable() As Boolean
    Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
    m_Moveable = New_Moveable
    PropertyChanged "Moveable"
End Property

