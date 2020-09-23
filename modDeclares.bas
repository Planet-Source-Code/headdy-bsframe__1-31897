Attribute VB_Name = "modDeclares"
Option Explicit

' The difference between a pen nd a brush - a pen is used for
' drawing borders, while a brush is used for filling areas.
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

' Rectangle
' An API call for drawing rectangles. The style of the rectangle is decided by the form's
' (or picture object's, or any object with an hDC) drawing styles, such as FillColor,
' ForeColor and FillStyle. One of the most useful API calls, but why didn't they just use
' a RECT variable?
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    
' DrawText
' An API call for drawing text. It uses the ForeColor property
' of the object to determine the colour of the text. To draw
' disabled text you will need to use the DrawStateString API
' call.
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
    (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, _
    lpRect As RECT, ByVal uFormat As Long) As Long

Public Const DT_WORD_ELLIPSIS = &H40000
'gets rid of extra words at the end and truncates the last word.

Public Const DT_MODIFYSTRING = &H10000
'The function is allowed to modify the passed string to fit the
'text inside the rectangle. It has no effect if DT_WORD_ELLIPSIS
'is not also specified, and also only seems to work on single
'lines of text (using the DT_SINGLELINE flag).

Public Const DT_CALCRECT = &H400
'Calculates the exact space the text takes up inside the RECT
'variable. (Works like the Autosize property of a Label
'control.)
                               
Public Const DT_SINGLELINE = &H20   'All the text will be placed on a single line.

Public Const DT_VCENTER = &H4        'Text is centered vertically in the RECT
                                                        'variable.
                 
Public Const DT_CENTER = &H1          'Centers the text horizontally within the
                                                        'rectangle.
Public Const DT_LEFT = &H0               'Left aligns the text (by default).
Public Const DT_RIGHT = &H2            'Right aligns the text.
Public Const DT_WORDBREAK = &H10    'Allows for more than one line of text if
                                                        'necessary.

' SetRect
' An API call for quickly setting the coordinates of a RECT variable.
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, _
   ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


' OleTranslateColor
' ----------------------------
' If you select a system colour in an ActiveX property that is
' of the OLE_COLOR type, this API call will help you to turn
' it into a long (recognisable) value. This'll make things a
' hell of a lot easier for you.
Declare Function OleTranslateColor Lib "oleaut32.dll" _
   (ByVal lOleColor As Long, ByVal lHPalette As Long, _
   lColorRef As Long) As Long

Private Const CLR_INVALID = -1

' The RECT type is very common amongst API calls, and is used
' to define a rectangle.
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


' MoveToEx
' ----------------------------
' This changes the drawing point on the device context. If you
' want you can store the old coordinates in lpPoint.
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
   ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long


' LineTo
' ----------------------------
' This draws a line from the current point on the hDC. The
' point can be changed by using MoveToEx if you want to stay
' with API. Note that this draws up to, but not including, the
' given coordinates. The new point is set to these
' coordinates.
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
   ByVal X As Long, ByVal Y As Long) As Long

Public Const Author = "Drew (aka The Bad One)"

' Pen constant (needed?)
Const PS_SOLID = 0


' DrawStateString
' ----------------------------
' This (hidden?) API call lets you draw text in a number of
' ways. For example, you can draw it "disabled" or selected.
' I say this call is hidden because I don't think it's used
' in many places and documented well. I got this from an
' example at vbaccelerator.com.
Declare Function DrawStateString Lib "user32" _
   Alias "DrawStateA" (ByVal hdc As Long, _
   ByVal hBrush As Long, ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, ByVal cbStringLen As Long, _
   ByVal X As Long, ByVal Y As Long, ByVal cX As Long, _
   ByVal cY As Long, ByVal fuFlags As Long) As Long

' Used flags for DrawStateString
Public Const DSS_DISABLED = &H20
Public Const DST_TEXT = &H1

Function TranslateColor(ByVal oClr As OLE_COLOR, _
   Optional hPal As Long = 0) As Long
   ' Convert Automation color to Windows color
   If OleTranslateColor(oClr, hPal, TranslateColor) Then
       TranslateColor = CLR_INVALID
   End If
End Function

