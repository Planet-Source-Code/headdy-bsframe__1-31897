VERSION 5.00
Begin VB.UserControl bsFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   MaskColor       =   &H00808000&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "bsFrame.ctx":0000
End
Attribute VB_Name = "bsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------
' BadSoft bsFrame control
' ©2002 BadSoft Entertainment, all rights reserved.
'
' Starting date 14/02/2002
'-------------------------------------

' Okay, I'm sitting here eating a whole chocolate trifle, and
' BadSoft needs to make a comeback. Yes in case you didn't
' realise, I've been away trying to sort my "organisation" out,
' and now is the time for a comeback. We're gonna bounce back
' into this hot s*** and liven up the mailing list a little bit.

' The bsFrame control will be designed to provide more
' options than Microsoft's intrinsic one. It'll have support for
' border styles, colours, frame styles and disabled text, and in
' the future text alignment, text/header positioning and maybe
' even images. But the goal is to make it as lightweight as
' possible.

' Let's see if I can remember how to do this...!

'Default Property Values:
Const m_def_FrameStyle = 0
Const m_def_CaptionColour = vbButtonText
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_FlatBorderColour = vbBlack
Const m_def_BackColour = vbButtonFace
Const m_def_BorderStyle = 6

'Property Variables:
Dim m_FlatBorderColour As OLE_COLOR
Dim m_FrameStyle As bsfFrameType
Dim m_CaptionColour As OLE_COLOR
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_BackColour As OLE_COLOR
Dim m_Caption As String
Dim m_BorderStyle As bsfBorderStyle


' ENUMERATIONS
' ----------------------------
Enum bsfCaptionAlignment
   bsTopLeft
   bsTopMiddle
   bsTopRight
   bsLeftTop
   bsLeftMiddle
   bsLeftBottom
   bsRightTop
   bsRightMiddle
   bsRightBottom
   bsBottomLeft
   bsBottomMiddle
   bsBottomRight
End Enum

Public Enum bsfBorderStyle
   bsfNone
   bsfFlat
   bsfRaisedThin
   bsfRaised3D
   bsfSunkenThin
   bsfSunken3D
   bsfEtched
   bsfBump
End Enum

Public Enum bsfFrameType
   bsfStandardFrame
   bsfPlainFrame
   bsfHeaderFrame
End Enum

' BackColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColour() As OLE_COLOR
Attribute BackColour.VB_Description = "The colour of the bsFrame's background."
   BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As OLE_COLOR)
   m_BackColour = New_BackColour
   PropertyChanged "BackColour"
   DrawFrame
End Property

' Caption()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text at the top of the bsFrame."
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawFrame
End Property

' BorderStyle()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get BorderStyle() As bsfBorderStyle
Attribute BorderStyle.VB_Description = "The style in which the bsFrame's edges are drawn."
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsfBorderStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawFrame
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
   DrawFrame
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Caption = UserControl.Extender.Name
   m_BorderStyle = m_def_BorderStyle
   m_FrameStyle = m_def_FrameStyle
   OleTranslateColor m_def_FlatBorderColour, 0, m_FlatBorderColour
   OleTranslateColor m_def_HighlightColour, 0, m_HighlightColour
   OleTranslateColor m_def_HighlightDKColour, 0, m_HighlightDKColour
   OleTranslateColor m_def_ShadowColour, 0, m_ShadowColour
   OleTranslateColor m_def_ShadowDKColour, 0, m_ShadowDKColour
   OleTranslateColor m_def_CaptionColour, 0, m_CaptionColour
   OleTranslateColor m_def_BackColour, 0, m_BackColour
   Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_BackColour = PropBag.ReadProperty("BackColour", m_def_BackColour)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDKColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDKColour", m_def_ShadowDKColour)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   Set UserControl.Font = PropBag.ReadProperty("Fount", Ambient.Font)
   m_FrameStyle = PropBag.ReadProperty("FrameStyle", m_def_FrameStyle)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
   DrawFrame
End Sub

Private Sub UserControl_Show()
   DrawFrame
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColour", m_BackColour, m_def_BackColour)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDKColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDKColour", m_ShadowDKColour, m_def_ShadowDKColour)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("Fount", UserControl.Font, Ambient.Font)
   Call PropBag.WriteProperty("FrameStyle", m_FrameStyle, m_def_FrameStyle)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub


' DoPlainEdges()
' --------------------------
Private Sub DoPlainEdges()

   Dim lPen As Long
   
   Select Case m_BorderStyle
      Case bsfFlat
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_FlatBorderColour))
         SelectObject UserControl.hdc, lPen
         Rectangle UserControl.hdc, 0, 0, _
            UserControl.ScaleWidth, UserControl.ScaleHeight
         DeleteObject lPen
         
      Case bsfRaisedThin
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         
      Case bsfSunkenThin
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         
      Case bsfRaised3D
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen
      
      Case bsfSunken3D
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, m_HighlightColour)
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen
         
      Case bsfEtched
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen

         
      Case bsfBump
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen

      End Select
End Sub

' DoHeaderEdges()
' --------------------------
Private Sub DoHeaderEdges(ByVal TextHeight As Long)

   Dim lPen As Long
   
   Select Case BorderStyle
      Case bsfFlat
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_FlatBorderColour))
         SelectObject UserControl.hdc, lPen
         Rectangle UserControl.hdc, 0, 0, _
            UserControl.ScaleWidth, TextHeight + 4
         'body
         Rectangle UserControl.hdc, 0, TextHeight + 3, _
            UserControl.ScaleWidth, UserControl.ScaleHeight
            
         DeleteObject lPen
         
      Case bsfRaisedThin
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, ScaleWidth - 1, _
            TextHeight + 4, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         DeleteObject lPen

      Case bsfSunkenThin
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, ScaleWidth - 1, _
            TextHeight + 4, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         DeleteObject lPen
         
      Case bsfRaised3D
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, 1, 0
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 2, 1
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            TextHeight + 5, 0
         LineTo UserControl.hdc, 0, TextHeight + 5
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 5
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            TextHeight + 6, 0
         LineTo UserControl.hdc, 1, TextHeight + 6
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 6
         DeleteObject lPen
      
      Case bsfSunken3D
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, 1, 0
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 2, 1
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            TextHeight + 5, 0
         LineTo UserControl.hdc, 0, TextHeight + 5
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 5
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            TextHeight + 6, 0
         LineTo UserControl.hdc, 1, TextHeight + 6
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 6
         DeleteObject lPen
      
      Case bsfEtched
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         LineTo UserControl.hdc, ScaleWidth - 1, -1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, 1, 0
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 2, 0
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            TextHeight + 5, 0
         LineTo UserControl.hdc, 0, TextHeight + 5
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 5
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            TextHeight + 6, 0
         LineTo UserControl.hdc, 1, TextHeight + 6
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 5
         DeleteObject lPen

      Case bsfBump
         'header
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, 0, 0
         LineTo UserControl.hdc, 0, 0
         LineTo UserControl.hdc, 0, TextHeight + 4
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 4
         LineTo UserControl.hdc, ScaleWidth - 1, 0
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, 1, 0
         LineTo UserControl.hdc, 1, 1
         LineTo UserControl.hdc, 1, TextHeight + 3
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 3
         LineTo UserControl.hdc, ScaleWidth - 2, 1
         DeleteObject lPen
         'body
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, _
            TextHeight + 5, 0
         LineTo UserControl.hdc, 0, TextHeight + 5
         LineTo UserControl.hdc, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, TextHeight + 5
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, _
            TextHeight + 6, 0
         LineTo UserControl.hdc, 1, TextHeight + 6
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, TextHeight + 6
         DeleteObject lPen
         
   End Select
End Sub

' DoStandardEdges()
' --------------------------
Private Sub DoStandardEdges(ByVal TextHeight As Long)

   Dim lPen As Long, lBrush As Long, rectTemp As RECT
   Dim iMargin As Integer
   
   iMargin = TextHeight / 2
   
   Select Case BorderStyle
      Case bsfFlat
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_FlatBorderColour))
         SelectObject UserControl.hdc, lPen
         Rectangle UserControl.hdc, 0, iMargin, _
            UserControl.ScaleWidth, UserControl.ScaleHeight
         DeleteObject lPen
         
      Case bsfRaisedThin
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen

      Case bsfSunkenThin
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         
      Case bsfRaised3D
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen
      
      Case bsfSunken3D
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowDKColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightDKColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen
         
      Case bsfEtched
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin
         DeleteObject lPen

      Case bsfBump
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 1, iMargin, 0
         LineTo UserControl.hdc, 0, iMargin
         LineTo UserControl.hdc, 0, UserControl.ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hdc, ScaleWidth - 1, iMargin
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_ShadowColour))
         SelectObject UserControl.hdc, lPen
         MoveToEx UserControl.hdc, UserControl.ScaleWidth - 2, iMargin + 1, 0
         LineTo UserControl.hdc, 1, iMargin + 1
         LineTo UserControl.hdc, 1, UserControl.ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(PS_SOLID, 1, _
            TranslateColor(m_HighlightColour))
         SelectObject UserControl.hdc, lPen
         LineTo UserControl.hdc, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hdc, ScaleWidth - 2, iMargin + 1
         DeleteObject lPen

      End Select
End Sub

' DrawFrame()
' --------------------------
Private Sub DrawFrame()

   Dim rectCaption As RECT, rectTemp As RECT
   Dim strCaption As String
   Dim iCaptionHeight As Long
   Dim realX As Integer, realY As Integer
   
   'Clear everything
   lBrush = CreateSolidBrush(TranslateColor(BackColour))
   SetRect rectTemp, 0, 0, ScaleWidth, ScaleHeight
   FillRect UserControl.hdc, rectTemp, lBrush
   DeleteObject lBrush
   
   'Before we draw anything we need to calculate the size space
   'the text will take up. This is so that the side that holds
   'the text can be drawn properly.
   'We use a separate string for checking the size of the
   'caption because if it's too long we can truncate it with an
   'ellipsis (...).
   
   strCaption = m_Caption
   rectCaption.Right = UserControl.ScaleWidth
   rectCaption.Left = 8
   rectCaption.Bottom = 2
  
   'Calling DrawText with the DT_CALCRECT flag doesn't draw any
   'text.
   Call DrawText(UserControl.hdc, strCaption, _
      Len(strCaption), rectCaption, DT_CALCRECT)
   
   Select Case m_FrameStyle
      Case bsfStandardFrame
         DoStandardEdges rectCaption.Bottom
            
         'rectangle behind text
         lBrush = CreateSolidBrush(TranslateColor(BackColour))
         With rectCaption
            SetRect rectTemp, 6, 0, .Right + 2, .Bottom
         End With
         FillRect UserControl.hdc, rectTemp, lBrush
         DeleteObject lBrush
      
      Case bsfPlainFrame
         DoPlainEdges
         rectCaption.Left = 4
         rectCaption.Top = 2
         
      Case bsfHeaderFrame
         DoHeaderEdges rectCaption.Bottom
         rectCaption.Left = 4
         rectCaption.Top = 2
   End Select
         
   'draw text
   rectCaption.Bottom = rectCaption.Bottom + 2
   
   If UserControl.Enabled = False And _
      UserControl.Ambient.UserMode = True Then
   'We don't need a special API function, here's how to draw
   'disabled text.
      With rectCaption
         SetRect rectCaption, .Left + 1, .Top + 1, .Right + 1, _
            .Bottom + 1
         SetTextColor UserControl.hdc, _
            TranslateColor(m_HighlightColour)
         Call DrawText(UserControl.hdc, strCaption, _
            Len(strCaption), rectCaption, 0)
            
         SetRect rectCaption, .Left - 1, .Top - 1, .Right - 1, _
            .Bottom - 1
         SetTextColor UserControl.hdc, _
            TranslateColor(m_ShadowColour)
         Call DrawText(UserControl.hdc, strCaption, _
            Len(strCaption), rectCaption, 0)
      End With
   Else
   'Draw text as normal.
      SetTextColor UserControl.hdc, _
         TranslateColor(m_CaptionColour)
      Call DrawText(UserControl.hdc, strCaption, _
         Len(strCaption), rectCaption, 0)
   End If
   
   'Very important...
   Refresh
End Sub

' FlatBorderColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FlatBorderColour() As OLE_COLOR
Attribute FlatBorderColour.VB_Description = "The colour of the edges when BorderStyle is set to bsfFlat."

   FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
   m_FlatBorderColour = New_FlatBorderColour
   PropertyChanged "FlatBorderColour"
   DrawFrame
End Property

' HighlightColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightColour() As OLE_COLOR
Attribute HighlightColour.VB_Description = "The colour of the lightest border colour."
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawFrame
End Property

' HighlightDKColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightDKColour() As OLE_COLOR
Attribute HighlightDKColour.VB_Description = "The colour of the second lightest border colour."
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDKColour"
   DrawFrame
End Property

' ShadowColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowColour() As OLE_COLOR
Attribute ShadowColour.VB_Description = "The second darkest border colour."
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawFrame
End Property

' ShadowDKColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowDKColour() As OLE_COLOR
Attribute ShadowDKColour.VB_Description = "The darkest border colour."
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDKColour"
   DrawFrame
End Property

' CaptionColour()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbButtonText
Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the bsFrame's Caption text."
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawFrame
End Property

' Fount()
' --------------------------
' Before you say anything, Fount is the English word for Font.
' Hence Font is American. Because I'm British, I use British
' words. Unlike many of you I haven't sold myself out.

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Fount() As Font
Attribute Fount.VB_Description = "The font used for the Caption. (Fount is the English word for font.)"
   Set Fount = UserControl.Font
End Property

Public Property Set Fount(ByVal New_Fount As Font)
   Set UserControl.Font = New_Fount
   PropertyChanged "Fount"
   DrawFrame
End Property

' FrameStyle()
' --------------------------
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get FrameStyle() As bsfFrameType
Attribute FrameStyle.VB_Description = "The design of the bsFrame."
   FrameStyle = m_FrameStyle
End Property

Public Property Let FrameStyle(ByVal New_FrameStyle As bsfFrameType)
   m_FrameStyle = New_FrameStyle
   PropertyChanged "FrameStyle"
   DrawFrame
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Whether or not the control can respond to the user's actions."
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
   DrawFrame
End Property

Private Sub ShowAbout()
   frmAbout.Show vbModal
End Sub

Public Sub About()
Attribute About.VB_Description = "Shows information about this control."
Attribute About.VB_UserMemId = -552
   ShowAbout
End Sub
