VERSION 5.00
Begin VB.UserControl EZFrame 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   HitBehavior     =   0  'None
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
End
Attribute VB_Name = "EZFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0

Private Const DEF_TextBoxHeight As Long = 16
Private Const DEF_TextColor As Long = vbBlack
Private Const DEF_TextBoxColor As Long = vbWhite
Private Const DEF_FrameColor As Long = vbBlack

Private p_FrameColor As OLE_COLOR
Private p_TextBoxColor As OLE_COLOR
Private p_BackColor As OLE_COLOR
Private p_Caption As String
Private p_TextBoxHeight As Long
Private p_TextColor As Long
Private p_Alignment As Long
Private p_Font As StdFont

Private ControlRight As Long
Private ControlBottom As Long
Private TextBoxCenter As Long
Private TextDrawParams As Long

Private lpp As POINTAPI

'TextBox RECT
Private tbRECT As RECT

'Text RECT
Private tRECT As RECT


Private Sub UserControl_Initialize()
    
    'Initialize font
    
    Set p_Font = New StdFont
    Set UserControl.Font = p_Font
    
End Sub

Private Sub UserControl_InitProperties()

    'Set default properties
    
    p_TextBoxHeight = DEF_TextBoxHeight&
    TextBoxCenter = DEF_TextBoxHeight \ 2&
    SetTextDrawParams
    p_TextBoxColor = DEF_TextBoxColor
    p_TextColor = DEF_TextColor
    p_FrameColor = DEF_FrameColor
    p_BackColor = Ambient.BackColor
    p_Caption = Ambient.DisplayName
    
End Sub

Private Sub UserControl_Paint()
    
    'Brush handle
    Dim bHandle As Long
    
    'Set frame color
    UserControl.ForeColor = p_FrameColor
    
    'Clear user control
    UserControl.Cls
    
    '''''''''''''''''''''''
    'Draw surrounding lines
    '''''''''''''''''''''''
    
    MoveToEx UserControl.hdc, 3&, TextBoxCenter, lpp
    LineTo UserControl.hdc, 0&, TextBoxCenter
    LineTo UserControl.hdc, 0&, ControlBottom
    LineTo UserControl.hdc, ControlRight, ControlBottom
    LineTo UserControl.hdc, ControlRight, TextBoxCenter
    LineTo UserControl.hdc, ControlRight - 4&, TextBoxCenter
    
    ''''''''''''''
    'Draw text box
    ''''''''''''''
    
    'Create solid brush, fill text box rect with the color and delete brush
    bHandle = CreateSolidBrush(p_TextBoxColor)
    FillRect UserControl.hdc, tbRECT, bHandle
    DeleteObject bHandle
    
    'Draw text box borders
    MoveToEx UserControl.hdc, 4&, 0&, lpp
    LineTo UserControl.hdc, 4&, p_TextBoxHeight
    LineTo UserControl.hdc, ControlRight - 4&, p_TextBoxHeight
    LineTo UserControl.hdc, ControlRight - 4&, 0&
    LineTo UserControl.hdc, 4&, 0&

    '''''''''''''
    'Draw caption
    '''''''''''''
    
    'Set text color
    UserControl.ForeColor = p_TextColor
    
    'Draw text
    DrawTextEx UserControl.hdc, p_Caption, Len(p_Caption), tRECT, TextDrawParams, ByVal 0&

End Sub

Private Sub UserControl_Resize()
    ControlRight = UserControl.ScaleWidth - 1&
    ControlBottom = UserControl.ScaleHeight - 1&
    SetTextBoxRect
    
    UserControl_Paint
End Sub

'Properties ->

Public Property Let FrameColor(ByRef new_FrameColor As OLE_COLOR)
    p_FrameColor = new_FrameColor
    PropertyChanged "FrameColor"
    UserControl_Paint
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = p_FrameColor
End Property

Public Property Let Caption(ByRef new_caption As String)
    p_Caption = new_caption
    UserControl_Paint
End Property

Public Property Get Caption() As String
    Caption = p_Caption
End Property

Public Property Let TextBoxHeight(ByRef new_TextBoxHeight As Long)
    p_TextBoxHeight = new_TextBoxHeight
    TextBoxCenter = p_TextBoxHeight \ 2
    SetTextBoxRect
    
    PropertyChanged "TextBoxHeight"
    UserControl_Paint
End Property

Public Property Get TextBoxHeight() As Long
    TextBoxHeight = p_TextBoxHeight
End Property

Public Property Let Alignment(ByRef new_Alignment As AlignmentConstants)
    p_Alignment = new_Alignment
    SetTextDrawParams
    
    PropertyChanged "Alignment"
    UserControl_Paint
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = p_Alignment
End Property

Public Property Let TextColor(ByRef new_TextColor As OLE_COLOR)
    p_TextColor = new_TextColor
    
    PropertyChanged "TextColor"
    UserControl_Paint
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = p_TextColor
End Property

Public Property Let TextBoxColor(ByRef new_TextBoxColor As OLE_COLOR)
    p_TextBoxColor = new_TextBoxColor
    
    PropertyChanged "TextBoxColor"
    UserControl_Paint
End Property

Public Property Get TextBoxColor() As OLE_COLOR
    TextBoxColor = p_TextBoxColor
End Property

Public Property Let BackColor(ByRef new_BackColor As OLE_COLOR)
    p_BackColor = new_BackColor
    UserControl.BackColor = p_BackColor
    
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = p_BackColor
End Property

Public Property Set Font(ByRef new_font As StdFont)
    SetFont new_font
    
    PropertyChanged "Font"
    UserControl_Paint
End Property

Public Property Let Font(ByRef new_font As StdFont)
    SetFont new_font
    
    PropertyChanged "Font"
    UserControl_Paint
End Property
Public Property Get Font() As StdFont
    Set Font = p_Font
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'Write properties
    'Occures when changing parameters or saving project etc.
    PropBag.WriteProperty "FrameColor", p_FrameColor, DEF_FrameColor
    PropBag.WriteProperty "BackColor", p_BackColor, Ambient.BackColor
    PropBag.WriteProperty "TextBoxColor", p_TextBoxColor, DEF_TextBoxColor
    PropBag.WriteProperty "Caption", p_Caption, ""
    PropBag.WriteProperty "TextBoxHeight", p_TextBoxHeight, DEF_TextBoxHeight
    PropBag.WriteProperty "TextColor", p_TextColor, DEF_TextColor
    PropBag.WriteProperty "Alignment", p_Alignment, vbLeftJustify
    PropBag.WriteProperty "Font", p_Font, Ambient.Font
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    'Read properties
    'Occures when loading form etc.
    p_FrameColor = PropBag.ReadProperty("FrameColor", DEF_FrameColor)
    p_BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    p_TextBoxColor = PropBag.ReadProperty("TextBoxColor", DEF_TextBoxColor)
    p_Caption = PropBag.ReadProperty("Caption", "")
    p_TextBoxHeight = PropBag.ReadProperty("TextBoxHeight", DEF_TextBoxHeight)
    p_TextColor = PropBag.ReadProperty("TextColor", DEF_TextColor)
    p_Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    Set p_Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    'Add properties
    UserControl.BackColor = p_BackColor
    TextBoxCenter = p_TextBoxHeight \ 2
    SetTextBoxRect
    SetTextDrawParams
    SetFont p_Font
    
    'Paint control
    UserControl_Paint
    
End Sub

'Other functions ->

Private Sub SetTextBoxRect()
    
    'Precalculating text rect and text box rect saves cpu a little bit
    
    'Set text box rect
    tbRECT.Top = 0&
    tbRECT.Left = 4&
    tbRECT.Right = ControlRight - 4&
    tbRECT.Bottom = p_TextBoxHeight
    
    'Set text rect
    tRECT.Top = 1&
    tRECT.Left = 7&
    tRECT.Right = ControlRight - 7&
    tRECT.Bottom = p_TextBoxHeight - 1&
    
End Sub

Private Sub SetTextDrawParams()
    
    'Set text draw params using p_Alignment
    
    If p_Alignment = vbLeftJustify Then
        TextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
    ElseIf p_Alignment = vbRightJustify Then
        TextDrawParams = DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
    Else:
        TextDrawParams = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    End If
    
End Sub

Private Sub SetFont(ByRef new_font As StdFont)
    
    'Set text font using new_font
    
    With p_Font
        .Bold = new_font.Bold
        .Italic = new_font.Italic
        .Name = new_font.Name
        .Size = new_font.Size
    End With
    Set UserControl.Font = p_Font

End Sub
