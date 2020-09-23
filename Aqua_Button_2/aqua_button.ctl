VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   ScaleHeight     =   405
   ScaleWidth      =   1605
   ToolboxBitmap   =   "aqua_button.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   510
      TabIndex        =   0
      Top             =   90
      Width           =   480
   End
   Begin VB.Image picOff 
      Height          =   420
      Left            =   -15
      Picture         =   "aqua_button.ctx":0312
      Top             =   -15
      Width           =   1650
   End
   Begin VB.Image picOn 
      Height          =   540
      Left            =   -60
      Picture         =   "aqua_button.ctx":056C
      Top             =   -60
      Width           =   1785
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Click() 'MappingInfo=Label1,Label1,-1,Click
'Event Declarations:
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
'Default Property Values:
Const m_def_FontTransparent = 0
'Property Variables:
Dim m_FontTransparent As Boolean





Private Sub picOn_Click()
    RaiseEvent Click
End Sub
Private Sub Label1_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    MsgBox "HitTest"
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picOn.Visible = True
    picOff.Visible = False
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picOn.Visible = False
    picOff.Visible = True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
Attribute ActiveControl.VB_Description = "Returns the control that has focus."
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Label1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    'Label1.FontName = PropBag.ReadProperty("FontName", "")
    Label1.FontSize = PropBag.ReadProperty("FontSize", 8)
    Label1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
    Label1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Label1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Label1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Label1.FontName, "")
    Call PropBag.WriteProperty("FontSize", Label1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", Label1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", m_FontTransparent, m_def_FontTransparent)
    Call PropBag.WriteProperty("FontUnderline", Label1.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Label1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Label1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Label1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = m_FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    m_FontTransparent = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Label1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Label1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FontTransparent = m_def_FontTransparent
End Sub

