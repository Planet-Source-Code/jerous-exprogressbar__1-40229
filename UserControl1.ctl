VERSION 5.00
Begin VB.UserControl exProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ToolboxBitmap   =   "UserControl1.ctx":0000
End
Attribute VB_Name = "exProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_SegmentWidth = 0.7
Const m_def_SegmentSize = 0
'Const m_def_PercentColor = 0
Const m_def_ShowPercent = -1
Const m_def_RelativeScroll = 0
Const m_def_StartColor = 0
Const m_def_EndColor = 0
Const m_def_Value = 0
Const m_def_Min = 0
Const m_def_Max = 100
'Property Variables:
Dim m_SegmentWidth As Single
Dim m_SegmentSize As Single
'Dim m_PercentColor As OLE_COLOR
Dim m_ShowPercent As Boolean
Dim m_RelativeScroll As Boolean
Dim m_StartColor As OLE_COLOR
Dim m_EndColor As OLE_COLOR
Dim m_Value As Integer
Dim m_Min As Integer
Dim m_Max As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Enum e_Borderstyle
  Flat
  D3D
End Enum
Private Sub DrawBar()
Dim SKleur As Long, EKleur As Long, G1 As Single, R1 As Single, B1 As Single
Dim R As Single, G As Single, B As Single
Dim Wi As Single, M%
Cls
If m_Value = m_Min Then
  GoTo Tekst
End If
ScaleWidth = Abs(m_Min) + Abs(m_Max)
SKleur = OleColorToRGB(m_StartColor)
EKleur = OleColorToRGB(m_EndColor)
R1 = RValue(SKleur)
G1 = GValue(SKleur)
B1 = BValue(SKleur)
M = Abs(m_Max) + Abs(m_Min)
If m_RelativeScroll Then
  M = Abs(Min) + m_Value
End If
G = ((GValue(EKleur) - G1)) / (M)
B = ((BValue(EKleur) - B1)) / (M)
R = ((RValue(EKleur) - R1)) / (M)
If UserControl.Enabled = False Then
  R1 = 255 / 2
  G1 = R1
  B1 = R1
  G = 0
  B = 0
  R = 0
End If
For Wi = 0 To m_Value + Abs(m_Min) 'ScaleHeight - 1
  Line (Wi, 0)-(Wi + 0 + 1, ScaleHeight), RGB(R1, G1, B1), BF
  R1 = R1 + R
  G1 = G1 + G
  B1 = B1 + B
Next Wi
If m_SegmentSize <> 0 And m_SegmentWidth <> 0 Then
  For Wi = 0 To m_Value + Abs(m_Min) Step m_SegmentSize
    If Wi + m_SegmentSize > m_Value + Abs(m_Min) Then
      Line (Wi, 0)-(ScaleWidth, ScaleHeight), UserControl.BackColor, BF
      Exit For
    End If
    Line (Wi, 0)-(Wi + m_SegmentWidth, ScaleHeight), UserControl.BackColor, BF
    If Int(Wi / 20) = Wi / 20 Then DoEvents
  Next
End If
Tekst:
If m_ShowPercent Then
  If m_RelativeScroll Then
    CurrentX = m_Value / 2 - TextWidth("8") / 2
  Else
    CurrentX = ScaleWidth / 2 - TextWidth("8") / 2
  End If
  CurrentY = ScaleHeight / 2 - TextHeight("8") / 2
  Print Trim(Round(m_Value + Abs(m_Min) / (Abs(m_Min) + Abs(m_Max)), 0)) + "%"
End If
End Sub
Private Function RValue(RGB As Long) As Integer
RValue = Dec#(Right(Bin(CDbl(RGB), 32), 8))
End Function
Private Function BValue(RGB As Long) As Integer
BValue = Dec#(Mid(Bin(CDbl(RGB), 32), 9, 8))
End Function
Private Function GValue(RGB As Long) As Integer
GValue = Dec#(Mid(Bin(CDbl(RGB), 32), 17, 8))
End Function
Private Function OleColorToRGB(ByVal clr As OLE_COLOR) As Long
If clr And &H80000000 Then
  OleColorToRGB = GetSysColor(clr Xor &H80000000)
Else
  OleColorToRGB = clr
End If
End Function
Private Function Bin(Value#, Optional Lengte%) As String
Dim X&, Dummy#, Rest#, Y&, M&, Res$
Value = Abs(Value)
X = -1
While Dummy <= Value
  X = X + 1
  Dummy = 1 * 2 ^ X
Wend
M = X - 1
Rest = Value
For Y = M To 0 Step -1
  Rest = Rest - (1 * 2 ^ Y)
  If Rest >= 0 Then Res$ = "1"
  If Rest < 0 Then
    Rest = Rest + (1 * 2 ^ Y)
    Res$ = "0"
  End If
  Bin = Bin + Res$
Next
If Bin = "" Then Bin = "0"
If Lengte = 0 Then
  Bin = Bin
Else
  Bin = String(Lengte - Len(Bin), "0") + Bin
End If
End Function
Private Function Dec#(Value$)
Dim X%, B%
For X = 1 To Len(Value)
  B% = Val(Mid(Value, X, 1))
  Dec = Dec + B * 2 ^ (Len(Value) - X)
Next
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As e_Borderstyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As e_Borderstyle)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
  UserControl.MousePointer() = New_MousePointer
  PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
  Set UserControl.MouseIcon = New_MouseIcon
  PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
  UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get StartColor() As OLE_COLOR
  StartColor = m_StartColor
End Property

Public Property Let StartColor(ByVal New_StartColor As OLE_COLOR)
  m_StartColor = New_StartColor
  PropertyChanged "StartColor"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get EndColor() As OLE_COLOR
  EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal New_EndColor As OLE_COLOR)
  m_EndColor = New_EndColor
  PropertyChanged "EndColor"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
  If New_Value > m_Max Then
    MsgBox "Value has to be less than Max.", vbExclamation
    Exit Property
  End If
  If New_Value < m_Min Then
    MsgBox "Value has to be bigger than Min.", vbExclamation
    Exit Property
  End If
  m_Value = New_Value
  PropertyChanged "Value"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
  Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
  If New_Min >= m_Max Then
    MsgBox "Min has to be less than Max.", vbExclamation
    Exit Property
  End If
  m_Min = New_Min
  PropertyChanged "Min"
  If m_Value < m_Min Then m_Value = m_Min
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Max() As Integer
  Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
  If New_Max <= m_Min Then
    MsgBox "Max has to be bigger than Min.", vbExclamation
    Exit Property
  End If
  m_Max = New_Max
  PropertyChanged "Max"
  If m_Value > m_Max Then m_Value = m_Max
  DrawBar
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_StartColor = m_def_StartColor
  m_EndColor = m_def_EndColor
  m_Value = m_def_Value
  m_Min = m_def_Min
  m_Max = m_def_Max
  m_ShowPercent = m_def_ShowPercent
  m_RelativeScroll = m_def_RelativeScroll
'  m_PercentColor = m_def_PercentColor
  m_SegmentSize = m_def_SegmentSize
  m_SegmentWidth = m_def_SegmentWidth
End Sub

Private Sub UserControl_Paint()
DrawBar
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  m_StartColor = PropBag.ReadProperty("StartColor", m_def_StartColor)
  m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  m_Min = PropBag.ReadProperty("Min", m_def_Min)
  m_Max = PropBag.ReadProperty("Max", m_def_Max)
  m_ShowPercent = PropBag.ReadProperty("ShowPercent", m_def_ShowPercent)
  m_RelativeScroll = PropBag.ReadProperty("RelativeScroll", m_def_RelativeScroll)
  m_SegmentSize = PropBag.ReadProperty("SegmentSize", m_def_SegmentSize)
  m_SegmentWidth = PropBag.ReadProperty("SegmentWidth", m_def_SegmentWidth)
  UserControl.ForeColor = PropBag.ReadProperty("PercentColor", &H80000012)
'If Ambient.UserMode = False Then
  DrawBar
'End If
End Sub

Private Sub UserControl_Resize()
DrawBar
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("StartColor", m_StartColor, m_def_StartColor)
  Call PropBag.WriteProperty("EndColor", m_EndColor, m_def_EndColor)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
  Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
  Call PropBag.WriteProperty("ShowPercent", m_ShowPercent, m_def_ShowPercent)
  Call PropBag.WriteProperty("RelativeScroll", m_RelativeScroll, m_def_RelativeScroll)
  Call PropBag.WriteProperty("SegmentSize", m_SegmentSize, m_def_SegmentSize)
  Call PropBag.WriteProperty("SegmentWidth", m_SegmentWidth, m_def_SegmentWidth)
  Call PropBag.WriteProperty("PercentColor", UserControl.ForeColor, &H80000012)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,-1
Public Property Get ShowPercent() As Boolean
  ShowPercent = m_ShowPercent
End Property

Public Property Let ShowPercent(ByVal New_ShowPercent As Boolean)
  m_ShowPercent = New_ShowPercent
  PropertyChanged "ShowPercent"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RelativeScroll() As Boolean
  RelativeScroll = m_RelativeScroll
End Property

Public Property Let RelativeScroll(ByVal New_RelativeScroll As Boolean)
  m_RelativeScroll = New_RelativeScroll
  PropertyChanged "RelativeScroll"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SegmentSize() As Single
  SegmentSize = m_SegmentSize
End Property

Public Property Let SegmentSize(ByVal New_SegmentSize As Single)
  If New_SegmentSize < m_SegmentWidth Then
    MsgBox "Segmentsize has to be more than Segmentwidth!", vbExclamation
    Exit Property
  End If
  m_SegmentSize = Abs(New_SegmentSize)
  PropertyChanged "SegmentSize"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,.7
Public Property Get SegmentWidth() As Single
  SegmentWidth = m_SegmentWidth
End Property

Public Property Let SegmentWidth(ByVal New_SegmentWidth As Single)
  If New_SegmentWidth > m_SegmentSize Then
    MsgBox "SegmentWidth has to be less than Segmentsize!", vbExclamation
    Exit Property
  End If
  m_SegmentWidth = Abs(New_SegmentWidth)
  PropertyChanged "SegmentWidth"
  DrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,2,0
Public Property Get Percent() As Integer
Attribute Percent.VB_MemberFlags = "400"
  Percent = Int(m_Value + Abs(m_Min)) / (Abs(m_Min) + Abs(m_Max)) * 100
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get PercentColor() As OLE_COLOR
Attribute PercentColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  PercentColor = UserControl.ForeColor
End Property

Public Property Let PercentColor(ByVal New_PercentColor As OLE_COLOR)
  UserControl.ForeColor() = New_PercentColor
  PropertyChanged "PercentColor"
  DrawBar
End Property

