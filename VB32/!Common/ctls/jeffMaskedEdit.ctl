VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl jeffMaskedEdit 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   DataBindingBehavior=   1  'vbSimpleBound
   ScaleHeight     =   960
   ScaleWidth      =   3090
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   2385
   End
End
Attribute VB_Name = "jeffMaskedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This code was written by The Frog Prince
'
' If you have questions or comments, I can be reached at
'        TheFrogPrince@hotmail.com
' If you wanna see more cool vb user controls, classes, code,
' and add-ins like this one, or updates to this code, go to
' my web page at
'        http://members.tripod.com/the__frog__prince/
' You are free to use, re-write, or otherwise do as you wish
' with this code.  However, if you do a cool enhancement, I
' would appreciate it if you could e-mail it to me.  I like
' to see what people do with my stuff.  =)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
' SPECIAL NOTES
'
'   In order to extend the validate event of the masked edit
'   box, I had to rename the both the event Validate to
'   Validation, and the CausesValidation property to
'   CausesValidate.  I have no clue why since the user control
'   has no means by which to turn on and off the validate
'   internally... it only shows up externally.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Option Explicit


'Default Property Values:
Dim m_AutoTab As Boolean

Const m_def_ClipText = ""
Const m_def_HideSelection = 0
Const m_def_SelLength = 0
Const m_def_SelStart = 0
Const m_def_SelText = "0"
Const m_def_FormattedText = ""
Const m_def_Text = ""
Const m_def_Locked = 0
'Property Variables:
Dim m_ClipText As String
Dim m_HideSelection As Boolean
Dim m_SelLength As Integer
Dim m_SelStart As Integer
Dim m_SelText As String
Dim m_FormattedText As String
Dim m_Text As String
Dim m_Locked As Boolean
Dim m_AllowedKeys As jmeAllowedKeys
Dim m_AutoSelectAll As Boolean

Private bHasFocus As Boolean


'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "OLEDragOver event"
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "OLEDragDrop event"
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "OLEGiveFeedback event"
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "OLEStartDrag event"
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLESetData
Attribute OLESetData.VB_Description = "OLESetData event"
Event OLECompleteDrag(Effect As Long) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "OLECompleteDrag event"
Event Change() 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Change
Event ValidationError(InvalidText As String, StartPosition As Integer) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,ValidationError
Event Validation(Cancel As Boolean) 'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Validate

' This is a public copy of the ctlKeyPress enum in mAPIconstants and must
' be synchronized.
Public Enum jmeAllowedKeys
    
    [jmeNoSpaces] = 2 ^ 0
    [jmeNoSingleQuotes] = 2 ^ 1
    [jmeNoDoubleQuotes] = 2 ^ 2
    
    [jmeUppercase] = 2 ^ 3
    [jmeLowercase] = 2 ^ 4
    
    [jmeAllowDecimal] = 2 ^ 5
    [jmeAllowNegative] = 2 ^ 6
    [jmeAllowSpaces] = 2 ^ 7
    [jmeAllowStars] = 2 ^ 8
    [jmeAllowPounds] = 2 ^ 9
    [jmeAllowForwardSlash] = 2 ^ 10
    [jmeAllowParenthesis] = 2 ^ 11
    [jmeAllowDollarSigns] = 2 ^ 12
    [jmeAllowColon] = 2 ^ 13
    [jmeAllowAMPM] = 2 ^ 14
    
    [jmeOnlyNUMBERS] = 2 ^ 15
    [jmeOnlyPHONE] = jmeOnlyNUMBERS + jmeAllowParenthesis + jmeAllowSpaces + jmeAllowNegative + jmeAllowPounds
    [jmeOnlyDATE] = jmeOnlyNUMBERS + jmeAllowForwardSlash
    [jmeOnlyMONEY] = jmeOnlyNUMBERS + jmeAllowDollarSigns + jmeAllowNegative + jmeAllowDecimal
    [jmeOnlyTIME] = jmeOnlyNUMBERS + jmeAllowColon + jmeAllowSpaces + jmeAllowAMPM
    [jmeOnlyNUMBERSwithDecimals] = jmeOnlyNUMBERS + jmeAllowDecimal
    
End Enum



Public Property Get AutoSelectAll() As Boolean
    AutoSelectAll = m_AutoSelectAll
End Property

Public Property Let AutoSelectAll(ByVal bAuto As Boolean)
    m_AutoSelectAll = bAuto
    PropertyChanged "AutoSelectAll"
End Property



Private Function KeyNotAllowed(ByVal KeyAscii As Integer, ByVal eAllowed As jmeAllowedKeys) As Boolean
    KeyNotAllowed = True
    Dim sChar As String
    sChar = Chr(KeyAscii)
    Select Case True
        Case (eAllowed And jmeNoDoubleQuotes) And sChar = """"
        Case (eAllowed And jmeNoSingleQuotes) And sChar = "'"
        Case (eAllowed And jmeNoSpaces) And sChar = " "
        Case Else
            KeyNotAllowed = False
    End Select
End Function

Private Function KeyAllowed(ByVal KeyAscii As Integer, ByVal eAllowed As jmeAllowedKeys) As Boolean
    KeyAllowed = True
    Dim sChar As String
    sChar = Chr(KeyAscii)
    Select Case True
        Case (eAllowed And jmeAllowAMPM) And InStr("AMP", UCase(sChar)) > 0
        Case (eAllowed And jmeAllowColon) And sChar = ":"
        Case (eAllowed And jmeAllowDecimal) And sChar = "." And InStr(txtUnselected(MaskEdBox1), ".") = 0
        Case (eAllowed And jmeAllowDollarSigns) And sChar = "$"
        Case (eAllowed And jmeAllowForwardSlash) And sChar = "/"
        Case (eAllowed And jmeAllowNegative) And sChar = "-"
        Case (eAllowed And jmeAllowParenthesis) And InStr("()", sChar) > 0
        Case (eAllowed And jmeAllowPounds) And sChar = "#"
        Case (eAllowed And jmeAllowSpaces) And sChar = " "
        Case (eAllowed And jmeAllowStars) And sChar = "*"
        Case Else
            KeyAllowed = False
    End Select
End Function

Private Function OnKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal eAllowed As jmeAllowedKeys) As Integer
    
    ' Allow All System Keys Through
    If KeyAscii < 32 Then
        
        ' AutoTab
        If m_AutoTab And KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
        End If
        
        OnKeyPress = KeyAscii
        Exit Function
    End If
    
    ' Uliminate In-Eligible Keystrokes
    Select Case True
        Case KeyAllowed(KeyAscii, eAllowed)
        Case KeyNotAllowed(KeyAscii, eAllowed)
            KeyAscii = 0
        Case (eAllowed And jmeOnlyNUMBERS) And (KeyAscii < vbKey0 Or KeyAscii > vbKey9)
            KeyAscii = 0
    End Select
    
    ' Coerce values
    Select Case True
        Case (eAllowed And jmeLowercase)
            KeyAscii = Asc(LCase(Chr(KeyAscii)))
        Case (eAllowed And jmeUppercase)
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
    
    OnKeyPress = KeyAscii
    
End Function


Public Property Get AllowedKeys() As jmeAllowedKeys
    AllowedKeys = m_AllowedKeys
End Property

Public Property Let AllowedKeys(ByVal eKeys As jmeAllowedKeys)
    m_AllowedKeys = eKeys
    PropertyChanged "AllowedKeys"
    
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = Text1.Alignment
    
End Property

Public Property Let Alignment(ByVal eAlignment As AlignmentConstants)
    Text1.Alignment = eAlignment
    PropertyChanged "Alignment"
End Property



Private Sub MaskEdBox1_GotFocus()
    bHasFocus = True
    HideMaskedEdit False
    If m_AutoSelectAll Then
        MaskEdBox1.SelStart = 0
        MaskEdBox1.SelLength = Len(MaskEdBox1.FormattedText)
    End If
    
End Sub

Private Sub MaskEdBox1_LostFocus()
    bHasFocus = False
    ApplyMask
    HideMaskedEdit True

End Sub

Public Function HideMaskedEdit(bHide As Boolean)
    If bHide Then
        MaskEdBox1.Move -MaskEdBox1.Width, -MaskEdBox1.Height
    Else
        MaskEdBox1.Move 0, 0
    End If
End Function


Private Sub MaskEdBox1_OLEDragDrop(Data As MSMask.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub MaskEdBox1_OLEDragOver(Data As MSMask.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub MaskEdBox1_OLEStartDrag(Data As MSMask.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub Text1_GotFocus()
    bHasFocus = True
    HideMaskedEdit False
    ts.ctlSetFocus MaskEdBox1
    
End Sub

Private Sub Text1_LostFocus()
    bHasFocus = False
    
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_GotFocus()
    bHasFocus = True
    HideMaskedEdit False
    ts.ctlSetFocus MaskEdBox1
    
End Sub

Private Sub UserControl_Initialize()
    HideMaskedEdit True
    m_AutoSelectAll = True
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SelLength = m_def_SelLength
    m_SelStart = m_def_SelStart
    m_SelText = m_def_SelText
    m_FormattedText = m_def_FormattedText
    m_Text = m_def_Text
    m_ClipText = m_def_ClipText
    m_HideSelection = m_def_HideSelection
        
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    HideMaskedEdit True
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Me.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Me.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Me.Appearance = PropBag.ReadProperty("Appearance", 1)
    Me.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Me.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Me.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Me.ClipMode = PropBag.ReadProperty("ClipMode", 0)
    Me.PromptInclude = PropBag.ReadProperty("PromptInclude", True)
    Me.AllowPrompt = PropBag.ReadProperty("AllowPrompt", False)
    Me.AutoTab = PropBag.ReadProperty("AutoTab", False)
    Me.MaxLength = PropBag.ReadProperty("MaxLength", 64)
    Me.Format = PropBag.ReadProperty("Format", "")
    Me.Mask = PropBag.ReadProperty("Mask", "")
    Me.PromptChar = PropBag.ReadProperty("PromptChar", "_")
    Me.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    Me.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    MaskEdBox1.CausesValidation = PropBag.ReadProperty("CausesValidate", True)
    Me.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
'    m_HideSelection = PropBag.ReadProperty("HideSelection", m_def_HideSelection)
    Me.SelLength = PropBag.ReadProperty("SelLength", m_def_SelLength)
    Me.SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
    Me.SelText = PropBag.ReadProperty("SelText", m_def_SelText)
    m_FormattedText = PropBag.ReadProperty("FormattedText", m_def_FormattedText)
    Me.Text = PropBag.ReadProperty("Text", m_def_Text)
    m_ClipText = PropBag.ReadProperty("ClipText", m_def_ClipText)
'    MaskEdBox1.HideSelection = PropBag.ReadProperty("HideSelection", m_def_HideSelection)
    m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
    Me.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_AllowedKeys = PropBag.ReadProperty("AllowedKeys", 0)
    m_AutoSelectAll = PropBag.ReadProperty("AutoSelectAll", -1)
    
End Sub

Private Sub UserControl_Resize()
    UserControl.MaskEdBox1.Move UserControl.MaskEdBox1.Left, UserControl.MaskEdBox1.Top, UserControl.Width, UserControl.Height
    UserControl.Text1.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", MaskEdBox1.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", MaskEdBox1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", MaskEdBox1.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", MaskEdBox1.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Font", MaskEdBox1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", MaskEdBox1.Appearance, 1)
    Call PropBag.WriteProperty("OLEDropMode", MaskEdBox1.OLEDropMode, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", MaskEdBox1.BorderStyle, 1)
    Call PropBag.WriteProperty("ClipMode", MaskEdBox1.ClipMode, 0)
    Call PropBag.WriteProperty("PromptInclude", MaskEdBox1.PromptInclude, True)
    Call PropBag.WriteProperty("AllowPrompt", MaskEdBox1.AllowPrompt, False)
    Call PropBag.WriteProperty("AutoTab", MaskEdBox1.AutoTab, False)
    Call PropBag.WriteProperty("MaxLength", MaskEdBox1.MaxLength, 64)
    Call PropBag.WriteProperty("Format", MaskEdBox1.Format, "")
    Call PropBag.WriteProperty("Mask", MaskEdBox1.Mask, "")
    Call PropBag.WriteProperty("PromptChar", MaskEdBox1.PromptChar, "_")
    Call PropBag.WriteProperty("OLEDragMode", MaskEdBox1.OLEDragMode, 0)
    Call PropBag.WriteProperty("WhatsThisHelpID", MaskEdBox1.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("CausesValidate", MaskEdBox1.CausesValidation, True)
    Call PropBag.WriteProperty("ToolTipText", MaskEdBox1.ToolTipText, "")
    Call PropBag.WriteProperty("SelLength", MaskEdBox1.SelLength, m_def_SelLength)
    Call PropBag.WriteProperty("SelStart", MaskEdBox1.SelStart, m_def_SelStart)
    Call PropBag.WriteProperty("SelText", MaskEdBox1.SelText, m_def_SelText)
    Call PropBag.WriteProperty("FormattedText", m_FormattedText, m_def_FormattedText)
    Call PropBag.WriteProperty("Text", MaskEdBox1.Text, m_def_Text)
    Call PropBag.WriteProperty("ClipText", m_ClipText, m_def_ClipText)
    Call PropBag.WriteProperty("HideSelection", MaskEdBox1.HideSelection, m_def_HideSelection)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("AllowedKeys", m_AllowedKeys, 0)
    Call PropBag.WriteProperty("AutoSelectAll", m_AutoSelectAll, -1)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = MaskEdBox1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    MaskEdBox1.BackColor() = New_BackColor
    Text1.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = MaskEdBox1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    MaskEdBox1.ForeColor() = New_ForeColor
    Text1.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = MaskEdBox1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MaskEdBox1.Enabled() = New_Enabled
    Text1.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the three-dimensional style of the check box caption."
    MousePointer = MaskEdBox1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    MaskEdBox1.MousePointer() = New_MousePointer
    Text1.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    Set MouseIcon = MaskEdBox1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set MaskEdBox1.MouseIcon = New_MouseIcon
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font Property"
Attribute Font.VB_UserMemId = -512
    Set Font = MaskEdBox1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MaskEdBox1.Font = New_Font
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Sets whether the control has a flat or sunken 3d appearance"
    Appearance = MaskEdBox1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    MaskEdBox1.Appearance() = New_Appearance
    Text1.Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this control can act as an OLE drop target."
    OLEDropMode = MaskEdBox1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    MaskEdBox1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,BorderStyle
Public Property Get BorderStyle() As MSMask.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = MaskEdBox1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSMask.BorderStyleConstants)
    MaskEdBox1.BorderStyle() = New_BorderStyle
    Text1.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    MaskEdBox1.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    MaskEdBox1.OLEDrag
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
    
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    'msgbox "KeyDown: " & KeyCode
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    'msgbox "KeyPress: " & KeyAscii
    If m_Locked Then
        KeyAscii = 0
    End If
    KeyAscii = OnKeyPress(KeyAscii, m_AllowedKeys)
    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
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



Private Sub MaskEdBox1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub


Private Sub MaskEdBox1_OLESetData(Data As MSMask.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub MaskEdBox1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,ClipMode
Public Property Get ClipMode() As ClipModeConstants
Attribute ClipMode.VB_Description = "Determines whether to include or exclude the literal characters in the input mask when doing a cut or copy command."
    ClipMode = MaskEdBox1.ClipMode
End Property

Public Property Let ClipMode(ByVal New_ClipMode As ClipModeConstants)
    MaskEdBox1.ClipMode() = New_ClipMode
    PropertyChanged "ClipMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,PromptInclude
Public Property Get PromptInclude() As Boolean
Attribute PromptInclude.VB_Description = "Specifies whether prompt characters are contained in the Text property value."
    PromptInclude = MaskEdBox1.PromptInclude
End Property

Public Property Let PromptInclude(ByVal New_PromptInclude As Boolean)
    MaskEdBox1.PromptInclude() = New_PromptInclude
    PropertyChanged "PromptInclude"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,AllowPrompt
Public Property Get AllowPrompt() As Boolean
Attribute AllowPrompt.VB_Description = "Determines whether or not the prompt character is a valid input character."
    AllowPrompt = MaskEdBox1.AllowPrompt
End Property

Public Property Let AllowPrompt(ByVal New_AllowPrompt As Boolean)
    MaskEdBox1.AllowPrompt() = New_AllowPrompt
    
    PropertyChanged "AllowPrompt"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,AutoTab
Public Property Get AutoTab() As Boolean
Attribute AutoTab.VB_Description = "Determines whether or not the next control in the tab order receives the focus."
    AutoTab = MaskEdBox1.AutoTab
End Property

Public Property Let AutoTab(ByVal New_AutoTab As Boolean)
    MaskEdBox1.AutoTab() = New_AutoTab
    m_AutoTab = New_AutoTab
    PropertyChanged "AutoTab"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,MaxLength
Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Sets/returns the maximum length of the masked edit control."
    MaxLength = MaskEdBox1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Integer)
    MaskEdBox1.MaxLength() = New_MaxLength
    Text1.MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Format
Public Property Get Format() As String
Attribute Format.VB_Description = "Specifies the format for displaying and printing numbers, dates, times, and text."
    Format = MaskEdBox1.Format
End Property

Public Property Let Format(ByVal New_Format As String)
    MaskEdBox1.Format = New_Format
    ApplyMask
    PropertyChanged "Format"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Mask
Public Property Get Mask() As String
Attribute Mask.VB_Description = "Determines the input mask for the control."
    Mask = MaskEdBox1.Mask
End Property

Public Property Let Mask(ByVal New_Mask As String)
    MaskEdBox1.Mask = New_Mask
    PropertyChanged "Mask"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,PromptChar
Public Property Get PromptChar() As String
Attribute PromptChar.VB_Description = "Sets/returns the character used to prompt a user for input."
    PromptChar = MaskEdBox1.PromptChar
End Property

Public Property Let PromptChar(ByVal New_PromptChar As String)
    MaskEdBox1.PromptChar() = New_PromptChar
    PropertyChanged "PromptChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = MaskEdBox1.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    MaskEdBox1.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = MaskEdBox1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    MaskEdBox1.WhatsThisHelpID() = New_WhatsThisHelpID
    Text1.WhatsThisHelpID = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,CausesValidation
Public Property Get CausesValidate() As Boolean
    CausesValidate = MaskEdBox1.CausesValidation
End Property

Public Property Let CausesValidate(ByVal New_CausesValidation As Boolean)
    MaskEdBox1.CausesValidation = New_CausesValidation
    PropertyChanged "CausesValidate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = MaskEdBox1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    MaskEdBox1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Function ApplyMask()
    On Error Resume Next
    If Len(MaskEdBox1.FormattedText) > Text1.MaxLength And Text1.MaxLength > 0 Then
        Me.MaxLength = Len(MaskEdBox1.FormattedText)
    End If
    Text1.Text = MaskEdBox1.FormattedText
    
End Function

Private Sub MaskEdBox1_Change()
    ApplyMask
    If bHasFocus Then
        MaskEdBox1.DataChanged = True
        Extender.DataChanged = True
        RaiseEvent Change
    End If
End Sub

Private Sub MaskEdBox1_ValidationError(InvalidText As String, StartPosition As Integer)
    RaiseEvent ValidationError(InvalidText, StartPosition)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    
    RaiseEvent Validation(Cancel)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SelLength() As Integer
    SelLength = MaskEdBox1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Integer)
    MaskEdBox1.SelLength = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SelStart() As Integer
    SelStart = MaskEdBox1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Integer)
    MaskEdBox1.SelStart = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get SelText() As String
    SelText = MaskEdBox1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    MaskEdBox1.SelText = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get FormattedText() As String
    FormattedText = MaskEdBox1.FormattedText
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Text() As Variant
Attribute Text.VB_MemberFlags = "34"
    Text = MaskEdBox1.Text
    If Text = "" Then
        Text = Empty
    End If
End Property

Public Property Let Text(ByVal New_Text As Variant)
    On Error Resume Next
    
    If Trim(NoNull(New_Text)) = "" Then
        If MaskEdBox1.Mask <> "" Then
            MaskEdBox1.Text = Replace(MaskEdBox1.Mask, "#", MaskEdBox1.PromptChar)
        Else
            MaskEdBox1.Text = Empty
        End If
    Else
        MaskEdBox1.Text = NoNull(New_Text)
    End If

'    If Err.Number = 0 Then
    If bHasFocus Then
        UserControl.PropertyChanged "Text"
    End If
'    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ClipText() As String
    ClipText = MaskEdBox1.ClipText
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HideSelection() As Boolean
    HideSelection = MaskEdBox1.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    MaskEdBox1.HideSelection = New_HideSelection
'    Text1.HideSelection = New_HideSelection
    PropertyChanged "HideSelection"
End Property


Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal new_LockValue As Boolean)
    m_Locked = new_LockValue
    Text1.Locked = new_LockValue
    PropertyChanged "Locked"
End Property

Public Property Let Value(ByVal new_Value As Variant)
    
    On Error Resume Next
    MaskEdBox1.Text = new_Value
    PropertyChanged "Value"
    
End Property

Public Property Get Value()
    Value = MaskEdBox1.Text
End Property


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' NoNull
'    NoNull is a wrapper to put around field values coming
'    out of a database when you cannot have them be null.
Private Function NoNull(ByRef vValue As Variant) As Variant
    If Not IsNull(vValue) Then
        NoNull = vValue
        Exit Function
    End If
    NoNull = Empty
End Function

