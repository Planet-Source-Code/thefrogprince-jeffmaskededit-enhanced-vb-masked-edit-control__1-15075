VERSION 5.00
Begin VB.UserControl jeffFrame 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ControlContainer=   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   6660
End
Attribute VB_Name = "jeffFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Option Explicit

Private mvarBorderStyle As enumBorderEdges
Private remoteRef As jeffFrame
Private mvarBorderWidth As Long
Private mvarBorderEdge As enumBorderFlags

Public Event Resize()

Public Property Get BorderEdge() As enumBorderFlags
    BorderEdge = mvarBorderEdge
End Property

Public Property Let BorderEdge(newEdge As enumBorderFlags)
    mvarBorderEdge = newEdge
    PropertyChanged "BorderEdge"
    Me.Refresh
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = mvarBorderWidth
End Property
Public Property Let BorderWidth(ByVal newWidth As Long)
    mvarBorderWidth = newWidth
    PropertyChanged "BorderWidth"
    Me.Refresh
    
End Property

Public Function GetRemoteReference()
    Dim l As Long
    For l = 1 To UserControl.Parent.Controls.Count
        If UserControl.Parent.Controls(l) Is Me Then
            Set remoteRef = UserControl.Parent.Controls(l)
            Exit For
        End If
    Next l
    
End Function

Private Sub UserControl_Initialize()
'    GetRemoteReference

End Sub

Private Sub UserControl_InitProperties()
    mvarBorderWidth = 1
    mvarBorderStyle = EDGE_ETCHED
    mvarBorderEdge = BF_RECT
    
End Sub

Private Sub UserControl_Paint()
    Static bPainting As Boolean
    If Not bPainting Then
        bPainting = True
            
        Dim tr As Rect
        tr.Left = 0
        tr.Top = 0
        tr.Right = (UserControl.ScaleWidth \ Screen.TwipsPerPixelX) - 0
        tr.Bottom = (UserControl.ScaleHeight \ Screen.TwipsPerPixelY) - 0
        UserControl.Cls
        
        Dim lOffset As Long
    '    lOffset = mvarBorderWidth - 1
    '    DrawEdge UserControl.hDC, rectMake(tr.Left, tr.Top, tr.Left + lOffset, tr.Top + lOffset), mvarBorderStyle, BF_TOPLEFT
    '    DrawEdge UserControl.hDC, rectMake(tr.Left + mvarBorderWidth, tr.Top, tr.Right - mvarBorderWidth, tr.Top + lOffset), mvarBorderStyle, BF_TOP
    '    DrawEdge UserControl.hDC, rectMake(tr.Right - lOffset, tr.Top, tr.Right, tr.Top + lOffset), mvarBorderStyle, BF_TOPRIGHT
        
        
        
        ' Paint BackGround
    '    Debug.Print UserControl.Parent.hDC
    '    Debug.Print UserControl.CurrentX & "," & UserControl.CurrentY
        'BitBlt UserControl.hDC, 1, 1, UserControl.Width - 2, UserControl.Height - 2, UserControl.Parent.hDC, UserControl.CurrentX, UserControl.CurrentY, 0
        'UserControl.Parent.Image.Render UserControl.hDC,  0,  0, UserControl.Width, UserControl.Height, UserControl.CurrentX, UserControl.CurrentY, UserControl.Width, UserControl.Height, Null
    '    UserControl.Picture = UserControl.Parent.Image
        
        'UserControl.Image.Render UserControl.hDC, 0, 0, UserControl.Width, UserControl.Height, UserControl.CurrentX, UserControl.CurrentY, UserControl.Width, UserControl.Height, 0
        
    '    UserControl.PaintPicture UserControl.Parent.Picture, UserControl.CurrentX, UserControl.CurrentY
        
        
        
        
        
        Dim l As Long
        For l = 1 To Int((mvarBorderWidth / 2) + 0.99)
            Dim lOffsetOut As Long, lOffsetIn As Long
            lOffsetOut = l - 1
            lOffsetIn = mvarBorderWidth - l
            DrawEdge UserControl.hdc, rectMake(tr.Left + lOffsetOut, tr.Top + lOffsetOut, tr.Right - lOffsetOut, tr.Bottom - lOffsetOut), mvarBorderStyle, mvarBorderEdge
            DrawEdge UserControl.hdc, rectMake(tr.Left + lOffsetIn, tr.Top + lOffsetIn, tr.Right - lOffsetIn, tr.Bottom - lOffsetIn), mvarBorderStyle, mvarBorderEdge
        Next l
            
    '    tr.Left = 1
    '    tr.Top = 1
    '    tr.Bottom = tr.Bottom - 1
    '    tr.Right = tr.Right - 1
    '    DrawEdge UserControl.hDC, tr, mvarBorderStyle, BF_RECT
    
        bPainting = False
    End If
    
End Sub

Public Function Refresh()
    UserControl.Refresh
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As enumBorderEdges
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = mvarBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumBorderEdges)
    mvarBorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Me.Refresh
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    mvarBorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    mvarBorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    mvarBorderEdge = PropBag.ReadProperty("BorderEdge", BF_RECT)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

Private Sub UserControl_Terminate()
'    Set remoteRef = Nothing
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", mvarBorderStyle, 0)
    Call PropBag.WriteProperty("BorderWidth", mvarBorderWidth, 1)
    Call PropBag.WriteProperty("BorderEdge", mvarBorderEdge, BF_RECT)
End Sub

