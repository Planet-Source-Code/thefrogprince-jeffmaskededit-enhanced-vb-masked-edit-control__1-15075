VERSION 5.00
Begin VB.Form frmTestFrame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Harness For jeffFrame"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmTestFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin ControlTestHarness.jeffFrame frmExample 
      Height          =   1185
      Left            =   510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3900
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   2090
      BorderStyle     =   6
      Begin VB.PictureBox picJEFF 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   3450
         Picture         =   "frmTestFrame.frx":27A2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblLabel1 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the example frame.  Modify the properties above to see the variations you can create."
         Height          =   765
         Left            =   330
         TabIndex        =   13
         Top             =   270
         Width           =   2925
      End
   End
   Begin ControlTestHarness.jeffFrame frmParams 
      Height          =   2355
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4154
      BorderStyle     =   2
      BorderWidth     =   5
      Begin VB.VScrollBar scrollBorderWidth 
         Height          =   345
         Left            =   1590
         Max             =   0
         Min             =   -50
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Value           =   -1
         Width           =   225
      End
      Begin VB.ListBox listBorderEdges 
         Height          =   1755
         IntegralHeight  =   0   'False
         Left            =   2550
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   420
         Width           =   2355
      End
      Begin VB.ListBox listBorderStyles 
         Height          =   1410
         IntegralHeight  =   0   'False
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   780
         Width           =   2325
      End
      Begin VB.TextBox txtBorderWidth 
         Height          =   345
         Left            =   1200
         TabIndex        =   5
         Text            =   "1"
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblBorderEdges 
         AutoSize        =   -1  'True
         Caption         =   "Border &Edges"
         Height          =   195
         Left            =   2580
         TabIndex        =   9
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblBorderStyles 
         AutoSize        =   -1  'True
         Caption         =   "Border &Styles"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   930
      End
      Begin VB.Label lblBorderWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Border &Width:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   975
      End
   End
   Begin ControlTestHarness.jeffFrame frmTitle 
      Height          =   615
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1085
      BorderStyle     =   6
      BorderWidth     =   5
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Public Frame Test Harness"
         Top             =   150
         Width           =   2895
      End
   End
   Begin ControlTestHarness.jeffFrame frmDivider 
      Height          =   105
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3180
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   185
      BorderStyle     =   6
      BorderWidth     =   2
      BorderEdge      =   2
   End
   Begin VB.Label lblExample 
      AutoSize        =   -1  'True
      Caption         =   "Example Frame"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3510
      Width           =   1080
   End
End
Attribute VB_Name = "frmTestFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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



Private Sub Form_Load()
    Dim check As enumBorderEdges
    
    With Me.listBorderStyles
        .Clear
        .AddItem "BDR_RAISEDINNER"
        .ItemData(.NewIndex) = BDR_RAISEDINNER
        .AddItem "BDR_RAISEDOUTER"
        .ItemData(.NewIndex) = BDR_RAISEDOUTER
        .AddItem "BDR_SUNKENINNER"
        .ItemData(.NewIndex) = BDR_SUNKENINNER
        .AddItem "BDR_SUNKENOUTER"
        .ItemData(.NewIndex) = BDR_SUNKENOUTER
        .AddItem "EDGE_BUMP"
        .ItemData(.NewIndex) = EDGE_BUMP
        .AddItem "EDGE_ETCHED"
        .ItemData(.NewIndex) = EDGE_ETCHED
        .Selected(.NewIndex) = True
        .AddItem "EDGE_RAISED"
        .ItemData(.NewIndex) = EDGE_RAISED
        .AddItem "EDGE_SUNKEN"
        .ItemData(.NewIndex) = EDGE_SUNKEN
        
    End With
    
    With Me.listBorderEdges
        .Clear
        .AddItem "BF_ADJUST"
        .ItemData(.NewIndex) = BF_ADJUST
        .AddItem "BF_BOTTOM"
        .ItemData(.NewIndex) = BF_BOTTOM
        .AddItem "BF_BOTTOMLEFT"
        .ItemData(.NewIndex) = BF_BOTTOMLEFT
        .AddItem "BF_BOTTOMRIGHT"
        .ItemData(.NewIndex) = BF_BOTTOMRIGHT
        .AddItem "BF_DIAGONAL"
        .ItemData(.NewIndex) = BF_DIAGONAL
        .AddItem "BF_DIAGONAL_ENDBOTTOMLEFT"
        .ItemData(.NewIndex) = BF_DIAGONAL_ENDBOTTOMLEFT
        .AddItem "BF_DIAGONAL_ENDBOTTOMRIGHT"
        .ItemData(.NewIndex) = BF_DIAGONAL_ENDBOTTOMRIGHT
        .AddItem "BF_DIAGONAL_ENDTOPLEFT"
        .ItemData(.NewIndex) = BF_DIAGONAL_ENDTOPLEFT
        .AddItem "BF_DIAGONAL_ENDTOPRIGHT"
        .ItemData(.NewIndex) = BF_DIAGONAL_ENDTOPRIGHT
        .AddItem "BF_FLAT"
        .ItemData(.NewIndex) = BF_FLAT
        .AddItem "BF_LEFT"
        .ItemData(.NewIndex) = BF_LEFT
        .AddItem "BF_MIDDLE"
        .ItemData(.NewIndex) = BF_MIDDLE
        .AddItem "BF_MONO"
        .ItemData(.NewIndex) = BF_MONO
        .AddItem "BF_RECT"
        .ItemData(.NewIndex) = BF_RECT
        .Selected(.NewIndex) = True
        .AddItem "BF_RIGHT"
        .ItemData(.NewIndex) = BF_RIGHT
        .AddItem "BF_SOFT"
        .ItemData(.NewIndex) = BF_SOFT
        .AddItem "BF_TOP"
        .ItemData(.NewIndex) = BF_TOP
        .AddItem "BF_TOPLEFT"
        .ItemData(.NewIndex) = BF_TOPLEFT
        .AddItem "BF_TOPRIGHT"
        .ItemData(.NewIndex) = BF_TOPRIGHT
    End With
    
End Sub

Private Sub listBorderEdges_Click()
    Me.ApplyProperties
End Sub

Private Sub listBorderStyles_Click()
    Me.ApplyProperties
End Sub

Private Sub scrollBorderWidth_Change()
    Me.txtBorderWidth.Text = Abs(Me.scrollBorderWidth.Value)
    Me.ApplyProperties
End Sub

Private Sub txtBorderWidth_Change()
    If Val(Me.txtBorderWidth) > 50 Then
        Me.txtBorderWidth.Text = 50
    End If
    Me.scrollBorderWidth.Value = -1 * Me.txtBorderWidth.Text
    Me.ApplyProperties
End Sub

Private Sub txtBorderWidth_KeyPress(KeyAscii As Integer)
    ts.ctlKeyPress KeyAscii, NumbersOnly
End Sub


Public Function ApplyProperties()
    Dim l As Long
    Dim e As Long
    With Me.listBorderStyles
        For l = 0 To .ListCount - 1
            If .Selected(l) Then
                e = e Or .ItemData(l)
            End If
        Next l
    End With
    Me.frmExample.BorderStyle = e
    Me.frmExample.BorderWidth = Val(Me.txtBorderWidth.Text)
    e = 0
    With Me.listBorderEdges
        For l = 0 To .ListCount - 1
            If .Selected(l) Then
                e = e Or .ItemData(l)
            End If
        Next l
    End With
    Me.frmExample.BorderEdge = e
End Function
