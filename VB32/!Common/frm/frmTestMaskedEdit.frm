VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTestMaskedEdit 
   Caption         =   "Test Harness For jeffMaskedEdit"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   Icon            =   "frmTestMaskedEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin ControlTestHarness.jeffMaskedEdit meDBExample 
      DataField       =   "BoundField1"
      DataSource      =   "dcTest"
      Height          =   345
      Left            =   780
      TabIndex        =   5
      Top             =   1860
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      MouseIcon       =   "frmTestMaskedEdit.frx":27A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   8
      Mask            =   "?#######"
      SelText         =   ""
      Text            =   "________"
      HideSelection   =   -1  'True
      AllowedKeys     =   8
   End
   Begin MSAdodcLib.Adodc dcTest 
      Height          =   345
      Left            =   180
      Top             =   1470
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbTest.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbTest.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblMaskedEdit"
      Caption         =   "DB Example"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frmExamples 
      Caption         =   " Visual Examples "
      Height          =   1185
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2925
      Begin ControlTestHarness.jeffMaskedEdit meAMPM 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         MouseIcon       =   "frmTestMaskedEdit.frx":27BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   8
         Mask            =   "##:## ??"
         SelText         =   ""
         Text            =   "__:__ __"
         HideSelection   =   -1  'True
         Alignment       =   2
         AllowedKeys     =   57472
      End
      Begin ControlTestHarness.jeffMaskedEdit meMoney 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         MouseIcon       =   "frmTestMaskedEdit.frx":27DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,0.00"
         SelText         =   ""
         HideSelection   =   -1  'True
         Alignment       =   1
         AllowedKeys     =   36960
      End
      Begin ControlTestHarness.jeffMaskedEdit meDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   300
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         MouseIcon       =   "frmTestMaskedEdit.frx":27F6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Mask            =   "##/##/####"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Alignment       =   1
         AllowedKeys     =   33792
      End
      Begin ControlTestHarness.jeffMaskedEdit mePhone 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         MouseIcon       =   "frmTestMaskedEdit.frx":2812
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Mask            =   "(###) ###-####"
         SelText         =   ""
         Text            =   "(___) ___-____"
         HideSelection   =   -1  'True
         AllowedKeys     =   35520
      End
   End
   Begin VB.ListBox listAllowedKeys 
      Height          =   4785
      Left            =   3270
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   300
      Width           =   1755
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4170
      Width           =   1365
   End
   Begin VB.TextBox txtMask 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Top             =   3570
      Width           =   1515
   End
   Begin VB.TextBox txtFormat 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   3570
      Width           =   1485
   End
   Begin ControlTestHarness.jeffMaskedEdit meMainExample 
      Height          =   315
      Left            =   480
      TabIndex        =   15
      Top             =   4740
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BackColor       =   8438015
      MouseIcon       =   "frmTestMaskedEdit.frx":282E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelText         =   ""
      Text            =   "meMainExample"
      HideSelection   =   -1  'True
   End
   Begin VB.Label lblDesc 
      Caption         =   $"frmTestMaskedEdit.frx":284A
      Height          =   855
      Left            =   90
      TabIndex        =   6
      Top             =   2340
      Width           =   3105
   End
   Begin VB.Label lblAllowedKeys 
      AutoSize        =   -1  'True
      Caption         =   "Allowed Keys:"
      Height          =   195
      Left            =   3300
      TabIndex        =   13
      Top             =   60
      Width           =   990
   End
   Begin VB.Label lblAlignment 
      AutoSize        =   -1  'True
      Caption         =   "Alignment:"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   3930
      Width           =   735
   End
   Begin VB.Label lblMask 
      AutoSize        =   -1  'True
      Caption         =   "Mask:"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   3330
      Width           =   435
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      Caption         =   "Format:"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   3330
      Width           =   525
   End
End
Attribute VB_Name = "frmTestMaskedEdit"
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

Private Sub cboAlignment_Click()
    Me.meMainExample.Alignment = Me.cboAlignment.ItemData(Me.cboAlignment.ListIndex)

End Sub

Private Sub Form_Load()
    With Me.cboAlignment
        .AddItem "Left"
        .ItemData(.NewIndex) = vbLeftJustify
        .AddItem "Center"
        .ItemData(.NewIndex) = vbCenter
        .AddItem "Right"
        .ItemData(.NewIndex) = vbRightJustify
        .ListIndex = 0
    End With
    
    With Me.listAllowedKeys

        .AddItem "OnlyDATE"
        .ItemData(.NewIndex) = jmeOnlyDATE
        .AddItem "OnlyMONEY"
        .ItemData(.NewIndex) = jmeOnlyMONEY
        .AddItem "OnlyNUMBERS"
        .ItemData(.NewIndex) = jmeOnlyNUMBERS
        .AddItem "OnlyPHONE"
        .ItemData(.NewIndex) = jmeOnlyPHONE
        .AddItem "OnlyTIME"
        .ItemData(.NewIndex) = jmeOnlyTIME
        .AddItem "CaseUpper"
        .ItemData(.NewIndex) = jmeUppercase
        .AddItem "CaseLower"
        .ItemData(.NewIndex) = jmeLowercase
        .AddItem "NoDoubleQuotes"
        .ItemData(.NewIndex) = jmeNoDoubleQuotes
        .AddItem "NoSingleQuotes"
        .ItemData(.NewIndex) = jmeNoSingleQuotes
        .AddItem "NoSpaces"
        .ItemData(.NewIndex) = jmeNoSpaces
        .AddItem "AllowAMPM"
        .ItemData(.NewIndex) = jmeAllowAMPM
        .AddItem "AllowColon"
        .ItemData(.NewIndex) = jmeAllowColon
        .AddItem "AllowDecimal"
        .ItemData(.NewIndex) = jmeAllowDecimal
        .AddItem "AllowDollarSigns"
        .ItemData(.NewIndex) = jmeAllowDollarSigns
        .AddItem "AllowForwardSlash"
        .ItemData(.NewIndex) = jmeAllowForwardSlash
        .AddItem "AllowNegative"
        .ItemData(.NewIndex) = jmeAllowNegative
        .AddItem "AllowParenthesis"
        .ItemData(.NewIndex) = jmeAllowParenthesis
        .AddItem "AllowPounds"
        .ItemData(.NewIndex) = jmeAllowPounds
        .AddItem "AllowSpaces"
        .ItemData(.NewIndex) = jmeAllowSpaces
        .AddItem "AllowStars"
        .ItemData(.NewIndex) = jmeAllowStars

    End With
    
    
End Sub

Private Sub listAllowedKeys_Click()
    Dim e As jmeAllowedKeys
    Dim l As Long
    With Me.listAllowedKeys
        For l = 0 To .ListCount - 1
            If .Selected(l) Then
                e = e Or .ItemData(l)
            End If
        Next l
    End With
    Me.meMainExample.AllowedKeys = e
    
End Sub

Private Sub txtFormat_Change()
    On Error Resume Next
    Me.meMainExample.Format = Me.txtFormat.Text
End Sub

Private Sub txtMask_Change()
    On Error Resume Next
    Me.meMainExample.Mask = Me.txtMask.Text
    
End Sub
