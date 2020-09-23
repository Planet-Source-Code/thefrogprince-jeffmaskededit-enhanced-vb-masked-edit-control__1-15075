VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Main Menu For Jeff's Examples"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJeffMaskedEdit 
      Caption         =   "jeff&MaskedEdit"
      Height          =   555
      Left            =   2010
      TabIndex        =   1
      Top             =   270
      Width           =   1515
   End
   Begin VB.CommandButton cmdJeffFrame 
      Caption         =   "jeff&Frame"
      Height          =   555
      Left            =   330
      TabIndex        =   0
      Top             =   270
      Width           =   1515
   End
End
Attribute VB_Name = "frmMainMenu"
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

Private Sub cmdJeffFrame_Click()
    frmTestFrame.Show vbModal, Me
    Set frmTestFrame = Nothing
    
End Sub

Private Sub cmdJeffMaskedEdit_Click()
    frmTestMaskedEdit.Show vbModal, Me
    Set frmTestMaskedEdit = Nothing
    
End Sub
