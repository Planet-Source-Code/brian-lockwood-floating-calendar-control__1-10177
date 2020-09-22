VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmArchive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archive Information"
   ClientHeight    =   3750
   ClientLeft      =   405
   ClientTop       =   840
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   780
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3810
      TabIndex        =   10
      Top             =   3405
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   285
      Left            =   2940
      TabIndex        =   9
      Top             =   3420
      Width           =   825
   End
   Begin VB.Frame frmContainer 
      Caption         =   "Container"
      Height          =   1185
      Left            =   435
      TabIndex        =   5
      Top             =   1500
      Width           =   3855
      Begin VB.TextBox txtContainerDate 
         Height          =   300
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Req"
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdCal 
         Height          =   250
         Index           =   2
         Left            =   3210
         Picture         =   "Testfrm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Within a Frame:"
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   525
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCal 
      Height          =   250
      Index           =   0
      Left            =   3375
      Picture         =   "Testfrm.frx":0383
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   450
      Width           =   230
   End
   Begin VB.CommandButton cmdCal 
      Height          =   250
      Index           =   1
      Left            =   3360
      Picture         =   "Testfrm.frx":0706
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   795
      Width           =   230
   End
   Begin VB.TextBox txtArchDate 
      Height          =   300
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Req"
      Top             =   442
      Width           =   1200
   End
   Begin VB.Label lblArchDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Standard Text Box >>:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblDestrDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Microsoft Mask Control >>:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   1995
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCal_Click(Index As Integer)
 Dim ctlNm As Control

    
    Select Case Index
        Case Is = 0
            Set ctlNm = txtArchDate
            lib_frmCal.intTopAdjust = g_intMDI_CHILD + g_intOTHER_FORM_BORDER_STYLE
            lib_frmCal.intLeftAdjust = 0
        Case Is = 1
            Set ctlNm = mskEdit
            lib_frmCal.intTopAdjust = g_intMDI_CHILD + g_intOTHER_FORM_BORDER_STYLE
            lib_frmCal.intLeftAdjust = 0
        Case Is = 2
            Set ctlNm = txtContainerDate
            lib_frmCal.intTopAdjust = frmContainer.Top + g_intMDI_CHILD + g_intOTHER_FORM_BORDER_STYLE
            lib_frmCal.intLeftAdjust = frmContainer.Left + 0
        Case Else
            MsgBox "Case Else Error", vbCritical
    End Select
    
    '   Make adjustment because this is a Fixed dialog form
    

    Call lib_frmCal.GetValue(ctl:=ctlNm)
    
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
    End
End Sub

































