VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form lib_frmCal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboYears 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   915
   End
   Begin VB.ComboBox cboMonths 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   1440
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   2160
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2535
      TabIndex        =   1
      Top             =   30
      Width           =   240
   End
   Begin MSACAL.Calendar cal 
      Height          =   1890
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   2790
      _Version        =   524288
      _ExtentX        =   4921
      _ExtentY        =   3334
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1998
      Month           =   8
      Day             =   31
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   0   'False
      ShowDays        =   0   'False
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "lib_frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ******************************************************************
'   Project:    Calendar control
'
'   Author:     LockwoodTech
'
'   Legal:      You are free to use and distribute
'   ******************************************************************
Option Explicit

Private Const m_strMY_NAME As String = "lib_frmCal"    ' Objects name


'   CALL THIS FORM OBJECT USING THE FOLLOWING SYNTAX:
'   Call lib_frmCal.GetValue(ctl:=mebTest, F:=mdiMain)

'   Note:  MS Cal has a bug where the Month and Year combos periodically
'   migrate down on the form making it unusable.  For this reason I
'   added my own picklists.

 Private m_strValue       As String
 Private m_bolCancel      As Boolean
 Private m_bolClear       As Boolean

'   POSITION ADJUSTMENTS
'   The TOP and LEFT values for the calander control are highly
'   dependant on the forms/controls that exist on the interface
'   (i.e. MDI child form, Tab controls, Outlines etc.).  You may
'   have to adjust the normal position calculation to compensate.
'   In this CASE my app. is MDI with a toolbar, so I have to adjust
'   for this extra space on top.  I also have to adjust the left
'   value to compensate for the fact that my control is on a form
'   that shares space with another form (frmTree_View_Control)
'   to the left of it.

Public intTopAdjust     As Integer
Public intLeftAdjust    As Integer

 Sub cal_Click()
    On Error GoTo PROC_ERR
    
    m_strValue = Format(cal, "MM/DD/YYYY")
    cboMonths = Format(m_strValue, "Mmm")
    cboYears = Format(m_strValue, "YYYY")
    Me.Hide
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Resume PROC_EXIT:
End Sub

Public Sub GetValue(ByRef ctl As Control, Optional ByVal intContainerTop As Integer)
 Dim intTop As Integer
 Dim intLeft As Integer

    On Error GoTo PROC_ERR
    
    m_bolCancel = True
    m_bolClear = False

    intTop = intTopAdjust + intContainerTop + Screen.ActiveForm.Top + ctl.Top + ctl.Height
    intLeft = intLeftAdjust + Screen.ActiveForm.Left + ctl.Left - cal.Width + ctl.Width
    
    m_strValue = IIf(IsDate(ctl.Text), ctl.Text, "")
    
    If m_strValue = "" Or m_strValue = "__/__/____" Then
        cal.Today
        cboMonths = Format(Now, "Mmm")
        cboYears = Format(Now, "YYYY")
    Else
        m_strValue = ctl.Text
        cal.Value = m_strValue
        cboMonths = Format(m_strValue, "Mmm")
        cboYears = Format(m_strValue, "YYYY")
    End If
    
    m_bolCancel = False      '   to clean this up
    Me.Top = intTop
    Me.Left = intLeft

    Me.Show vbModal
    
    If m_bolClear Then
        ctl.Text = IIf(TypeOf ctl Is MaskEdBox, "__/__/____", m_strValue)
    ElseIf Not m_bolCancel And Len(m_strValue) <> 0 Then
        ctl.Text = m_strValue
    End If
    
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Resume PROC_EXIT:
End Sub

Private Sub cboMonths_Click()
    cal.Month = cboMonths.ListIndex + 1
End Sub

Private Sub cboYears_Click()
    cal.Year = cboYears
End Sub

Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    m_bolClear = True
    m_strValue = ""
    Unload Me
End Sub

Private Sub Form_Initialize()
 Dim n As Integer
    On Error GoTo PROC_ERR
    
    cboMonths.AddItem "Jan"
    cboMonths.AddItem "Feb"
    cboMonths.AddItem "Mar"
    cboMonths.AddItem "Apr"
    cboMonths.AddItem "May"
    cboMonths.AddItem "Jun"
    cboMonths.AddItem "Jul"
    cboMonths.AddItem "Aug"
    cboMonths.AddItem "Sep"
    cboMonths.AddItem "Oct"
    cboMonths.AddItem "Nov"
    cboMonths.AddItem "Dec"
    
    For n = 1970 To 2020
        cboYears.AddItem n
    Next n
    
    cal.Today
    
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Resume PROC_EXIT:
End Sub

Sub Form_Unload(Cancel As Integer)
    If Screen.ActiveForm Is Me Then Cancel = True
    Me.Hide
    m_bolCancel = True
End Sub

Private Sub cal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case Is = 13    '   Enter
        Call cal_Click
    Case Is = 27    '   Escape
        Unload Me
    End Select
End Sub

