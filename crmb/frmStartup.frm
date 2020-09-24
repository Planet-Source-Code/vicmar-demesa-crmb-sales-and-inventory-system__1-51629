VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4830
   ClientLeft      =   2145
   ClientTop       =   915
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00404000&
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   7695
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   5040
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   960
         Top             =   5040
      End
      Begin MSComctlLib.ProgressBar Bar 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Loading..."
         Top             =   5160
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000005&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   240
         Top             =   1440
         Width           =   7455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ussiness Software Solution"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2010
         TabIndex        =   9
         Top             =   960
         Width           =   3720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRMB"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "tekhah@yahoo.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   4800
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Comments Or Suggestions Please Email At "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   3840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3450
         TabIndex        =   4
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unauthorized Used of This Software Is Strictly Prohibited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   4665
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003  SCaVeNGeR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Left            =   7800
         TabIndex        =   2
         Top             =   5340
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y As Integer


Private Sub Timer1_Timer()
Dim a As Long, b As Integer
Bar.Value = Bar.Value + 2
Screen.MousePointer = vbHourglass

If Bar.Value <= 30 Then
Label1 = "Initializing....."
ElseIf Bar.Value <= 50 Then
Label1 = "Loading components....."
ElseIf Bar.Value <= 70 Then
Label1 = "Integrating Database...."
ElseIf Bar.Value <= 100 Then
Label1 = "Please wait..."
End If
If Bar.Value = 100 Then
If Timer1.Interval >= 1 Then
Unload frmStartup
Screen.MousePointer = vbDefault
Load frmLogin
frmLogin.Show
End If
End If
End Sub

