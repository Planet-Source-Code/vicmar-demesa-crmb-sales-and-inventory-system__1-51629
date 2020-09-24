VERSION 5.00
Begin VB.Form frmAbout 
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -360
      Width           =   7695
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1
         Top             =   4440
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000005&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   240
         Top             =   1560
         Width           =   7335
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   7695
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   7695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Comments Or Suggestions Please Email Us At "
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   4095
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
         TabIndex        =   8
         Top             =   4080
         Width           =   1635
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
         TabIndex        =   7
         Top             =   600
         Width           =   1725
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
         TabIndex        =   6
         Top             =   1080
         Width           =   3720
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Height          =   1575
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   7695
      End
      Begin VB.Shape Shape1 
         Height          =   315
         Left            =   3720
         Top             =   4920
         Width           =   3930
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Developer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblWarning 
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
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   6735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Copyrights (c) 2003 Developed By SCaVeNGeR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         Top             =   4920
         Width           =   3930
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
         TabIndex        =   1
         Top             =   5340
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Timer1.Interval = Timer1.Interval - 2
Screen.MousePointer = vbHourglass
If Timer1.Interval <= 1 Then
Unload frmAbout
Screen.MousePointer = vbDefault
End If
End Sub


