VERSION 5.00
Begin VB.Form frmChangepassword 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2085
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   6015
   Icon            =   "frmChangepassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txtconfirmpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4320
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   645
         Width           =   1575
      End
      Begin VB.TextBox txtnewpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1575
      End
      Begin VB.TextBox txtUserid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Confirm Password"
         Height          =   270
         Index           =   4
         Left            =   3000
         TabIndex        =   16
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         Caption         =   "New Password"
         Height          =   270
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "Old Password"
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "User "
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   1140
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6540
      TabIndex        =   6
      Top             =   2940
      Visible         =   0   'False
      Width           =   6540
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmChangepassword.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmChangepassword.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmChangepassword.frx":0690
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmChangepassword.frx":09D2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   11
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmChangepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mbDataChanged As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
    If txtnewpass.Text = "" Or txtconfirmpass.Text = "" Then
        Beep
        MsgBox "One of your login fields is blank, please try again.", vbInformation, "Attention"
        Exit Sub
    End If
    adoPrimaryRS.MoveFirst
      Do Until found Or adoPrimaryRS.EOF
        usercode = adoPrimaryRS.Fields("user_id").Value
        If usercode = txtUserid.Text Then
            found = True
            user_id = usercode
            Exit Do
        Else
            adoPrimaryRS.MoveNext
        End If
        Loop
        
       If found Then
            Password = adoPrimaryRS.Fields("user_pass").Value
            If Not Password = txtPassword.Text Then
               MsgBox "Old password is incorrect!", vbCritical, "Warning"
                txtPassword.SetFocus
            Else
               If txtnewpass.Text = txtconfirmpass.Text Then
               adoPrimaryRS!user_pass = Trim(txtnewpass.Text)
               adoPrimaryRS.Update
                 MsgBox "Password Has Been Changed !", vbInformation, "Information"
                 Unload Me
               Else
                MsgBox "New password & Confirmed password should be the same !", vbCritical, "Warning"
                txtnewpass.SetFocus
               End If
            End If
  End If
End Sub

Private Sub Form_Load()

  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\psmdb.mdb;"
  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select user_id,user_pass from users Order by user_id", db, adOpenStatic, adLockOptimistic

'  Dim oText As TextBox
'  'Bind the text boxes to the data provider
'  For Each oText In Me.txtFields
'    Set oText.DataSource = adoPrimaryRS
'  Next
  txtUserid.Text = usercode
  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub



Private Sub Form_Unload(cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

