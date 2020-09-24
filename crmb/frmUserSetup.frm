VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmUserSetup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   5355
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmUserSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4140
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "User"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   3360
      TabIndex        =   31
      Top             =   6000
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmUserSetup.frx":000C
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmUserSetup.frx":0020
      TabIndex        =   23
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Frame Frame2 
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
      Height          =   1725
      Left            =   -120
      TabIndex        =   24
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton cmd7 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   300
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Appearance      =   0  'Flat
      Caption         =   "&Update"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Appearance      =   0  'Flat
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Caption         =   "&Add"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmd6 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmd4 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Users"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   3975
      Begin VB.ComboBox Combo1 
         DataField       =   "Database"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo7 
         DataField       =   "Usersetup"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Text            =   "Combo7"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo8 
         DataField       =   "Transactions"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Text            =   "Combo8"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo9 
         DataField       =   "Files"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Text            =   "Combo9"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo10 
         DataField       =   "Reports"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   15
         Text            =   "Combo10"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Database"
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   29
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Users"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Transact"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Reports"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Files"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   3975
      Begin VB.TextBox Text2 
         DataField       =   "Password"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   1
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   "User"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "User"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   25
      Left            =   240
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "   User Setup"
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
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmUserSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
On Error GoTo AddErr
If Text1.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
ButtonEnabled
Data1.Recordset.AddNew
Text1.SetFocus
End If
Exit Sub
AddErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd4_Click()
On Error GoTo CancelErr
If Text1.Enabled = True Then
Data1.Recordset.CancelUpdate
ButtonEnabled1
Else
MsgBox "Cancel without record to add/edit.", vbExclamation, "Attention"
End If
Exit Sub
CancelErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd6_Click()
On Error Resume Next
If Text1.Enabled = True Then
Exit Sub
End If
Dim a
If Text1.Text = "" Then
MsgBox "No current record to delete.", vbExclamation, "Attention"
Exit Sub
Else
a = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm")
If a = vbYes Then
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
Else
Cancel = True
Exit Sub
End If
End If
End Sub


Private Sub cmd2_Click()
On Error GoTo ModifyErr
If Text1.Enabled = True Then
MsgBox "Please save first before edit.", vbExclamation, "Attention"
Exit Sub
End If
Data1.Recordset.Edit
Combo2.Text = Combo1.Text
Combo3.Text = Combo7.Text
Combo4.Text = Combo8.Text
Combo5.Text = Combo9.Text
Combo6.Text = Combo10.Text
ButtonEnabled
Text7.Enabled = True
Text1.SetFocus
Exit Sub
ModifyErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"

End Sub

Private Sub cmd7_Click()
On Error GoTo CloseErr
Dim a
If Text1.Enabled = True Then
a = MsgBox("Do you want to save the record before you exit?", vbYesNo + vbQuestion, "Confirm")
If a = vbYes Then
cmd3_Click
Exit Sub
Else
cmd4_Click
Unload Me
End If
End If
Unload Me
Exit Sub
CloseErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub



Private Sub Form_Load()
  Combo2.AddItem ("0")
  Combo2.AddItem ("1")
  Combo2.ListIndex = 0
  Combo3.AddItem ("0")
  Combo3.AddItem ("1")
  Combo3.ListIndex = 0
  Combo4.AddItem ("0")
  Combo4.AddItem ("1")
  Combo4.ListIndex = 0
  Combo5.AddItem ("0")
  Combo5.AddItem ("1")
  Combo5.ListIndex = 0
  Combo6.AddItem ("0")
  Combo6.AddItem ("1")
  Combo6.ListIndex = 0
CenterForm Me
Text7.Enabled = False
Data1.DatabaseName = App.Path + "\crmdb.mdb"
Data2.DatabaseName = App.Path + "\crmdb.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a, ans
If Cancel = 0 Then
    If Text2.Enabled = True Then
    a = MsgBox("Do you want to save the record before you exit?", vbYesNo + vbQuestion, "Confirm")
    If a = vbYes Then
            Cancel = 1
            cmd3_Click
            Exit Sub
            
        Else
        Cancel = 1
        cmd4_Click
        Unload Me
         End If
        End If
    End If
End Sub

Private Sub cmd3_Click()
On Error GoTo SaveErr
Dim a, b
If Text2.Enabled = False Then
MsgBox "Please add/edit before saving.", vbExclamation, "Attention"
Exit Sub
End If
Combo1.Text = Combo2.Text
Combo7.Text = Combo3.Text
Combo8.Text = Combo4.Text
Combo9.Text = Combo5.Text
Combo10.Text = Combo6.Text

If Text1.Text = "" Then
MsgBox "Please enter the user.", vbExclamation, "Attention"
Text1.SetFocus
Exit Sub
ElseIf Text2.Text = "" Then
MsgBox "Please enter the password.", vbExclamation, "Attention"
Text2.SetFocus
Exit Sub
End If
If Text7.Enabled = False Then
a = Text1.Text
Data2.RecordSource = ("Select * from Users where user = '" & a & "'")
Data2.Refresh
If Text3.Text = "" Then
Data2.RecordSource = ("Select * from Users order by user")
Data2.Refresh
Else
MsgBox "User already exist, please try again.", vbExclamation, "Attention"
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
Else
Text7.Enabled = False
End If
MsgBox "Record successfully saved.", vbInformation, "Save"
Data1.Recordset.Update
Data1.Refresh
ButtonEnabled1
Exit Sub
SaveErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub


Public Sub ButtonEnabled()
Text1.Enabled = True
Text2.Enabled = True
cmd1.Visible = False
cmd2.Visible = False
cmd3.Visible = True
cmd4.Visible = True
cmd6.Enabled = False
DBGrid1.Enabled = False
Combo2.Visible = True
Combo1.Visible = False
Combo3.Visible = True
Combo6.Visible = True
Combo4.Visible = True
Combo5.Visible = True
Combo7.Visible = False
Combo8.Visible = False
Combo9.Visible = False
Combo10.Visible = False
End Sub

Public Sub ButtonEnabled1()
Text1.Enabled = False
Text2.Enabled = False
Combo2.Visible = False
Combo1.Visible = True
Combo3.Visible = False
Combo6.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Combo7.Visible = True
Combo8.Visible = True
Combo9.Visible = True
Combo10.Visible = True
cmd1.Visible = True
cmd2.Visible = True
cmd3.Visible = False
cmd4.Visible = False
cmd6.Enabled = True
DBGrid1.Enabled = True
End Sub






