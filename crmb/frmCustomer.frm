VERSION 5.00
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   4545
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   9225
   ControlBox      =   0   'False
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9225
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   7680
      TabIndex        =   33
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search"
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
      Height          =   1215
      Left            =   6480
      TabIndex        =   26
      Top             =   720
      Width           =   2655
      Begin VB.ComboBox Combo3 
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
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox comSearch 
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
         Height          =   315
         ItemData        =   "frmCustomer.frx":000C
         Left            =   120
         List            =   "frmCustomer.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   6975
      Begin VB.CommandButton cmd4 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd6 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd1 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd2 
         Appearance      =   0  'Flat
         Caption         =   "&Modify"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd3 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd7 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
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
      ForeColor       =   &H00404000&
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   6255
      Begin VB.TextBox Text10 
         DataField       =   "Address"
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
         Height          =   315
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         DataField       =   "CustomerCode"
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
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         DataField       =   "FirstName"
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
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         DataField       =   "Phone"
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
         Height          =   315
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   4
         Top             =   1980
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         DataField       =   "LastName"
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
         Height          =   315
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         DataField       =   "Email"
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
         Height          =   315
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         DataField       =   "FirstName"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Text            =   "Text8"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         DataField       =   "CustomerCode"
         DataSource      =   "Data4"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Last Name"
         Height          =   375
         Index           =   4
         Left            =   3840
         TabIndex        =   30
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "E-mail Address"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Phone Number"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Code"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "autodept"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ctr"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from Customers order by LastName"
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   240
      Top             =   360
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "   Customer Database"
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
      TabIndex        =   24
      Top             =   0
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   240
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      ForeColor       =   &H00404040&
      Height          =   1260
      Index           =   1
      Left            =   6480
      TabIndex        =   21
      Top             =   720
      Width           =   2700
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   2895
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   780
      Width           =   6255
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoCustomer As Recordset
Private Sub cmd1_Click()
On Error GoTo AddErr
If Text2.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
ButtonEnabled
Text6.Text = Val(Text6.Text) + 1
Text6.Text = Format$(Text6.Text, "000")
Data1.Recordset.AddNew
Text1.Text = Text6.Text
Text2.SetFocus
End If
Exit Sub
AddErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd4_Click()
On Error GoTo CancelErr
If Text2.Enabled = True Then
Data1.Recordset.CancelUpdate
ButtonEnabled1
Else
MsgBox "Cancel without record to add/edit.", vbExclamation, "Attention"
End If
If Text7.Enabled = False Then
Text6.Text = Val(Text6.Text) - 1
Else
Text7.Enabled = False
End If
Exit Sub
CancelErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub Combo1_Click()
On Error GoTo Search1Err
Dim a
If comSearch.Text = "Customer Code" Then
a = Combo1.Text
Data1.RecordSource = ("Select * from Customers where CustomerCode = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by CustomerCode")
Data1.Refresh
End If
ElseIf comSearch.Text = "First Name" Then
a = Combo2.Text
Data1.RecordSource = ("Select * from Customers where FirstName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by FirstName")
Data1.Refresh
End If
ElseIf comSearch.Text = "Last Name" Then
a = Combo3.Text
Data1.RecordSource = ("Select * from Customers where LastName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo3.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
End If
End If
Exit Sub
Search1Err:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub Combo2_Click()
On Error GoTo Search2Err
Dim a
If comSearch.Text = "Customer Code" Then
a = Combo1.Text
Data1.RecordSource = ("Select * from Customers where CustomerCode = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by CustomerCode")
Data1.Refresh
End If
ElseIf comSearch.Text = "First Name" Then
a = Combo2.Text
Data1.RecordSource = ("Select * from Customers where FirstName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by FirstName")
Data1.Refresh
End If
ElseIf comSearch.Text = "Last Name" Then
a = Combo3.Text
Data1.RecordSource = ("Select * from Customers where LastName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo3.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
End If
End If
Exit Sub
Search2Err:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub Combo3_Click()
On Error GoTo Search3Err
Dim a
If comSearch.Text = "Customer Code" Then
a = Combo1.Text
Data1.RecordSource = ("Select * from Customers where CustomerCode = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo1.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by CustomerCode")
Data1.Refresh
End If
ElseIf comSearch.Text = "First Name" Then
a = Combo2.Text
Data1.RecordSource = ("Select * from Customers where FirstName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo2.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by FirstName")
Data1.Refresh
End If
ElseIf comSearch.Text = "Last Name" Then
a = Combo3.Text
Data1.RecordSource = ("Select * from Customers where LastName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo3.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
End If
End If
Exit Sub
Search3Err:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub Command1_Click()
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
End Sub

Private Sub comSearch_Change()
If comSearch.Text = "First Name" Then
Combo1.Visible = False
Combo2.Visible = True
Combo3.Visible = False
ElseIf comSearch.Text = "Customer Code" Then
Combo1.Visible = True
Combo2.Visible = False
Combo3.Visible = False
ElseIf comSearch.Text = "Last Name" Then
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = True
End If
End Sub

Private Sub comSearch_LostFocus()
If comSearch.Text = "First Name" Then
Combo1.Visible = False
Combo2.Visible = True
Combo3.Visible = False
ElseIf comSearch.Text = "Customer Code" Then
Combo1.Visible = True
Combo2.Visible = False
Combo3.Visible = False
ElseIf comSearch.Text = "Last Name" Then
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = True
End If
End Sub


Private Sub cmd6_Click()
On Error Resume Next
If Text2.Enabled = True Then
Exit Sub
End If
Dim a
If Text2.Text = "" Then
MsgBox "No current record to delete.", vbExclamation, "Attention"
Exit Sub
Else
a = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm")
If a = vbYes Then
Text6.Text = Val(Text6.Text)
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
Else
Cancel = True
Exit Sub
End If
End If
End Sub

Private Sub cmd2_Click()
On Error GoTo ModifyErr
If Text2.Enabled = True Then
MsgBox "Please save first before edit.", vbExclamation, "Attention"
Exit Sub
End If
Data1.Recordset.Edit
ButtonEnabled
Text2.SetFocus
Text7.Enabled = True
Exit Sub
ModifyErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub cmd7_Click()
On Error GoTo CloseErr
Dim a
If Text2.Enabled = True Then
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
Dim X As Integer
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\crmdb.mdb;"
  Set AdoCustomer = New Recordset
     AdoCustomer.Open "select CustomerCode,Firstname,LastName from Customers Order by LastName", db, adOpenStatic, adLockOptimistic
  AdoCustomer.Requery
  For X = 1 To AdoCustomer.RecordCount
    Combo2.AddItem AdoCustomer.Fields("Firstname")
    AdoCustomer.MoveNext
  Next X
  AdoCustomer.Requery
  For X = 1 To AdoCustomer.RecordCount
    Combo1.AddItem AdoCustomer.Fields("CustomerCode")
    AdoCustomer.MoveNext
  Next X
  AdoCustomer.Requery
  For X = 1 To AdoCustomer.RecordCount
    Combo3.AddItem AdoCustomer.Fields("LastName")
    AdoCustomer.MoveNext
  Next X
   Combo1.ListIndex = 0
  Combo2.ListIndex = 0
  Combo3.ListIndex = 0
  comSearch.AddItem ("First Name")
  comSearch.AddItem ("Last Name")
  comSearch.AddItem ("Customer Code")
  comSearch.ListIndex = 1
CenterForm Me
Text7.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = False
Data1.DatabaseName = App.Path + "\crmdb.mdb"
Data2.DatabaseName = App.Path + "\crmdb.mdb"
Data3.DatabaseName = App.Path + "\crmdb.mdb"
Data4.DatabaseName = App.Path + "\crmdb.mdb"
End Sub

Private Sub cmd3_Click()
On Error GoTo SaveErr
Dim a, b
If Text2.Enabled = False Then
MsgBox "Please add/edit before saving.", vbExclamation, "Attention"
Exit Sub
End If
If Text7.Enabled = False Then
a = Text2.Text
b = Text1.Text
c = Text4.Text
Data4.RecordSource = ("Select * from Customers where CustomerCode= '" & b & "'")
Data4.Refresh
Data3.RecordSource = ("Select * from Customers where FirstName = '" & a & "' and LastName ='" & c & "'")
Data3.Refresh
If Text9.Text = "" Then
Data4.RecordSource = ("Select * from Customers order by CustomerCode")
Data4.Refresh
Else
MsgBox "Customer Code already exist, please try again.", vbExclamation, "Attention"
Text1.Locked = False
Text1.SetFocus
Exit Sub
End If
If Text8.Text = "" Then
Data3.RecordSource = ("Select * from Customers order by FirstName")
Data3.Refresh
Else
MsgBox "Customer name already exist, please try again.", vbExclamation, "Attention"
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If
Else
Text7.Enabled = False
End If

If Text2.Text = "" Then
MsgBox "Please enter the first name.", vbExclamation, "Attention"
Text2.SetFocus
Exit Sub
ElseIf Text4.Text = "" Then
MsgBox "Please enter the last name.", vbExclamation, "Attention"
Text4.SetFocus
Exit Sub
ElseIf Text5.Text = "" Then
MsgBox "Please enter the e-mail address.", vbExclamation, "Attention"
Text5.SetFocus
Exit Sub
ElseIf Text3.Text = "" Then
MsgBox "Please enter the phone number or enter (-) if phone number is not available.", vbExclamation, "Attention"
Text3.SetFocus
Exit Sub
ElseIf Text10.Text = "" Then
MsgBox "Please enter the address.", vbExclamation, "Attention"
Text10.SetFocus
Exit Sub
End If
MsgBox "Record successfully saved.", vbInformation, "Save"
Data1.Recordset.Update
Data1.RecordSource = ("Select * from Customers order by LastName")
Data1.Refresh
ButtonEnabled1
Exit Sub
SaveErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"
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

Private Sub text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 45 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Public Sub ButtonEnabled()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text6.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text10.Enabled = True
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = True
cmd4.Enabled = True
cmd6.Enabled = False
Command1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
comSearch.Enabled = False
Command1.Enabled = False
Text1.Locked = True
End Sub

Public Sub ButtonEnabled1()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text10.Enabled = False
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
comSearch.Enabled = True
Command1.Enabled = True
cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = False
cmd4.Enabled = False
cmd6.Enabled = True
Command1.Enabled = True
End Sub
