VERSION 5.00
Begin VB.Form frmProduct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   3960
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   10095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   10095
   Begin VB.Frame Frame3 
      Caption         =   "Search"
      ForeColor       =   &H00404040&
      Height          =   1695
      Left            =   7320
      TabIndex        =   27
      Top             =   840
      Width           =   2655
      Begin VB.ComboBox Combo6 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox Combo5 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox comSearch 
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmProduct.frx":000C
         Left            =   120
         List            =   "frmProduct.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1830
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   1200
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   6975
      Begin VB.CommandButton cmd7 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd3 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd2 
         Appearance      =   0  'Flat
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd1 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmd6 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmd4 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Products"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text6 
      DataField       =   "autoprod"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00404000&
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.TextBox Text2 
         DataField       =   "ProductName"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         DataField       =   "ProductName"
         DataSource      =   "Data5"
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         DataField       =   "ReorderLevel"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5040
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         DataField       =   "UnitsInStock"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataField       =   "ProductCode"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         DataField       =   "UnitPrice"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Reorder Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3840
         TabIndex        =   23
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Units in Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1395
      End
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   21
      Text            =   "Text7"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox ctr1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox ctr2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   8880
      Width           =   255
   End
   Begin VB.TextBox ctr3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   8820
      Width           =   255
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ctr"
      Top             =   9120
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from Products order by ProductName"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      ForeColor       =   &H00404040&
      Height          =   1740
      Index           =   1
      Left            =   7320
      TabIndex        =   31
      Top             =   840
      Width           =   2700
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   240
      Top             =   360
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "   Product Database"
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
      TabIndex        =   25
      Top             =   0
      Width           =   10455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   4020
      TabIndex        =   24
      Top             =   2460
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   6915
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoProduct As Recordset
Private Sub cmd1_Click()
On Error GoTo AddErr
If Text2.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
ButtonEnabled
Text6.Text = Val(Text6.Text) + 1
Text6.Text = Format$(Text6.Text, "00000")
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

Private Sub Combo6_Click()
On Error GoTo SearchErr
Dim a
If comSearch.Text = "Product Code" Then
a = Combo5.Text
Data1.RecordSource = ("Select * from Products where ProductCode = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo5.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Products order by ProductCode")
Data1.Refresh
End If
Else
If comSearch.Text = "Product Name" Then
a = Combo6.Text
Data1.RecordSource = ("Select * from Products where ProductName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo6.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Products order by ProductName")
Data1.Refresh
End If
End If
End If
Exit Sub
SearchErr:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub
Private Sub Combo5_Click()
On Error GoTo Search1Err
Dim a
If comSearch.Text = "Product Code" Then
a = Combo5.Text
Data1.RecordSource = ("Select * from Products where ProductCode = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo5.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Products order by ProductCode")
Data1.Refresh
End If
Else
If comSearch.Text = "Product Name" Then
a = Combo6.Text
Data1.RecordSource = ("Select * from Products where ProductName = '" & a & "'")
Data1.Refresh
If Text2.Text = "" Then
a = Combo6.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data1.RecordSource = ("Select * from Products order by ProductName")
Data1.Refresh
End If
End If
End If
Exit Sub
Search1Err:
 MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub Command1_Click()
Data1.RecordSource = ("Select * from Products order by ProductName")
Data1.Refresh
End Sub

Private Sub comSearch_Change()
If comSearch.Text = "Product Name" Then
Combo5.Visible = False
Combo6.Visible = True

Else
If comSearch.Text = "Product Code" Then
Combo5.Visible = True
Combo6.Visible = False

End If
End If
End Sub

Private Sub comSearch_LostFocus()
If comSearch.Text = "Product Name" Then
Combo5.Visible = False
Combo6.Visible = True

Else
If comSearch.Text = "Product Code" Then
Combo5.Visible = True
Combo6.Visible = False

End If
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
Data1.RecordSource = ("Select * from Products order by ProductName")
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
  Set AdoProduct = New Recordset
     AdoProduct.Open "select ProductCode,Productname from Products Order by ProductName", db, adOpenStatic, adLockOptimistic
  
For X = 1 To AdoProduct.RecordCount
    Combo6.AddItem AdoProduct.Fields("Productname")
    AdoProduct.MoveNext
  Next X
  AdoProduct.Requery
  For X = 1 To AdoProduct.RecordCount
    Combo5.AddItem AdoProduct.Fields("ProductCode")
    AdoProduct.MoveNext
  Next X
  Combo5.ListIndex = 0
  Combo6.ListIndex = 0
CenterForm Me
  comSearch.AddItem ("Product Name")
  comSearch.AddItem ("Product Code")
  comSearch.ListIndex = 0
Data1.DatabaseName = App.Path + "\crmdb.mdb"
Data2.DatabaseName = App.Path + "\crmdb.mdb"
Data5.DatabaseName = App.Path + "\crmdb.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a
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
Dim a
If Text2.Enabled = False Then
MsgBox "Please add/edit before saving.", vbExclamation, "Attention"
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please enter the product name.", vbExclamation, "Attention"
Text2.SetFocus
Exit Sub
ElseIf Text8.Text = "" Or Val(Text8.Text) = "0" Then
MsgBox "Please enter the units stock.", vbExclamation, "Attention"
Text8.SetFocus
Exit Sub
ElseIf Text9.Text = "" Or Val(Text9.Text) = "0" Then
MsgBox "Please enter the reorder level.", vbExclamation, "Attention"
Text9.SetFocus
Exit Sub
ElseIf Text5.Text = "" Or Val(Text5.Text) = "0" Then
MsgBox "Please enter the unit price.", vbExclamation, "Attention"
Text5.SetFocus
Exit Sub
ElseIf Val(Text8.Text) < Val(Text9.Text) Then
MsgBox "Units in stock must be greater than reorder level, please try again.", vbExclamation, "Attention"
Text8.Text = ""
Text8.SetFocus
Exit Sub
End If
If Text7.Enabled = False Then
a = Text2.Text
Data5.RecordSource = ("Select * from Products where ProductName = '" & a & "'")
Data5.Refresh
If Text3.Text = "" Then
Data5.RecordSource = ("Select * from Products order by ProductName")
Data5.Refresh
Else
MsgBox "Product name already exist, please try again.", vbExclamation, "Attention"
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If
Else
Text7.Enabled = False
End If
MsgBox "Record successfully saved.", vbInformation, "Save"
Data1.Recordset.Update
Data1.RecordSource = ("Select * from Products order by ProductName")
Data1.Refresh
ButtonEnabled1
Exit Sub
SaveErr:
MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub text9_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Public Sub ButtonEnabled()
Text1.Enabled = True
Text2.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = True
cmd4.Enabled = True
cmd6.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
comSearch.Enabled = False
Command1.Enabled = False
Text1.Locked = True
End Sub

Public Sub ButtonEnabled1()
Text1.Enabled = False
Text2.Enabled = False
Text5.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = False
cmd4.Enabled = False
cmd6.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
comSearch.Enabled = True
Command1.Enabled = True
End Sub

