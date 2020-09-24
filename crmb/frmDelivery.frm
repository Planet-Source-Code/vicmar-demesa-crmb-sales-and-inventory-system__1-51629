VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   7590
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   8655
   ControlBox      =   0   'False
   Icon            =   "frmDelivery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd1 
      Caption         =   "&New"
      Height          =   375
      Left            =   7080
      Picture         =   "frmDelivery.frx":000C
      TabIndex        =   50
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   44
      Top             =   600
      Width           =   6735
      Begin VB.TextBox Text5 
         DataField       =   "Date"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         DataField       =   "DeliveryNo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         DataField       =   "DeliveryNo"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   1560
         TabIndex        =   49
         Text            =   "Text14"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Delivery No."
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         DataSource      =   "Data1"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   47
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   8160
      TabIndex        =   42
      Top             =   8760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check Product"
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
      Height          =   735
      Left            =   240
      TabIndex        =   37
      Top             =   2520
      Width           =   6735
      Begin VB.CommandButton cmd5 
         Caption         =   "Check"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   41
         Top             =   240
         Width           =   1065
      End
      Begin VB.ComboBox comSearch 
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
         ItemData        =   "frmDelivery.frx":01A9
         Left            =   720
         List            =   "frmDelivery.frx":01AB
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox Text16 
      DataField       =   "TotalPrice"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   480
      TabIndex        =   34
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   2640
      TabIndex        =   33
      Text            =   "Text17"
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmd4 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9915
      Width           =   1300
   End
   Begin VB.CommandButton cmd6 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9630
      Width           =   1300
   End
   Begin VB.CommandButton cmd2 
      Appearance      =   0  'Flat
      Caption         =   "&Modify"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9120
      Width           =   1300
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDelivery.frx":01AD
      Height          =   2055
      Left            =   360
      OleObjectBlob   =   "frmDelivery.frx":01C1
      TabIndex        =   23
      Top             =   4920
      Width           =   6495
   End
   Begin VB.Frame Frame3 
      Caption         =   "Purchase Order No."
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
      Height          =   2295
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   6615
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      RecordSource    =   "DeliverDetails"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Products"
      Top             =   8925
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Deliver"
      Top             =   9720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ctr"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox ctr3 
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   9180
      Width           =   255
   End
   Begin VB.TextBox ctr2 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   9240
      Width           =   255
   End
   Begin VB.TextBox ctr1 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   9240
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "autorec"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1860
      TabIndex        =   18
      Top             =   9060
      Width           =   1095
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DeliverDetails"
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Receive Item"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Picture         =   "frmDelivery.frx":0F0C
      TabIndex        =   1
      Top             =   3120
      Width           =   1425
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&Update Record"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Picture         =   "frmDelivery.frx":17D6
      TabIndex        =   2
      Top             =   3600
      Width           =   1425
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "Close"
      Height          =   375
      Left            =   7080
      Picture         =   "frmDelivery.frx":18D8
      TabIndex        =   17
      Top             =   4080
      Width           =   1425
   End
   Begin VB.Frame Frame4 
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
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   6735
      Begin VB.TextBox Text8 
         DataField       =   "UnitPrice"
         DataSource      =   "Data3"
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
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         DataField       =   "ProductCode"
         DataSource      =   "Data3"
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
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         DataField       =   "UnitsInStock"
         DataSource      =   "Data3"
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
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         DataField       =   "ProductName"
         DataSource      =   "Data3"
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
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text13 
         DataField       =   "ItemID"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Text            =   "Text13"
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         DataField       =   "Quantity"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   4800
         TabIndex        =   30
         Text            =   "Text11"
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Price"
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Product Code"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Units in Stock"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   12
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1395
      End
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
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6735
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frmDelivery.frx":1A75
         Left            =   1560
         List            =   "frmDelivery.frx":1A77
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox Combo4 
         DataField       =   "SupplierID"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmDelivery.frx":1A79
         Left            =   1560
         List            =   "frmDelivery.frx":1A7B
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1800
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text11 
      DataField       =   "UnitPrice"
      DataSource      =   "Data4"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3720
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   240
      Top             =   360
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "   Deliveries"
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
      TabIndex        =   43
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Amount"
      Height          =   195
      Left            =   4320
      TabIndex        =   36
      Top             =   7200
      Width           =   945
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   2175
      Index           =   0
      Left            =   420
      TabIndex        =   25
      Top             =   4860
      Width           =   6465
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   1815
      Index           =   3
      Left            =   420
      TabIndex        =   15
      Top             =   2820
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   1515
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   6675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
      DataSource      =   "Data1"
      Height          =   195
      Index           =   9
      Left            =   2760
      TabIndex        =   32
      Top             =   6480
      Width           =   585
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoSupplier As Recordset
Attribute AdoSupplier.VB_VarHelpID = -1
Dim AdoProduct As Recordset
Private Sub cmd1_Click()
On Error GoTo AddErr
If Text1.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
Text6.Text = Val(Text6.Text) + 1
Text6.Text = Format$(Text6.Text, "0000000")
Data1.Recordset.AddNew
Text1.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo2.Visible = True
Combo4.Visible = False
Combo2.Enabled = True
Combo4.Enabled = True
Text15.Text = 0
Text5.Text = Date
Text1.Text = Text6.Text
Text2.SetFocus
End If
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = False
cmd7.Enabled = False
Command1.Enabled = True
comSearch.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
Exit Sub
AddErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub


Private Sub cmd5_Click()
On Error GoTo SearchErr
Dim a

If comSearch.Text = "Product Code" Then
a = Combo5.Text
Data3.RecordSource = ("Select * from Products where ProductCode = '" & a & "'")
Data3.Refresh
If Text2.Text = "" Then
a = Combo5.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data3.RecordSource = ("Select * from Products order by ProductCode")
Data3.Refresh
Else

End If
Else
If comSearch.Text = "Product Name" Then
a = Combo6.Text
Data3.RecordSource = ("Select * from Products where ProductName = '" & a & "'")
Data3.Refresh
If Text2.Text = "" Then
a = Combo6.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data3.RecordSource = ("Select * from Products order by ProductName")
Data3.Refresh
Else

End If
End If
End If
cmd5.Enabled = True
Exit Sub
SearchErr:
MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub



Private Sub cmd7_Click()
Dim a
a = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Confirm")
If a = vbYes Then
Unload Me
End If
End Sub
Private Sub Command1_Click()
On Error GoTo SaveErr
Dim a
Dim b
Data4.Recordset.AddNew
b = InputBox("Enter Number of Quantity:", "Quantity", "")
If b = "" Then
Exit Sub
End If
If Text19.Text = Text10.Text Then
MsgBox "Product is already received, please try again.", vbExclamation, "Attention"
Exit Sub
End If
Data3.Recordset.Edit
Text9.Text = Val(Text9.Text) + b
Text19.Text = Text10.Text
Data3.Recordset.Update
Text4.Text = b
Text14.Text = Text1.Text
Text11.Text = Text8.Text
Text13.Text = Text10.Text
Text17.Text = Val(Text15.Text)
Text16.Text = b * Val(Text11.Text)
Text15.Text = Val(Text16.Text) + Val(Text17.Text)
If Text11.Text = "" Or Text4.Text = "" Or Text14.Text = "" Or Text13.Text = "" Then
MsgBox "Please enter Complete information about the Product", vbInformation, "Confirm"
Else
Data4.Recordset.Update
Data4.Refresh
cmd1.Visible = True
cmd1.Enabled = False
cmd3.Enabled = True
cmd5.Enabled = True
End If
Exit Sub
SaveErr:
MsgBox Err.Description & "", vbExclamation, "Attention"
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

Private Sub Form_Load()

Dim X As Integer
Dim a
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\crmdb.mdb;"
  Set AdoSupplier = New Recordset
     AdoSupplier.Open "select Supplierid,Companyname from Suppliers Order by Supplierid", db, adOpenStatic, adLockOptimistic
  Set AdoProduct = New Recordset
     AdoProduct.Open "select ProductCode,Productname from Products Order by ProductCode", db, adOpenStatic, adLockOptimistic
  For X = 1 To AdoSupplier.RecordCount
    Combo2.AddItem AdoSupplier.Fields("Companyname")
    AdoSupplier.MoveNext
Next X
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
Data3.DatabaseName = App.Path + "\crmdb.mdb"
Data4.DatabaseName = App.Path + "\crmdb.mdb"
Data5.DatabaseName = App.Path + "\crmdb.mdb"


End Sub




Private Sub cmd3_Click()

On Error GoTo SaveErr
Dim a
If Text1.Enabled = False Then
MsgBox "Please add/edit before saving.", vbExclamation, "Attention"
Exit Sub
End If

Text5.Text = Date
Combo4.Text = Combo2.Text
Text14.Text = Text1.Text
If Text2.Text = "" Then
MsgBox "Please enter the invoice number.", vbExclamation, "Attention"
Text2.SetFocus
Exit Sub
ElseIf Text12.Text = "" Then
MsgBox "Please enter the purchase order number.", vbExclamation, "Attention"
Text12.SetFocus
Exit Sub
ElseIf Combo2.Text = "" Then
MsgBox "Please enter the supplier to received product.", vbExclamation, "Attention"
Combo2.SetFocus
Exit Sub
End If
MsgBox "Received product successfully updated.", vbInformation, "Update"
Data1.Recordset.Update
Data1.Refresh
Text1.Enabled = False
Text2.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text12.Enabled = False
Combo2.Visible = False
Combo4.Visible = True
Combo2.Enabled = False
Combo4.Enabled = False
cmd7.Enabled = True
cmd1.Enabled = True
cmd3.Enabled = False
cmd5.Enabled = False
Command1.Enabled = False
Exit Sub
SaveErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub
Private Sub text12_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
   Beep
   KeyAscii = 0
End If
End Sub





