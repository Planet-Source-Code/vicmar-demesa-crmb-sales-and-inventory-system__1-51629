VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   "
   ClientHeight    =   7485
   ClientLeft      =   210
   ClientTop       =   495
   ClientWidth     =   8430
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
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8430
   Begin VB.CommandButton cmd1 
      Caption         =   "&New"
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
      Left            =   6960
      Picture         =   "frmOrder.frx":000C
      TabIndex        =   51
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   45
      Top             =   600
      Width           =   6735
      Begin VB.TextBox Text1 
         DataField       =   "OrderNo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         DataField       =   "Date"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text12 
         DataField       =   "OrderNo"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   1320
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
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
         Left            =   4200
         TabIndex        =   49
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Order No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   6360
      TabIndex        =   43
      Top             =   10440
      Width           =   975
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Check Product"
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   6735
      Begin VB.ComboBox Combo5 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox Combo6 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox comSearch 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmOrder.frx":01A9
         Left            =   840
         List            =   "frmOrder.frx":01AB
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "Check"
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
         Height          =   315
         Left            =   5520
         TabIndex        =   38
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   1200
      TabIndex        =   35
      Text            =   "Text18"
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   4200
      TabIndex        =   34
      Text            =   "Text17"
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      DataField       =   "TotalPrice"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   2640
      TabIndex        =   33
      Top             =   8880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Order"
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
      Left            =   6960
      Picture         =   "frmOrder.frx":01AD
      TabIndex        =   31
      Top             =   2880
      Width           =   1425
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&Update Record"
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
      Left            =   6960
      Picture         =   "frmOrder.frx":0A77
      TabIndex        =   30
      Top             =   3360
      Width           =   1425
   End
   Begin VB.CommandButton cmd4 
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
      Left            =   1200
      Picture         =   "frmOrder.frx":0B79
      TabIndex        =   29
      Top             =   10560
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmd6 
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
      Left            =   -120
      Picture         =   "frmOrder.frx":0C7B
      TabIndex        =   28
      Top             =   10320
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&Edit"
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
      Left            =   600
      Picture         =   "frmOrder.frx":0D7D
      TabIndex        =   27
      Top             =   10080
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmd7 
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
      Left            =   6960
      Picture         =   "frmOrder.frx":11BF
      TabIndex        =   26
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6735
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmOrder.frx":135C
         Left            =   1320
         List            =   "frmOrder.frx":135E
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox Combo4 
         DataField       =   "CustomerCode"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "frmOrder.frx":1360
         Left            =   1320
         List            =   "frmOrder.frx":1362
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Customer"
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OrderDetails"
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      ForeColor       =   &H00404000&
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   6735
      Begin VB.TextBox Text3 
         DataField       =   "ProductCode"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         DataField       =   "ProductName"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         DataField       =   "UnitsInStock"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         DataField       =   "UnitPrice"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         DataField       =   "ProductCode"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Text            =   "Text13"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         DataField       =   "UnitPrice"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   4800
         TabIndex        =   19
         Text            =   "Text11"
         Top             =   840
         Width           =   1455
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
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   840
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
         Index           =   7
         Left            =   3720
         TabIndex        =   16
         Top             =   840
         Width           =   1395
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
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1395
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
         Index           =   4
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   1395
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmOrder.frx":1364
      Height          =   2130
      Left            =   120
      OleObjectBlob   =   "frmOrder.frx":1378
      TabIndex        =   9
      Top             =   4800
      Width           =   6735
   End
   Begin VB.TextBox Text6 
      DataField       =   "autoissue"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Products"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1860
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
      Left            =   4800
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   8400
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
      Left            =   2400
      TabIndex        =   5
      Top             =   10080
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
      Left            =   3240
      TabIndex        =   4
      Top             =   10080
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
      Left            =   3600
      TabIndex        =   3
      Top             =   10020
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
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ctr"
      Top             =   10920
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
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Order"
      Top             =   10560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2040
      Top             =   10200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data4 
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
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OrderDetails"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text14 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   23
      Top             =   10320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "Quantity"
      DataSource      =   "Data4"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Order No. "
      ForeColor       =   &H00404040&
      Height          =   2295
      Left            =   0
      TabIndex        =   24
      Top             =   4560
      Width           =   6855
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
      Caption         =   "   Orders"
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
      TabIndex        =   44
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   32
      Top             =   7080
      Width           =   945
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   2100
      Index           =   1
      Left            =   180
      TabIndex        =   25
      Top             =   4860
      Width           =   6705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   22
      Top             =   9600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Label5"
      Height          =   2055
      Index           =   3
      Left            =   180
      TabIndex        =   13
      Top             =   2460
      Width           =   6735
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
      Height          =   1275
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6675
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoCustomer As Recordset
Attribute AdoCustomer.VB_VarHelpID = -1
Dim AdoProduct As Recordset
Private Sub cmd1_Click()
On Error GoTo AddErr
If Text1.Enabled = True Then
MsgBox "Please save first before adding", vbExclamation, "Attention"
Else
Text6.Text = Val(Text6.Text) + 1
Text6.Text = Format$(Text6.Text, "000000")
Data1.Recordset.AddNew
Text1.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo2.Visible = True
Combo4.Visible = False
Combo2.Enabled = True
Combo4.Enabled = True
Text15.Text = 0
Text19.Text = ""
Text5.Text = Date
Text1.Text = Text6.Text
Text18.Text = Text1.Text
Combo2.SetFocus
End If
cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = False
Command1.Enabled = True
comSearch.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
cmd7.Enabled = False
a = Text1.Text
  Data5.RecordSource = ("Select * from OrderDetails where OrderNo = '" & a & "'")
Data5.Refresh
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
If Text3.Text = "" Then
a = Combo5.Text
MsgBox "Product not found, please try again.", vbExclamation, "Attention"
Data3.RecordSource = ("Select * from Products order by ProductCode")
Data3.Refresh
Else
End If
Else
If comSearch.Text = "Product Name" Then
a = Combo6.Text
Data3.RecordSource = ("Select * from Products where ProductName = '" & a & "'")
Data3.Refresh
If Text3.Text = "" Then
a = Combo6.Text
MsgBox "Record not found, please try again.", vbExclamation, "Attention"
Data3.RecordSource = ("Select * from Products order by ProductName")
Data3.Refresh
Else
End If
End If
End If
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
Dim a, b
b = InputBox("Enter Number of Quantity:", "Quantity", "")
If b = "" Or Val(b) = "0" Then
MsgBox "Enter number of quantity to ordered, please try again.", vbExclamation, "Attention"
Exit Sub
End If
If b > Val(Text9.Text) Then
MsgBox "Quantity of products to be ordered exceeds the units in stock, please try again.", vbExclamation, "Attention"
Command1.SetFocus
Exit Sub
End If
b = Val(b)
Data3.Recordset.Edit
Text9.Text = Val(Text9.Text) - b
Text19.Text = Text10.Text
Data3.Recordset.Update
Data4.Recordset.AddNew
Text14.Text = b
Text4.Text = b
Text12.Text = Text18.Text
Text11.Text = Text8.Text
Text13.Text = Text10.Text
Text17.Text = Val(Text15.Text)
Text16.Text = b * Val(Text11.Text)
Text15.Text = Val(Text16.Text) + Val(Text17.Text)
If Text11.Text = "" Or Text4.Text = "" Or Text12.Text = "" Or Text13.Text = "" Then
MsgBox "Please enter Complete information about the Product", vbInformation, "Confirm"
Else
Data4.Recordset.Update
Data4.Refresh
cmd1.Visible = True
cmd1.Enabled = False
cmd3.Enabled = True
cmd5.Enabled = True
End If
a = Text1.Text
  Data5.RecordSource = ("Select * from OrderDetails where OrderNo = '" & a & "'")
Data5.Refresh
Text14.Text = ""
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
  Set AdoCustomer = New Recordset
     AdoCustomer.Open "select customercode,firstname from customers Order by customercode", db, adOpenStatic, adLockOptimistic
  Set AdoProduct = New Recordset
     AdoProduct.Open "select ProductCode,Productname from Products Order by ProductCode", db, adOpenStatic, adLockOptimistic
  For X = 1 To AdoCustomer.RecordCount
    Combo2.AddItem AdoCustomer.Fields("CustomerCode")
    AdoCustomer.MoveNext
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
  a = "0000001"
Data5.RecordSource = ("Select * from OrderDetails where OrderNo = '" & a & "'")
Data5.Refresh
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
If Combo2.Text = "" Then
MsgBox "Please enter the customer who order.", vbExclamation, "Attention"
Combo2.SetFocus
Exit Sub
End If
MsgBox "Ordered product successfully updated.", vbInformation, "Update"
Data1.Recordset.Update
Data1.Refresh
Text1.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Combo2.Visible = False
Combo4.Visible = True
Combo2.Enabled = False
Combo4.Enabled = False
cmd1.Enabled = True
cmd3.Enabled = False
cmd5.Enabled = False
cmd7.Enabled = True
Command1.Enabled = False
Exit Sub
SaveErr:
  MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub

