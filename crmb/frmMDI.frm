VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00808080&
   Caption         =   "CRMBusiness Software Solution - Sales System"
   ClientHeight    =   7800
   ClientLeft      =   255
   ClientTop       =   705
   ClientWidth     =   10125
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbr_Status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7500
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   18230
            MinWidth        =   18228
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "4/17/03"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:01 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3360
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":16FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1816
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1906
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":275A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":3936
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":4F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":5E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":74CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":83A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9A02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   1058
      ButtonWidth     =   820
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Process Orders..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Accept Deliveries..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Product Database..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Customer Database"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Supplier Database..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Sales Report..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Inventory Report..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Change Password..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User Setup..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About CRMBSS..."
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuoff 
         Caption         =   "&Log off..."
      End
      Begin VB.Menu mnufilespacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuorder 
         Caption         =   "&Orders..."
      End
      Begin VB.Menu mnuords 
         Caption         =   "-"
      End
      Begin VB.Menu mnudelivery 
         Caption         =   "Deliveries..."
      End
   End
   Begin VB.Menu mnudb 
      Caption         =   "&Database Management"
      Begin VB.Menu mnuproduct 
         Caption         =   "Product Database..."
      End
      Begin VB.Menu tr 
         Caption         =   "-"
      End
      Begin VB.Menu mnucustomer 
         Caption         =   "Customer Database..."
      End
      Begin VB.Menu tyu 
         Caption         =   "-"
      End
      Begin VB.Menu mnusupplier 
         Caption         =   "Supplier Database..."
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnusales 
         Caption         =   "Sales Report..."
      End
      Begin VB.Menu oi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinventory 
         Caption         =   "Inventory Report"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuchpass 
         Caption         =   "&Change Password..."
      End
      Begin VB.Menu uiy 
         Caption         =   "-"
      End
      Begin VB.Menu mnunewuser 
         Caption         =   "&User Setup..."
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoInventory As Recordset
Dim adoSales As Recordset
Private Sub MDIForm_Load()
frmMDI.sbr_Status.Panels(1).Text = "User: " & frmLogin.txtUserid.Text
frmMDI.sbr_Status.Panels(2).Text = "Welcome to CRMBusiness Software Solution Program Copyright(c) 2003 Developed by SCaVeNGeR."
frmMDI.mnuoff.Caption = "Log off " & frmLogin.txtUserid.Text + "..."

If database = 0 Then
mnuproduct.Enabled = False
mnucustomer.Enabled = False
mnusupplier.Enabled = False
Toolbar1.Buttons.Item(4).Enabled = False
Toolbar1.Buttons.Item(5).Enabled = False
Toolbar1.Buttons.Item(6).Enabled = False
End If

If transact = 0 Then
mnuorder.Enabled = False
mnudelivery.Enabled = False
Toolbar1.Buttons.Item(1).Enabled = False
Toolbar1.Buttons.Item(2).Enabled = False
End If

If reports = 0 Then
mnusales.Enabled = False
mnuinventory.Enabled = False
Toolbar1.Buttons.Item(8).Enabled = False
Toolbar1.Buttons.Item(9).Enabled = False
End If

If usersetup = 0 Then mnunewuser.Enabled = False

If usersetup = 0 Then Toolbar1.Buttons.Item(12).Enabled = False
End Sub






Private Sub MDIForm_Terminate()
End
End Sub

Private Sub mnuabout_Click()
Load frmAbout
frmAbout.Show 1
End Sub


Private Sub mnuchpass_Click()
frmChangepassword.Show
End Sub


Private Sub mnusales_Click()
Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\crmdb.mdb;"

  Set adoSales = New Recordset
     adoSales.Open "select * from OrderDetails", db, adOpenStatic, adLockOptimistic
   
Set SalesReport.DataSource = adoSales
SalesReport.Show
End Sub

Private Sub mnucustomer_Click()
frmCustomer.Show
End Sub

Private Sub mnudelivery_Click()
frmDeliver.Show
End Sub

Private Sub mnuexit_Click()
ans = MsgBox("Are you sure you want to exit this application?", vbYesNo + vbQuestion, "Exit ")
        If ans = vbYes Then
            End
        End If
End Sub

Private Sub mnuinventory_Click()
Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\crmdb.mdb;"

  Set adoInventory = New Recordset
  adoInventory.Open "select ProductCode,ProductName,UnitPrice,UnitsInStock,reorderLevel,(unitsinstock*unitprice) as [total] from Products order by ProductName", db, adOpenStatic, adLockOptimistic
  
Set InventoryReport.DataSource = adoInventory

InventoryReport.Show
End Sub


Private Sub mnunewuser_Click()
frmUserSetup.Show
End Sub

Private Sub mnuoff_Click()
 ans = MsgBox("Are you sure you want to log off?", vbYesNo + vbQuestion, "Log Off ")
        If ans = vbYes Then
            Unload Me
            frmLogin.Show
        End If

End Sub

Private Sub mnuorder_Click()
frmOrder.Show
End Sub

Private Sub mnuproduct_Click()
frmProduct.Show
End Sub


Private Sub mnusupplier_Click()
frmSupplier.Show
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
             Call mnuorder_Click
        Case 2:
            Call mnudelivery_Click
        Case 4:
            Call mnuproduct_Click
        Case 5:
            Call mnucustomer_Click
        Case 6:
            Call mnusupplier_Click
        Case 8:
            Call mnusales_Click
        Case 9:
            Call mnuinventory_Click
        Case 11:
            Call mnuchpass_Click
        Case 12:
            Call mnunewuser_Click
        Case 14:
            Call mnuabout_Click
               
           End Select
End Sub
