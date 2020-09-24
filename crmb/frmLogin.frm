VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2505
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3795
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1480.036
   ScaleMode       =   0  'User
   ScaleWidth      =   3563.3
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtUserid 
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
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   1260
   End
   Begin VB.TextBox txtUserid1 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Password"
      Top             =   2520
      Width           =   2355
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "user"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "password"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   25
      Left            =   120
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "usiness Software Solution"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1092
      TabIndex        =   11
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CRMB"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User:"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Adologin As Recordset


Private Sub cmdOK_Click()

    Dim found As Boolean
    found = False
    txtPassword.Text = LCase(txtPassword.Text)
  
    If txtPassword.Text = "" Or txtUserid.Text = "" Then
        Beep
        MsgBox "One of your login fields is blank, please try again.", vbInformation, "Attention"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    Adologin.MoveFirst
    usercode = Adologin.Fields("User").Value
    Do Until found Or Adologin.EOF
        usercode = Adologin.Fields("User").Value
        If usercode = txtUserid.Text Then
            found = True
            files = Val(Adologin!files)
            transact = Val(Adologin!Transactions)
            reports = Val(Adologin!reports)
            usersetup = Val(Adologin!usersetup)
            database = Val(Adologin!database)
            Exit Do
        Else
            Adologin.MoveNext
        End If
    Loop
        
    If found Then
        password = Adologin.Fields("password").Value
        If password = txtPassword.Text Then
            'frmMDI.Show
            'Unload Me
        Else
            Beep
            MsgBox "Password is incorrect, please try again.", vbExclamation, "Attention"
            txtPassword.SetFocus
            Exit Sub
        End If
    Else
        Beep
        MsgBox "User was not found, please try again.", vbExclamation, "Attention"
        txtUserid.SetFocus
    End If
    
            
            
                If found Then
                    Beep
                    MsgBox "Access Granted to " & txtUserid.Text, vbInformation, "Login"
                    Load frmMDI
                   frmMDI.Show
                   Unload Me
                End If

End Sub


Private Sub Form_Load()
 Dim X As Integer
 CenterForm Me
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\crmdb.mdb;"
  Set Adologin = New Recordset
     Adologin.Open "select User,password,files,transactions,reports,usersetup,database from users Order by User", db, adOpenStatic, adLockOptimistic
  
 For X = 1 To Adologin.RecordCount
    txtUserid.AddItem Adologin.Fields("User")
    Adologin.MoveNext
Next X
txtUserid.ListIndex = 0
End Sub
Private Sub cmdCancel_Click()
Beep
ans = MsgBox("Access Failed.", vbExclamation, "Login ")
End
End Sub


Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim found As Boolean
    found = False
    txtPassword.Text = LCase(txtPassword.Text)
    txtUserid.Text = LCase(txtUserid.Text)
    If txtPassword.Text = "" Or txtUserid.Text = "" Then
        Beep
        MsgBox "One of your login fields is blank, please try again.", vbInformation, "Attention"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    Adologin.MoveFirst
    usercode = Adologin.Fields("User").Value
    Do Until found Or Adologin.EOF
        usercode = Adologin.Fields("User").Value
        If usercode = txtUserid.Text Then
            found = True
            user = usercode
            files = Val(Adologin!files)
            transact = Val(Adologin!Transactions)
            reports = Val(Adologin!reports)
            usersetup = Val(Adologin!usersetup)
            database = Val(Adologin!database)
            Exit Do
        Else
            Adologin.MoveNext
        End If
    Loop
        
    If found Then
        password = Adologin.Fields("password").Value
        If password = txtPassword.Text Then
            'frmMDI.Show
            'Unload Me
        Else
            Beep
            MsgBox "Password is incorrect, please try again.", vbExclamation, "Attention"
            txtPassword.SetFocus
            Exit Sub
        End If
    Else
        Beep
        MsgBox "User was not found, please try again.", vbExclamation, "Attention"
        txtUserid.SetFocus
    End If
    
              
    
            
                If found Then
                    Beep
                  MsgBox "Access Granted to " & txtUserid.Text, vbInformation, "Login"
                  Load frmMDI
                   frmMDI.Show
                   Unload Me
                End If
End If
End Sub


Private Sub txtUserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
SendKeys "{home}+{end}"
txtUserid.Text = LCase(txtUserid.Text)
End If
End Sub

