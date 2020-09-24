Attribute VB_Name = "Module1"
Public files As Integer
Public usersetup As Integer
Public reports As Integer
Public transact As Integer
Public database As Integer
Public usercode As String
Public db As Connection

Public Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height * 0.85) \ 2 - frm.Height \ 2
    frm.Left = Screen.Width \ 2 - frm.Width \ 2
End Sub
