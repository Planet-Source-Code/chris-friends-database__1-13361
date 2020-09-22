VERSION 5.00
Begin VB.Form FrmChangePassword 
   BackColor       =   &H00000000&
   Caption         =   "Change Password"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   Icon            =   "FrmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4395
   Begin VB.Frame FrChangePassword 
      BackColor       =   &H00000000&
      Caption         =   "Change Password"
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtExistingPassword 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtNewPassword2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Enter Existing Password:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Enter New Password:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Confirm New Password:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.Frame fraFrame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Cancel"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   10
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
      Begin VB.Label cmdOk 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "&OK"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   8
         Top             =   130
         Width           =   1000
      End
   End
End
Attribute VB_Name = "FrmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    
    FrmChangePassword.Hide
    
End Sub


Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdOk.BackColor = &H0&
    cmdOk.ForeColor = &HFFFF00
    
    cmdCancel.ForeColor = &H0&
    cmdCancel.BackColor = &HFFFF00
    
End Sub

Private Sub cmdOK_Click()

Dim strTemp As String
Dim strPW As String
Dim strNewPW As String
Dim strEncryptNewPW As String
    'some error handling
    
    strPW = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
    strNewPW = LCase(txtNewPassword2.Text)
    'checks to see if you type int he correct password in the existing password field
        
     If FrmLogin.txtPassword = LCase(txtExistingPassword.Text) Then
        'checks the match of the new passwords
        
        If LCase(txtNewPassword1.Text) = strNewPW Then
            strEncryptNewPW = Encrypt(strNewPW)
            PutValue "Main", "Password", strEncryptNewPW, App.Path & "\" & con_INI_File
            MsgBox "Password changed!", 8, "Password Verfication"
        
        Else
            MsgBox "The New Passwords Do Not Match", 8, "Password Error"
            txtNewPassword1.SetFocus
            Exit Sub
        
        End If
        
    Else
        MsgBox "The Existing Password is Incorrect!", 8, "Password Error"
        txtExistingPassword.SetFocus
        Exit Sub
        
    End If
    'if the existing password matches the decrypted password and
    'both the new passwords match, then it changes the password to
    'be encrypted in the ini file (and then hides the change
    'password dialog box)
    
    FrmChangePassword.Hide
    DoEvents
    
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdOk.BackColor = &HFFFF00
    cmdOk.ForeColor = &H0&
    
    cmdCancel.BackColor = &H0&
    cmdCancel.ForeColor = &HFFFF00

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdOk.ForeColor = &HFFFF00
    cmdOk.BackColor = &H0&
    
    cmdCancel.ForeColor = &HFFFF00
    cmdCancel.BackColor = &H0&
    
End Sub

Private Sub FrChangePassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdOk.ForeColor = &HFFFF00
    cmdOk.BackColor = &H0&
    
    cmdCancel.ForeColor = &HFFFF00
    cmdCancel.BackColor = &H0&

End Sub
