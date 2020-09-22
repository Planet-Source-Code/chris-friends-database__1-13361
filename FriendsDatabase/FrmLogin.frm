VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Protected"
   ClientHeight    =   1290
   ClientLeft      =   2205
   ClientTop       =   2715
   ClientWidth     =   4440
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer2 
      Interval        =   3000
      Left            =   360
      Top             =   240
   End
   Begin Friends.CoolPgBar CoolPgBar2 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Please wait..."
      CaptionStyle    =   2
   End
   Begin VB.Timer tmrTimer1 
      Interval        =   53
      Left            =   0
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Enter Password"
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   720
         MouseIcon       =   "FrmLogin.frx":030A
         Picture         =   "FrmLogin.frx":0614
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Password:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   400
         Width           =   855
      End
   End
   Begin VB.Frame fraFrame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   1095
      Begin VB.Label cmdSubmit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "&OK"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   130
         Width           =   1005
      End
   End
   Begin VB.Frame fraFrame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   1095
      Begin VB.Label cmdCancel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   7
         Top             =   130
         Width           =   1000
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean
Private Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Private Sub cmdCancel_Click()
        
    PlaySound App.Path & "\sounds\groupopen.wav"
        
    ReadyToClose = True
    Unload Me
    End
    
End Sub



Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdCancel.BackColor = &HFFFF00
    cmdCancel.ForeColor = &H0&
    
    cmdSubmit.ForeColor = &HFFFF00
    cmdSubmit.BackColor = &H0&
    
    
End Sub

Private Sub cmdSubmit_Click()

    PlaySound App.Path & "\sounds\groupopen.wav"

    Dim strTest As String
    strTest = GetValue("Main", "Password", App.Path & "\" & con_INI_File)
   
  If LCase(txtPassword.Text) = Decrypt(strTest) Then
        
    frmFriends.Show
        ' The name of the main application
    FrmLogin.Hide
        ' Hides the login dialog box
        
  Else
    MsgBox "Enter a Valid Password for this System", 8, "Password Error"
    txtPassword.SetFocus
  Exit Sub
        
  End If
    
End Sub

Private Sub cmdSubmit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdSubmit.BackColor = &HFFFF00
    cmdSubmit.ForeColor = &H0&
    
    cmdCancel.BackColor = &H0&
    cmdCancel.ForeColor = &HFFFF00

End Sub

Private Sub coolpgBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Form_Load()
    
    PlaySound App.Path & "\sounds\start.wav"
    
    ' Remove the Close system menu item and the
    ' menu separator.
    RemoveMenus Me, False, False, _
    False, False, False, True, True
        
        
  If App.PrevInstance = True Then 'this will insure that your program cannot be opened more then once
   MsgBox "This program is already running !", vbCritical, "Program Error"
End
  End If
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = Not ReadyToClose
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdSubmit.ForeColor = &HFFFF00
    cmdSubmit.BackColor = &H0&
    
    cmdCancel.ForeColor = &HFFFF00
    cmdCancel.BackColor = &H0&
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End
    
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdSubmit.ForeColor = &HFFFF00
    cmdSubmit.BackColor = &H0&
    
    cmdCancel.ForeColor = &HFFFF00
    cmdCancel.BackColor = &H0&

End Sub

Private Sub tmrTimer1_Timer()


CoolPgBar2.Value = CoolPgBar2.Value + 2
  
     
End Sub

Private Sub tmrTimer2_Timer()

  On Error Resume Next
         
    txtPassword.Enabled = True
     
    cmdCancel.Enabled = True
    cmdSubmit.Enabled = True
          
    txtPassword.SetFocus
    CoolPgBar2.Visible = False

Resume Next

End Sub
