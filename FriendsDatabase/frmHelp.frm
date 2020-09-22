VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Friends Database"
   ClientHeight    =   3330
   ClientLeft      =   3000
   ClientTop       =   885
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFrame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
      Begin VB.Label lblLabel1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   40
         TabIndex        =   2
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.TextBox txtText20 
      Enabled         =   0   'False
      Height          =   3330
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   0
      Width           =   6630
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblLabel1_Click()
    
    Unload Me
    
End Sub

