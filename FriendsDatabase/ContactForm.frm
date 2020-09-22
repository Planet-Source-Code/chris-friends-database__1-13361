VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLSTBAR.OCX"
Begin VB.Form frmFriends 
   BackColor       =   &H00000000&
   Caption         =   "Friends Database"
   ClientHeight    =   8205
   ClientLeft      =   -105
   ClientTop       =   225
   ClientWidth     =   11955
   ForeColor       =   &H00000000&
   Icon            =   "ContactForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtText13 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4530
      Left            =   1040
      TabIndex        =   0
      Top             =   30
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   7990
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Address"
         Text            =   "Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Email"
         Text            =   "Email Address"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Phone"
         Text            =   "Phone Number"
         Object.Width           =   2338
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Cell"
         Text            =   "Cell Number"
         Object.Width           =   2338
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Beeper"
         Text            =   "Beeper Number"
         Object.Width           =   2338
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ICQ Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "AIM Handle"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "IRC Handle"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Gaming Handle"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "HomePage URL"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Sex"
         Object.Width           =   1766
      EndProperty
   End
   Begin VB.Frame fraScheduler 
      BackColor       =   &H00000000&
      Caption         =   "Memo"
      ForeColor       =   &H00FFFF00&
      Height          =   1800
      Left            =   1020
      TabIndex        =   41
      Top             =   4560
      Width           =   10860
      Begin VB.TextBox txtText7 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   10575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   15
      Top             =   7935
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   476
      SimpleText      =   "Caps"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   1695
            MinWidth        =   1060
            Text            =   "Caps Lock"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   1694
            MinWidth        =   1059
            Text            =   "Num Lock"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3457
            MinWidth        =   2822
            Text            =   "Christopher Palladino"
            TextSave        =   "Christopher Palladino"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   9456
            MinWidth        =   8821
            TextSave        =   "12/7/00"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4163
            MinWidth        =   3528
            TextSave        =   "10:49 AM"
         EndProperty
      EndProperty
   End
   Begin Listbar.SSListBar SSListBar1 
      Height          =   6660
      Left            =   45
      TabIndex        =   14
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   11748
      _Version        =   65537
      BackColor       =   0
      BorderStyle     =   2
      CaptionBackColor=   12632256
      CaptionForeColor=   0
      OLEDragMode     =   1
      OLEDropMode     =   2
      IconsLargeCount =   6
      Image(1).Index  =   1
      Image(1).Picture=   "ContactForm.frx":08CA
      Image(2).Index  =   2
      Image(2).Picture=   "ContactForm.frx":0BE6
      Image(3).Index  =   3
      Image(3).Picture=   "ContactForm.frx":14C2
      Image(4).Index  =   4
      Image(4).Picture=   "ContactForm.frx":1D9E
      Image(5).Index  =   5
      Image(5).Picture=   "ContactForm.frx":267A
      Image(6).Index  =   6
      Image(6).Picture=   "ContactForm.frx":2F56
      Groups(1).ItemCount=   6
      Groups(1).BackColor=   0
      Groups(1).ForeColor=   16776960
      Groups(1).PictureBackgroundMaskColor=   0
      Groups(1).CurrentGroup=   -1  'True
      Groups(1).Caption=   "Controls"
      Groups(1).ListItems(1).Text=   "Change Password"
      Groups(1).ListItems(1).Key=   "change"
      Groups(1).ListItems(1).IconLarge=   1
      Groups(1).ListItems(2).Index=   2
      Groups(1).ListItems(2).Text=   "Show Controls"
      Groups(1).ListItems(2).Key=   "show"
      Groups(1).ListItems(2).IconLarge=   2
      Groups(1).ListItems(3).Index=   3
      Groups(1).ListItems(3).Text=   "Hide Controls"
      Groups(1).ListItems(3).Key=   "hide"
      Groups(1).ListItems(3).IconLarge=   3
      Groups(1).ListItems(4).Index=   4
      Groups(1).ListItems(4).Text=   "Email Contact"
      Groups(1).ListItems(4).Key=   "email"
      Groups(1).ListItems(4).IconLarge=   5
      Groups(1).ListItems(5).Index=   5
      Groups(1).ListItems(5).Text=   "Help File"
      Groups(1).ListItems(5).Key=   "help"
      Groups(1).ListItems(5).IconLarge=   6
      Groups(1).ListItems(6).Index=   6
      Groups(1).ListItems(6).Text=   "Exit"
      Groups(1).ListItems(6).Key=   "exit"
      Groups(1).ListItems(6).IconLarge=   4
   End
   Begin VB.Frame fraFrame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1020
      TabIndex        =   23
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdFirst 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "First Contact"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   24
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdPrevious 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Previous"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   26
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3300
      TabIndex        =   27
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdNext 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Next"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   28
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdLast 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Last Contact"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   30
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5580
      TabIndex        =   31
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdAdd 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Add Contact"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   32
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   33
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdSave 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Save"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   34
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7860
      TabIndex        =   35
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdDelete 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Delete"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   36
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9000
      TabIndex        =   37
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdEdit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Edit Contact"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   38
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10140
      TabIndex        =   39
      Top             =   6300
      Width           =   1095
      Begin VB.Label cmdActivate 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Allow Edit"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   40
         Top             =   130
         Width           =   1000
      End
   End
   Begin VB.Frame fraFrame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11280
      TabIndex        =   42
      Top             =   6300
      Width           =   615
      Begin VB.Label cmdExtra 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Extra"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   43
         Top             =   130
         Width           =   520
      End
   End
   Begin VB.Frame fraFrame1 
      BackColor       =   &H00000000&
      Caption         =   "Contact Information Form"
      ForeColor       =   &H00FFFF00&
      Height          =   1095
      Left            =   50
      TabIndex        =   16
      Top             =   6840
      Width           =   11860
      Begin VB.TextBox txtText6 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtText5 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtText4 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtText3 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtText2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtText1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtText12 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   0
         Left            =   9060
         Locked          =   -1  'True
         MaxLength       =   250
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtText8 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtText9 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtText10 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtText11 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   1
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblLabel6 
         BackColor       =   &H00000000&
         Caption         =   "Beeper Number:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   7800
         TabIndex        =   22
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblLabel5 
         BackColor       =   &H00000000&
         Caption         =   "Cell Number:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   7800
         TabIndex        =   21
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblLabel4 
         BackColor       =   &H00000000&
         Caption         =   "Phone Number:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblLabel3 
         BackColor       =   &H00000000&
         Caption         =   "E-Mail Address:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblLabel2 
         BackColor       =   &H00000000&
         Caption         =   "Address:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblLabel1 
         BackColor       =   &H00000000&
         Caption         =   "Name:"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame fraFrame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11280
      TabIndex        =   44
      Top             =   6300
      Visible         =   0   'False
      Width           =   615
      Begin VB.Label cmdDefault 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Normal"
         ForeColor       =   &H00FFFF00&
         Height          =   200
         Left            =   40
         TabIndex        =   45
         Top             =   130
         Width           =   520
      End
   End
End
Attribute VB_Name = "frmFriends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*******************************************************************************
'*******************************************************************************
'*************************Coded By Christopher Palladino************************
'****************************Thanks To The Following****************************
'*** I'd like to thank Jerry Barnes for his tutorial for absolute beginners, ***
'*** if it wasn't for that this program would probably have been released    ***
'*** much much later. Thanks Jerry i learned a lot from your tutorials it's  ***
'*** people like you that keep people from pulling there hair out when       ***
'*** creating databases. Also I used someone elses source for the login box  ***
'*** but i do not remember who so feel free to e-mail me if you wish for me  ***
'*** to include your name here but many many thanks to you as well. Thanks!! ***
'*** I hope if you downloaded (of course you downloaded it) this source code ***
'*** and find it useful you will add credit to me by voting for my program   ***
'*** cause it took me tedious hours and days to get everything working       ***
'*** correctly. I'm surethere are bugs in it but i tried my best to kill as  ***
'*** many as possible. Good luck with your coding and may God bless you all. ***
'*******************************************************************************
'*******************************************************************************
'*******************************************************************************

Option Explicit
Private WithEvents connConnection As ADODB.Connection
Attribute connConnection.VB_VarHelpID = -1
Private WithEvents rsInfo As ADODB.Recordset
Attribute rsInfo.VB_VarHelpID = -1
Dim mblnAddMode As Boolean
Dim itmx As ListItem

Private Sub cmdActivate_Click()
    
  On Error GoTo ResumeNext
    
    PlaySound App.Path & "\sounds\groupopen.wav"
    
  If txtText1.Text = "" Then 'if the first text field is empty then
    MsgBox "You must choose a record from the database before enabling the edit option.", , "Database Error"
  Exit Sub
  
  Else
  
    cmdEdit.Enabled = True
    cmdAdd.Enabled = False 'make sure the user cannot enter during editing mode
    
    txtText1.SetFocus 'set the focus to the first text field
    
  End If
    
ResumeNext:
    
End Sub

Private Sub cmdActivate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdActivate.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdActivate.ForeColor = &H0&  'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdAdd_Click()
    
    On Error GoTo ResumeNext
    
    PlaySound App.Path & "\sounds\groupopen.wav"
    
    cmdEdit.Enabled = False 'disable the user from editing any records until after there is a record entered in the database
    cmdSave.Enabled = True  'enable the record to be saved
        
        
  Call DisableNavigation 'call the diable button feature placed earlier
        
        
    mblnAddMode = True 'We are now in addmode.
        
  Call ClearControls 'call up to the clear controls sub we created earlier
        
    cmdDelete.Enabled = False 'do not allow the record to be deleted till the user finishes saving
        
        
    txtText1.Locked = False 'unlock the text boxes so info can be input
    txtText2.Locked = False
    txtText3.Locked = False
    txtText4.Locked = False
    txtText5.Locked = False
    txtText6.Locked = False
    txtText7.Locked = False
    txtText8.Locked = False
    txtText9.Locked = False
    txtText10.Locked = False
    txtText11(1).Locked = False
    txtText12(0).Locked = False
    txtText13.Locked = False
        
    txtText1.SetFocus
    ListView1.Refresh
    
ResumeNext:
        
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdAdd.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdAdd.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
    
End Sub

Private Sub cmdDefault_Click()
    
    PlaySound App.Path & "\sounds\groupopen.wav"
    
    fraFrame12.Visible = False
    fraFrame11.Visible = True
    
    lblLabel1(0).Caption = "Name:"
    lblLabel2.Caption = "Address:"
    lblLabel3.Caption = "E-Mail Address:"
    lblLabel4.Caption = "Phone Number:"
    lblLabel5.Caption = "Cell Number:"
    lblLabel6.Caption = "Beeper Number:"
    
    txtText1.Visible = True
    txtText2.Visible = True
    txtText3.Visible = True
    txtText4.Visible = True
    txtText5.Visible = True
    txtText6.Visible = True
    txtText8.Visible = False
    txtText9.Visible = False
    txtText10.Visible = False
    txtText11(1).Visible = False
    txtText12(0).Visible = False
    txtText13.Visible = False
  
End Sub

Private Sub cmdDefault_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdDefault.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdDefault.ForeColor = &H0& 'change the color of the text when mouse hovers over
      
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
End Sub

Private Sub cmdDelete_Click()

  On Error GoTo ResumeNext

    PlaySound App.Path & "\sounds\groupopen.wav"

  If txtText1.Text = "" Then 'If the first text box is empty then
    MsgBox "You must choose a record to delete.", , "Database Error"
  
  Exit Sub
  
  Else
  If rsInfo.EOF = False And _
    rsInfo.BOF = False Then
        'Check to see if there is data in the database
        'and make sure it is open.

  On Error Resume Next
        'If there is an error, ignore it.
        
    connConnection.begtrans
        'Deleting a record is important so the begtrans method
        'is used.  It makes sure all actions between begtrans
        'and committrans are done at the same time.
    rsInfo.Delete
        'Delete the record.
        
    connConnection.CommitTrans
        'The actions have been committed.
        
    rsInfo.MoveNext
  If rsInfo.EOF = True Then
    rsInfo.MoveLast
            'If the user deletes the record in the last position
            'go to the new record in the last position.
            
  If rsInfo.BOF = True Then
  Call ClearControls
                'If the last record is deleted, clear the text
                'boxes.
                
    MsgBox "There is no data in the recordset!", , "Database Error!"
                'Alert the user that there is no more data in
                'the database.
  End If
  End If
    ElseIf rsInfo.EOF = True And rsInfo.BOF = True Then
        'Warn the user that he or she is trying to delete data
        'from a database with no records.
        
    MsgBox "There is no data in the recordset!", , "Database Error!"
  End If
    
    ListView1.ListItems.Clear
  While Not rsInfo.EOF()
  Set itmx = ListView1.ListItems.Add

    itmx.Text = rsInfo("Name")
    itmx.SubItems(1) = rsInfo("Address")
    itmx.SubItems(2) = rsInfo("Email")
    itmx.SubItems(3) = rsInfo("Phone")
    itmx.SubItems(4) = rsInfo("Cell")
    itmx.SubItems(5) = rsInfo("Beeper")
    itmx.SubItems(6) = rsInfo("ICQ Number")
    itmx.SubItems(7) = rsInfo("AIM Handle")
    itmx.SubItems(8) = rsInfo("IRC Handle")
    itmx.SubItems(9) = rsInfo("Gaming Handle")
    itmx.SubItems(10) = rsInfo("Homepage URL")
    itmx.SubItems(11) = rsInfo("Sex")
    
    rsInfo.MoveNext
  Wend
     
     txtText1.SetFocus
     ListView1.Refresh
     
  End If
  
ResumeNext:
     
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdDelete.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdDelete.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdEdit_Click()

  On Error GoTo ResumeNext

    PlaySound App.Path & "\sounds\groupopen.wav"

  If txtText1.Text = "" Then 'if the first text box has no info then
    MsgBox "You did not choose a contact to edit.", , "Edit Record Error"
  
  Exit Sub
   
  Else
    cmdSave.Enabled = True
        'Since a new record is being added, the user
        'should be allowed to save the data.
        
  Call DisableNavigation
        'No moves during edit.
        
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
        'Coded later to add more stability.
        
    txtText1.Locked = False
    txtText2.Locked = False
    txtText3.Locked = False
    txtText4.Locked = False
    txtText5.Locked = False
    txtText6.Locked = False
    txtText7.Locked = False
    txtText8.Locked = False
    txtText9.Locked = False
    txtText10.Locked = False
    txtText11(1).Locked = False
    txtText12(0).Locked = False
    txtText13.Locked = False
        'Allow the user to enter items into the text boxes.
             
  Call LoadDataInControls
        
    txtText1.SetFocus 'set the focus to the first text field
        
  End If
        
ResumeNext:
        
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdEdit.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdEdit.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdExtra_Click()
    
    PlaySound App.Path & "\sounds\groupopen.wav"
    
    fraFrame12.Visible = True
    fraFrame11.Visible = False
    
    lblLabel1(0).Caption = "ICQ:"
    lblLabel2.Caption = "AIM:"
    lblLabel3.Caption = "IRC Handle:"
    lblLabel4.Caption = "Gaming Handle:"
    lblLabel5.Caption = "Homepage URL:"
    lblLabel6.Caption = "Sex:"
    
    txtText1.Visible = False
    txtText2.Visible = False
    txtText3.Visible = False
    txtText4.Visible = False
    txtText5.Visible = False
    txtText6.Visible = False
    txtText8.Visible = True
    txtText9.Visible = True
    txtText10.Visible = True
    txtText11(1).Visible = True
    txtText12(0).Visible = True
    txtText13.Visible = True

End Sub

Private Sub cmdExtra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdExtra.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdExtra.ForeColor = &H0& 'change the color of the text when mouse hovers over
      
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
    
End Sub

Private Sub cmdFirst_Click()
  
  On Error GoTo ResumeNext

    PlaySound App.Path & "\sounds\groupopen.wav"
    
  If rsInfo.BOF = False Then
        rsInfo.MoveFirst
        'Move to the first record in the record set.
    ElseIf rsInfo.BOF = True _
        And rsInfo.EOF = True Then
        
        MsgBox "There is no data in the record set!", , "Record Error"
    End If
    
    txtText1.SetFocus
    ListView1.Refresh
    
ResumeNext:
    
End Sub

Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    cmdFirst.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdFirst.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
    
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
    
End Sub

Private Sub cmdLast_Click()

  On Error GoTo ResumeNext
  
    PlaySound App.Path & "\sounds\groupopen.wav"
  
  If rsInfo.EOF = False Then
        rsInfo.MoveLast
        'Move to the last record in the record set.
    ElseIf rsInfo.BOF = True _
        And rsInfo.EOF = True Then
        
        MsgBox "There is no data in the record set!", , "Database Error"
    End If
    
    txtText1.SetFocus
    ListView1.Refresh
    
ResumeNext:
      
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdLast.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdLast.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdNext_Click()

  On Error GoTo ResumeNext

    PlaySound App.Path & "\sounds\groupopen.wav"
  
  If rsInfo.EOF = False Then
        rsInfo.MoveNext
        
        If rsInfo.EOF Then
            
            
            rsInfo.MoveLast
        End If
    Else
        If rsInfo.BOF Then
            'Check to see if there is any data in the recordset.
        
            MsgBox "There is no data in the record set!", , "Database Error"
        Else
            rsInfo.MoveLast
            
        End If
    End If
    
    txtText1.SetFocus
    ListView1.Refresh
    
ResumeNext:
    
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdNext.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdNext.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdPrevious_Click()
  
  On Error GoTo ResumeNext
  
    PlaySound App.Path & "\sounds\groupopen.wav"
    
  If rsInfo.BOF = False Then
        rsInfo.MovePrevious
        'Check to see if you are at the front of the record set.
        'If you are not, then you can move forward.
        
        If rsInfo.BOF = True Then
            'This will prevent the user from moving to
            'the BOF marker if he or she is on the first record.
            rsInfo.MoveFirst
        End If
    Else
        If rsInfo.EOF Then
            'Check to see if there is any data in the record set.
            MsgBox "There is no data in the record set!", , "Database Error"
        Else
            rsInfo.MoveFirst
            
            
        End If
    End If
    
    txtText1.SetFocus
    ListView1.Refresh
    
ResumeNext:
  
End Sub


Private Sub DisableNavigation()
    
    cmdFirst.Enabled = False 'this will be called upon later
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    
    
End Sub

Private Sub EnableNavigation()
    
    cmdFirst.Enabled = True 'this will be called upon later
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    
    
End Sub

Private Sub WriteDataFromControls()
    
    
    'Again, there are several ways to manipulate field
    'values.  Some are shown.
    
    rsInfo("Name").Value = txtText1.Text
    
    rsInfo.Fields("Address").Value = txtText2.Text
    
    rsInfo.Fields("email").Value = txtText3.Text
    
    rsInfo.Fields("phone").Value = txtText4.Text
    
    rsInfo.Fields("cell").Value = txtText5.Text
    
    rsInfo.Fields("Schedule").Value = txtText7.Text
    
    rsInfo.Fields("ICQ Number").Value = txtText8.Text
    
    rsInfo.Fields("AIM Handle").Value = txtText9.Text
    
    rsInfo.Fields("IRC Handle").Value = txtText10.Text
    
    rsInfo.Fields("Gaming Handle").Value = txtText11(1).Text
    
    rsInfo.Fields("Homepage URL").Value = txtText12(0).Text
    
    rsInfo.Fields("Sex").Value = txtText13.Text
    
    rsInfo!Beeper = txtText6.Text
    
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdPrevious.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdPrevious.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub cmdSave_Click()
    
  On Error GoTo ResumeNext

    PlaySound App.Path & "\sounds\groupopen.wav"

  If txtText1.Text = "" Then 'If the first text area has no info then
    MsgBox "You must enter a name.", , "Database Error"
  
  Exit Sub
  
  Else
    cmdActivate.Enabled = True
    
  If cmdEdit.Enabled = False Then
    rsInfo.AddNew
  End If
    'This calls the Recordset MoveComplete event.
    'Before cmdEdit was added, there was no if-then statement.
    
  Call WriteDataFromControls
    
    rsInfo.Update
    mblnAddMode = False
    'Saving closes add mode.
    
    cmdSave.Enabled = False
    
    
    'These three steps prevent the user from adding
    'blank data or accidently changing data.
    txtText1.Locked = True
    txtText2.Locked = True
    txtText3.Locked = True
    txtText4.Locked = True
    txtText5.Locked = True
    txtText6.Locked = True
    txtText7.Locked = True
    txtText8.Locked = True
    txtText9.Locked = True
    txtText10.Locked = True
    txtText11(1).Locked = True
    txtText12(0).Locked = True
    txtText13.Locked = True
    
  Call EnableNavigation
    
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    'Add later to provide stability and consistency.
    
    rsInfo.Close
    rsInfo.Open
      
      ListView1.ListItems.Clear
           While Not rsInfo.EOF()
  Set itmx = ListView1.ListItems.Add

    itmx.Text = rsInfo("Name")
    itmx.SubItems(1) = rsInfo("Address")
    itmx.SubItems(2) = rsInfo("Email")
    itmx.SubItems(3) = rsInfo("Phone")
    itmx.SubItems(4) = rsInfo("Cell")
    itmx.SubItems(5) = rsInfo("Beeper")
    itmx.SubItems(6) = rsInfo("ICQ Number")
    itmx.SubItems(7) = rsInfo("AIM Handle")
    itmx.SubItems(8) = rsInfo("IRC Handle")
    itmx.SubItems(9) = rsInfo("Gaming Handle")
    itmx.SubItems(10) = rsInfo("Homepage URL")
    itmx.SubItems(11) = rsInfo("Sex")

    rsInfo.MoveNext
  Wend
  
  Call ClearControls
  
    txtText1.SetFocus
    ListView1.Refresh
  
  End If

ResumeNext:
    
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdSave.BackColor = &HFFFF00 'change the color of the background when mouse hovers over
    cmdSave.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    cmdFirst.BackColor = &H0&
    cmdFirst.ForeColor = &HFFFF00
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
        
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
        
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00

End Sub

Private Sub Form_Load()

Dim strConnect As String
    Dim strProvider As String 'name the database provider
    Dim strDataSource As String 'let the system know what data source to search in for the database
    Dim strDataBaseName As String 'name of the database
    strProvider = "Provider= Microsoft.Jet.OLEDB.4.0;"
    strDataSource = App.Path
    strDataBaseName = "\Grinch.mdb;"
    strDataSource = "Data Source=" & strDataSource & _
    strDataBaseName
    strConnect = strProvider & strDataSource
       
    Set connConnection = New ADODB.Connection
    connConnection.CursorLocation = adUseClient
    connConnection.Open strConnect
    
    Set rsInfo = New ADODB.Recordset
    rsInfo.CursorType = adOpenStatic
    rsInfo.CursorLocation = adUseClient
    rsInfo.LockType = adLockPessimistic
    rsInfo.Source = "Select * From [Info]"
    rsInfo.ActiveConnection = connConnection
    rsInfo.Open
    
     ListView1.ListItems.Clear
    While Not rsInfo.EOF()
  Set itmx = ListView1.ListItems.Add

    itmx.Text = rsInfo("Name")
    itmx.SubItems(1) = rsInfo("Address")
    itmx.SubItems(2) = rsInfo("Email")
    itmx.SubItems(3) = rsInfo("Phone")
    itmx.SubItems(4) = rsInfo("Cell")
    itmx.SubItems(5) = rsInfo("Beeper")
    itmx.SubItems(6) = rsInfo("ICQ Number")
    itmx.SubItems(7) = rsInfo("AIM Handle")
    itmx.SubItems(8) = rsInfo("IRC Handle")
    itmx.SubItems(9) = rsInfo("Gaming Handle")
    itmx.SubItems(10) = rsInfo("Homepage URL")
    itmx.SubItems(11) = rsInfo("Sex")

    rsInfo.MoveNext
  Wend
  
  cmdActivate.Enabled = True
  Call ClearControls
   
                                            
End Sub

Private Sub ClearControls()

    txtText1.Text = "" 'clear all text fields
    txtText2.Text = ""
    txtText3.Text = ""
    txtText4.Text = ""
    txtText5.Text = ""
    txtText6.Text = ""
    txtText7.Text = ""
    txtText8.Text = ""
    txtText9.Text = ""
    txtText10.Text = ""
    txtText11(1).Text = ""
    txtText12(0).Text = ""
    txtText13.Text = ""
    
End Sub

Private Sub LoadDataInControls()

    If rsInfo.BOF = True Or rsInfo.EOF = True Then
        Exit Sub
        
    End If
       
    txtText1.Text = rsInfo.Fields("Name").Value & " "
    txtText2.Text = rsInfo("Address").Value & " "
    txtText3.Text = rsInfo("Email").Value & " "
    txtText4.Text = rsInfo("Phone").Value & " "
    txtText5.Text = rsInfo("Cell").Value & " "
    txtText7.Text = rsInfo("Schedule").Value & " "
    txtText8.Text = rsInfo("ICQ Number").Value & " "
    txtText9.Text = rsInfo("AIM Handle").Value & " "
    txtText10.Text = rsInfo("IRC Handle").Value & " "
    txtText11(1).Text = rsInfo("Gaming Handle").Value & " "
    txtText12(0).Text = rsInfo("Homepage URL").Value & " "
    txtText13.Text = rsInfo("Sex").Value & " "
    txtText6.Text = rsInfo!Beeper & " "
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdFirst.BackColor = &H0& 'change the color of the background when mouse hovers over
    cmdFirst.ForeColor = &HFFFF00 'change the color of the text when mouse hovers over
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
    
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
        
End Sub

Private Sub fraScheduler_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdFirst.BackColor = &H0& 'change the color of the background when mouse hovers over
    cmdFirst.ForeColor = &HFFFF00 'change the color of the text when mouse hovers over
    
    cmdPrevious.BackColor = &H0&
    cmdPrevious.ForeColor = &HFFFF00
    
    cmdNext.BackColor = &H0&
    cmdNext.ForeColor = &HFFFF00
    
    cmdLast.BackColor = &H0&
    cmdLast.ForeColor = &HFFFF00
    
    cmdAdd.BackColor = &H0&
    cmdAdd.ForeColor = &HFFFF00
    
    cmdSave.BackColor = &H0&
    cmdSave.ForeColor = &HFFFF00
    
    cmdDelete.BackColor = &H0&
    cmdDelete.ForeColor = &HFFFF00
    
    cmdEdit.BackColor = &H0&
    cmdEdit.ForeColor = &HFFFF00
    
    cmdActivate.BackColor = &H0&
    cmdActivate.ForeColor = &HFFFF00
    
    cmdExtra.BackColor = &H0&
    cmdExtra.ForeColor = &HFFFF00
    
    cmdDefault.BackColor = &H0&
    cmdDefault.ForeColor = &HFFFF00
    
End Sub

Private Sub rsinfo_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If mblnAddMode = False Then 'if not in add mode
        
        Call LoadDataInControls 'call up to loaddataincontrols sub
        
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Unload Me 'remove from memory
    connConnection.Close
    Set connConnection = Nothing 'close the database connection
    
    Call ClearControls
    
  End
    
End Sub

Private Sub Form_Terminate()
  
  Unload Me 'remove from memory
    connConnection.Close
    Set connConnection = Nothing 'close the database connection
    

    Call ClearControls
    
  End
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
    Unload Me
    connConnection.Close
    Set connConnection = Nothing
    

    Call ClearControls
    
  End
  
End Sub

Private Sub SSListBar1_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)

    SSListBar1.PlaySoundFile App.Path & "\sounds\groupopen.wav" 'play the wav when a button is clicked on the toolbar"
    
    Select Case ItemClicked.Key
    Case "change"
        FrmChangePassword.Show 'show the change password form
    
    Case "show"
       
  If fraFrame1.Visible = True Then 'if the first frame is visible which means everything else is still visible also Then
     MsgBox "Controls are already showing, hide the controls before pressing this." 'Display the message box to the user
  
  Exit Sub 'cancel the users actions if the controls are visible
  
  End If
  
  If lblLabel1(0).Caption = "ICQ:" Then 'if the user has pressed the extra command then
    MsgBox "You cannot show the controls when showing the extra database entry fields." 'display the message box and let the user know he cannot show controls
  
  Exit Sub 'cancel the users choice if the show controls button was pressed after the extra command was pressed
  
  Else 'if extra was not pressed then the user can show the controls again
     fraScheduler.Visible = True 'make the frames visible again
     fraFrame1.Visible = True
     fraFrame2.Visible = True
     fraFrame3.Visible = True
     fraFrame4.Visible = True
     fraFrame5.Visible = True
     fraFrame6.Visible = True
     fraFrame7.Visible = True
     fraFrame8.Visible = True
     fraFrame9.Visible = True
     fraFrame10.Visible = True
     fraFrame11.Visible = True
          
     txtText7.Visible = True 'show the memo field
     txtText8.Visible = False 'to be sure the text boxes behind the extra command doesn't show
     txtText9.Visible = False
     txtText10.Visible = False
     txtText11(1).Visible = False
     txtText12(0).Visible = False
     txtText13.Visible = False
  End If
     
    Case "hide"
     
  If fraFrame1.Visible = False Then 'if the first frame is not visible which means everything else isn't Then
     MsgBox "Controls are already hidden, show the controls before pressing this." 'Display the message box to the user
  
  Exit Sub 'cancel the users actions if the controls aren't visible
  
  End If
  
  If lblLabel1(0).Caption = "ICQ:" Then 'if the user has pressed the extra command then
    MsgBox "You cannot hide the controls when showing the extra database entry fields." 'display the message box and let the user know he cannot hide the controls
  
  Exit Sub 'cancel the users choice if the hide controls button was pressed after the extra command was pressed
  
  Else 'if extra was not pressed then the user can hide the controls
     fraScheduler.Visible = False 'hide all frames
     fraFrame1.Visible = False
     fraFrame2.Visible = False
     fraFrame3.Visible = False
     fraFrame4.Visible = False
     fraFrame5.Visible = False
     fraFrame6.Visible = False
     fraFrame7.Visible = False
     fraFrame8.Visible = False
     fraFrame9.Visible = False
     fraFrame10.Visible = False
     fraFrame11.Visible = False
     
     txtText7.Visible = False 'make sure the text boxes are hidden as well
     txtText8.Visible = False
     txtText9.Visible = False
     txtText10.Visible = False
     txtText11(1).Visible = False
     txtText12(0).Visible = False
     txtText13.Visible = False
  End If
  
    Case "email"
    Shell ("Start mailto:" & txtText3.Text), vbHide 'opens your mail program and places the text in the To: field
        
    Case "help"
        frmHelp.Show
        
    Case "exit"
        Unload Me
        
  End Select
  
End Sub




Private Sub txtText1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtText10_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If


End Sub

Private Sub txtText11_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If


End Sub

Private Sub txtText12_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If


End Sub

Private Sub txtText13_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtText2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub
Private Sub txtText3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtText4_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtText5_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtText6_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtText7_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtText8_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtText9_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub
