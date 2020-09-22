VERSION 5.00
Begin VB.UserControl CoolPgBar 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   PropertyPages   =   "CoolPgBar.ctx":0000
   ScaleHeight     =   630
   ScaleWidth      =   3090
   ToolboxBitmap   =   "CoolPgBar.ctx":0016
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox picCover 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   270
         ScaleHeight     =   300
         ScaleWidth      =   630
         TabIndex        =   1
         Top             =   45
         Width           =   630
      End
   End
End
Attribute VB_Name = "CoolPgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*******************************************************************
'This ActiveX Control was created by MauTheMan                     *
'You may use it in your projects and distribute it freely          *
'I would appreciate an E-mail just to let me know about how many   *
'people are using my OCX (statistics purpose only)                 *
'This is the 1st version of this OCX, Report bugs please.          *
'                    mautheman@yahoo.com                           *
'******************************************************************
'How to use the CoolProgressBar Control                            *
'The same way you would use a regular ProgressBar :                *
'Especify the Max ( the number that will be considered 100%)       *
'do that "CoolPgBar1.value = CoolPgBar1.value + 1" to make it work *
'*******************************************************************

Public Enum BorderStyles
    None = 0
    [Fixed Single] = 1
End Enum

Public Enum ProgressStyle
    Normal = 0
    Graphic = 1
End Enum

Public Enum BarStyle
    Horizontal = 0
    Vertical = 1
End Enum

Public Enum CaptionStyles
    Default = 0
    Percentage = 1
    Custom = 2
End Enum

Const mcValue = 1
Const mcMax = 100
Private mValue As Long
Private mMax As Long
Private m_ProgressStyle As ProgressStyle
Private m_Orientation As BarStyle
Private m_Caption As String
Private m_CaptionStyle As CaptionStyles

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Value = mValue
    End Property

Public Property Let Value(ByVal vValue As Long)
    mValue = vValue
    PropertyChanged "Value"
    
    If Not mValue > mMax Then
        Progress Picture1, mMax, mValue
    Else
        Picture1.Cls
    End If
    
    End Property

Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Max = mMax
End Property

Public Property Let Max(ByVal vMax As Long)
    mMax = vMax
    PropertyChanged "Max"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Picture1.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    Picture1.ForeColor = vNewValue
    UserControl.PropertyChanged "ForeColor"
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    Picture1.BackColor = vNewValue
    UserControl.PropertyChanged "BackColor"
End Property

Public Property Get DrawMode() As DrawModeConstants
    DrawMode = Picture1.DrawMode
End Property

Public Property Let DrawMode(ByVal vNewValue As DrawModeConstants)
    Picture1.DrawMode = vNewValue
    UserControl.PropertyChanged "DrawMode"
End Property

Public Property Get Font() As StdFont
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set Picture1.Font = vNewValue
    UserControl.PropertyChanged "Font"
End Property

Public Property Get BorderStyle() As BorderStyles
    BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As BorderStyles)
    Picture1.BorderStyle = vNewValue
    UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    m_Caption = vNewValue
    UserControl.PropertyChanged "Caption"
End Property

Public Property Get Orientation() As BarStyle
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal vNewValue As BarStyle)
    If m_Orientation = Horizontal Or m_Orientation = Vertical Then
        m_Orientation = vNewValue
        UserControl.PropertyChanged "Orientation"
    Else
        Err.Raise Number:=vbObjectError + 32114, Description:="Invalid Style value (0 or 1 only)"
    End If
End Property

Public Property Get CaptionStyle() As CaptionStyles
    CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal vNewValue As CaptionStyles)
    If m_CaptionStyle = Custom Or m_CaptionStyle = Default Or m_CaptionStyle = Percentage Then
        m_CaptionStyle = vNewValue
        UserControl.PropertyChanged "CaptionStyle"
    Else
        Err.Raise Number:=vbObjectError + 32115, Description:="Invalid Style value (0, 1, or 2 only)"
    End If
    
End Property
Public Property Get Style() As ProgressStyle
    Style = m_ProgressStyle
End Property

Public Property Let Style(ByVal vNewValue As ProgressStyle)
    If vNewValue = Graphic Or vNewValue = Normal Then
        m_ProgressStyle = vNewValue
        
        If m_ProgressStyle = Normal Then
            picCover.Visible = False
            Set Picture1.Picture = Nothing
            Set picCover.Picture = Nothing
        End If
        
        If m_ProgressStyle = Graphic Then
            Caption = ""
            CaptionStyle = Default
            With picCover
                .Visible = True
                .Left = 0
                .Top = 0
                .Height = Picture1.Height
            End With
            
        End If
        UserControl.PropertyChanged "Style"
    Else
        Err.Raise Number:=vbObjectError + 32113, Description:="Invalid Style value (0 or 1 only)"
        
    End If
    
    
End Property
Public Property Get BackPicture() As StdPicture
Attribute BackPicture.VB_ProcData.VB_Invoke_Property = "StandardPicture;Appearance"
    Set BackPicture = Picture1.Picture
End Property

Public Property Set BackPicture(ByVal vNewValue As StdPicture)
    Set Picture1.Picture = vNewValue
    If vNewValue Is Nothing Then
        Style = Normal
    Else
        Style = Graphic
        Picture1.AutoSize = True
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
    UserControl.PropertyChanged "Picture"
    
End Property

Public Property Get CoverPicture() As StdPicture
Attribute CoverPicture.VB_ProcData.VB_Invoke_Property = "StandardPicture;Appearance"
    Set CoverPicture = picCover.Picture
End Property

Public Property Set CoverPicture(ByVal vNewValue As StdPicture)
    Set picCover.Picture = vNewValue
    If vNewValue Is Nothing Then
        Style = Normal
    Else
        Style = Graphic
    End If
    UserControl.PropertyChanged "Picture"
    
End Property

Private Sub UserControl_InitProperties()
    mValue = mcValue
    mMax = mcMax
    DrawMode = vbNotXorPen '10
    BackColor = vbButtonFace
    ForeColor = RGB(50, 50, 150)
    Style = m_ProgressStyle
    CaptionStyle = Default
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mValue = PropBag.ReadProperty("Value", "1")
    mMax = PropBag.ReadProperty("Max", "100")
    Caption = PropBag.ReadProperty("Caption", "")
    Picture1.ForeColor = PropBag.ReadProperty("ForeColor", RGB(50, 50, 150))
    Picture1.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    Picture1.DrawMode = PropBag.ReadProperty("DrawMode", 10)
    Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture1.Picture = PropBag.ReadProperty("BackPicture", Nothing)
    Set picCover.Picture = PropBag.ReadProperty("CoverPicture", Nothing)
    Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Style = PropBag.ReadProperty("Style", Normal)
    Orientation = PropBag.ReadProperty("Orientation", Horizontal)
    CaptionStyle = PropBag.ReadProperty("CaptionStyle", Default)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", mValue, "1")
    Call PropBag.WriteProperty("Max", mMax, "100")
    Call PropBag.WriteProperty("ForeColor", ForeColor, RGB(50, 50, 150))
    Call PropBag.WriteProperty("BackColor", BackColor, vbButtonFace)
    Call PropBag.WriteProperty("DrawMode", DrawMode, 10)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 1)
    Call PropBag.WriteProperty("Style", Style, Normal)
    Call PropBag.WriteProperty("BackPicture", BackPicture, Nothing)
    Call PropBag.WriteProperty("CoverPicture", CoverPicture, Nothing)
    Call PropBag.WriteProperty("Orientation", Orientation, Horizontal)
    Call PropBag.WriteProperty("Caption", Caption, "")
    Call PropBag.WriteProperty("CaptionStyle", CaptionStyle, Default)
    
End Sub

Private Sub UserControl_Resize()
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    If Style = Graphic Then
        picCover.Height = Picture1.Height
        picCover.Width = Picture1.Width
    End If
    
End Sub

Private Sub Progress(PictureProgress As PictureBox, MaxLength As Long, Value As Long)
    Dim i As Long
    i = (100 * Value) / MaxLength
    
    PictureProgress.ForeColor = ForeColor
    PictureProgress.BackColor = BackColor
    FillIt PictureProgress, i
    
End Sub
Private Sub FillIt(PicBox As Control, percent As Long)
    Dim Msg As String
        
    If Not PicBox.AutoRedraw Then
        PicBox.AutoRedraw = -1
    End If
    
    If Style = Normal Then
            PicBox.Cls
            PicBox.DrawMode = DrawMode
            PicBox.ScaleWidth = 100
            
            If CaptionStyle = Default Then Caption = ""
            If CaptionStyle = Percentage Then Caption = Format$(percent, "###") + "%"
            If CaptionStyle = Custom Then Caption = Caption
            PicBox.CurrentX = (PicBox.ScaleWidth - PicBox.TextWidth(Caption)) / 2
            PicBox.CurrentY = (PicBox.ScaleHeight - PicBox.TextHeight(Caption)) / 2
            PicBox.Print Caption
        
        If Orientation = Horizontal Then
            PicBox.ScaleWidth = 100
            PicBox.Line (0, 0)-(percent, PicBox.ScaleHeight), , BF
        Else
            PicBox.ScaleHeight = Max - Min
            PicBox.Line (0, PicBox.ScaleHeight)-(PicBox.ScaleWidth, PicBox.ScaleHeight - percent), , BF
        End If
        
        PicBox.Refresh
    End If
    
    If Style = Graphic Then
        PicBox.Cls
        If Orientation = Horizontal Then
            PicBox.ScaleWidth = 100
            picCover.Width = PicBox.Width
            picCover.Height = PicBox.Height
            picCover.Move picCover.Left + percent
        Else
            PicBox.ScaleHeight = Max - Min
            picCover.Move picCover.Left, picCover.Top - percent
        End If
        
    End If
    
End Sub




