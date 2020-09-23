VERSION 5.00
Begin VB.UserControl LayerItem 
   BackColor       =   &H80000005&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VBLayers.FlatPicture picView 
      Height          =   255
      Left            =   465
      TabIndex        =   8
      Top             =   105
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      AutoRedraw      =   -1  'True
   End
   Begin VBLayers.FlatPicture picEdit 
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   105
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      AutoRedraw      =   -1  'True
   End
   Begin VB.PictureBox picBltThumb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3810
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTempView 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2310
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTempEdit 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2025
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   945
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   390
      ScaleHeight     =   375
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   45
      Width           =   15
   End
   Begin VB.CommandButton cmdDummy 
      Enabled         =   0   'False
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   810
   End
   Begin VB.Line linBottomIndicator 
      BorderColor     =   &H8000000D&
      Visible         =   0   'False
      X1              =   312
      X2              =   265
      Y1              =   31
      Y2              =   31
   End
   Begin VB.Line linTopIndicator 
      BorderColor     =   &H8000000D&
      Visible         =   0   'False
      X1              =   314
      X2              =   267
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   225
      Left            =   1365
      TabIndex        =   3
      Top             =   150
      Width           =   3300
   End
   Begin VB.Line linTopSeperator 
      Visible         =   0   'False
      X1              =   76
      X2              =   305
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape shpHighlight 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   2115
      Top             =   75
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Shape shpBorder 
      Height          =   270
      Left            =   930
      Top             =   105
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line linBottomSeperator 
      Visible         =   0   'False
      X1              =   74
      X2              =   303
      Y1              =   31
      Y2              =   31
   End
End
Attribute VB_Name = "LayerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************************'
'*'
'*' Module    : LayerItem
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.20.2004
'*'
'*' Depends   :
'*'
'*' Purpose   :
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' API Declarations - gdi32.dll
'*'
'**********************************************************************************************************************'
Private Declare Function StretchBlt Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, _
        ByVal dwRop As Long) As Long

'**********************************************************************************************************************'
'*'
'*' API Declarations - msimg32.dll
'*'
'**********************************************************************************************************************'
Private Declare Function TransparentBlt Lib "msimg32.dll" ( _
        ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, _
        ByVal crTransparent As Long) As Boolean
 
'**********************************************************************************************************************'
'*'
'*' Private Member Declarations
'*'
'**********************************************************************************************************************'
Private m_bolEditable                   As Boolean
Private m_bolSelected                   As Boolean
Private m_bolShowBottomSeperator        As Boolean
Private m_bolShowTopSeperator           As Boolean
Private m_bolUseThumbnail               As Boolean
Private m_bolViewable                   As Boolean
Private m_lngThumbnailWidth             As Long
Private m_lngThumbnailHeight            As Long
Private m_spcThumbnail                  As StdPicture
Private m_strKey                        As String
Private m_strTag                        As String
Private m_strInternalIdentifier         As String

'**********************************************************************************************************************'
'*'
'*' Private Event Declarations
'*'
'**********************************************************************************************************************'
Event Click()
Event DblClick()
Event InitProperties()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ReadProperties(PropBag As PropertyBag)
Event Resize()
Event Show()
Event SetEditable(Value As Boolean)
Event SetVisible(Value As Boolean)
Event SelectionChange(Value As Boolean)
Event WriteProperties(PropBag As PropertyBag)

'**********************************************************************************************************************'
'*'
'*' Property  : BottomIndicator
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for showing the BottomIndicator on Control
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : BottomIndicator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get BottomIndicator() As Boolean
    BottomIndicator = linBottomIndicator.Visible
End Property
Public Property Let BottomIndicator(Value As Boolean)
    If Not linBottomIndicator.Visible = Value Then
        linBottomIndicator.Visible = Value
    End If
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Caption
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for setting the caption of a layer.
'*'
'*' Input     : Value (String)
'*'
'*' Output    : Caption (String)
'*'
'**********************************************************************************************************************'
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : InternalIdentifier
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for a unique InternalIdentifier within the Control (preferably GUID)
'*'
'*' Input     : Value (String)
'*'
'*' Output    : InternalIdentifier (String)
'*'
'**********************************************************************************************************************'
Public Property Get InternalIdentifier() As String
    InternalIdentifier = m_strInternalIdentifier
End Property
Public Property Let InternalIdentifier(Value As String)
    m_strInternalIdentifier = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : HasDC
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property Relay.
'*'
'*' Input     : None.
'*'
'*' Output    : HasDC (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : hDC
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property Relay.
'*'
'*' Input     : None.
'*'
'*' Output    : hDC (Long)
'*'
'**********************************************************************************************************************'
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : hWnd
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property Relay.
'*'
'*' Input     : None.
'*'
'*' Output    : hWnd (Long)
'*'
'**********************************************************************************************************************'
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Key
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for a Key identifier for the Control
'*'
'*' Input     : Value (String)
'*'
'*' Output    : Key (String)
'*'
'**********************************************************************************************************************'
Public Property Get Key() As String
    Key = m_strKey
End Property
Public Property Let Key(Value As String)
    m_strKey = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : LayerEditable
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for a flag to toggle editablity of the layer.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : LayerEditable (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Let LayerEditable(Value As Boolean)
    m_bolEditable = Value
    RaiseEvent SetEditable(Value)
    RedrawControl
End Property
Public Property Get LayerEditable() As Boolean
    LayerEditable = m_bolEditable
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : LayerViewable
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for a flag to toggle visibility of the layer.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : LayerViewable (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Let LayerViewable(Value As Boolean)
    m_bolViewable = Value
    RaiseEvent SetVisible(Value)
    RedrawControl
End Property
Public Property Get LayerViewable() As Boolean
    LayerViewable = m_bolViewable
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Picture
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property to set a piicture for the thumbnail
'*'
'*' Input     : Value (Picture)
'*'
'*' Output    : Picture (Picture)
'*'
'**********************************************************************************************************************'
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_spcThumbnail
End Property
Public Property Set Picture(ByVal Value As Picture)
    Set m_spcThumbnail = Value
    PropertyChanged "Picture"
    RedrawControl
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : ThumbnailHeight
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property to set the height of the thumbnail.
'*'
'*' Input     : Value (Long)
'*'
'*' Output    : ThumbnailHeight (Long)
'*'
'**********************************************************************************************************************'
Public Property Get ThumbnailHeight() As Long
    ThumbnailHeight = m_lngThumbnailHeight
End Property
Public Property Let ThumbnailHeight(Value As Long)
    m_lngThumbnailHeight = Value
    PropertyChanged "ThumbnailHeight"
    UserControl_Resize
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : ThumbnailWidth
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property to set the width of the thumbnail.
'*'
'*' Input     : Value (Long)
'*'
'*' Output    : ThumbnailWidth (Long)
'*'
'**********************************************************************************************************************'
Public Property Get ThumbnailWidth() As Long
    ThumbnailWidth = m_lngThumbnailWidth
End Property
Public Property Let ThumbnailWidth(Value As Long)
    m_lngThumbnailWidth = Value
    PropertyChanged "ThumbnailWidth"
    UserControl_Resize
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : UseThumbnail
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property to enable/disable use of a thumbnail.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : UseThumbnail (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get UseThumbnail() As Boolean
    UseThumbnail = m_bolUseThumbnail
End Property
Public Property Let UseThumbnail(Value As Boolean)
    m_bolUseThumbnail = Value
    PropertyChanged "UseThumbnail"
    UserControl_Resize
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Tag
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property to store a freeform tag on the individual layer.
'*'
'*' Input     : Value (String)
'*'
'*' Output    : Tag (String)
'*'
'**********************************************************************************************************************'
Public Property Get Tag() As String
    Tag = m_strTag
End Property
Public Property Let Tag(Value As String)
    m_strTag = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : TopIndicator
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for showing the TopIndicator on Control
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : TopIndicator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get TopIndicator() As Boolean
    TopIndicator = linTopIndicator.Visible
End Property
Public Property Let TopIndicator(Value As Boolean)
    If Not linTopIndicator.Visible = Value Then
        linTopIndicator.Visible = Value
    End If
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Selected
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for indicating if an individual layer is selected.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : Selected (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Let Selected(Value As Boolean)
    
    '*' Set the local member.
    '*'
    m_bolSelected = Value
    
    '*' Determine how to modify the display.
    '*'
    If m_bolSelected Then
    
        '*' Visible highlight, highlighted text, visible border.
        '*'
        shpHighlight.Visible = True
        lblCaption.ForeColor = vbHighlightText
        shpBorder.Visible = True
    Else
    
        '*' Hidden hightlight, standard text, hidden border.
        '*'
        shpHighlight.Visible = False
        lblCaption.ForeColor = vbWindowText
        shpBorder.Visible = False
    End If
    
    '*' Make sure that the controls are positioned and refreshed.  This is done since the rapid changes that can
    '*' occur in selection can cause events to be dropped in the calling parent.
    '*'
    UserControl_Resize
    UserControl.Refresh
    
End Property
Public Property Get Selected() As Boolean
    Selected = m_bolSelected
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : ShowBottomSeperator
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for indicating if a bottom seperator is rendered.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : ShowBottomSeperator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Let ShowBottomSeperator(Value As Boolean)
    m_bolShowBottomSeperator = Value
    linBottomSeperator.Visible = Value
End Property
Public Property Get ShowBottomSeperator() As Boolean
    ShowBottomSeperator = m_bolShowBottomSeperator
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : ShowTopSeperator
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Property for indicating if a top seperator is rendered.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : ShowTopSeperator (Booelan)
'*'
'**********************************************************************************************************************'
Public Property Let ShowTopSeperator(Value As Boolean)
    m_bolShowTopSeperator = Value
    linTopSeperator.Visible = Value
End Property
Public Property Get ShowTopSeperator() As Boolean
    ShowTopSeperator = m_bolShowTopSeperator
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : Refresh
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Tie in the refresh of all graphic intensive objects on the layer.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub Refresh()
    
    '*' Refresh all of the children controls that might run into redraw issues.
    '*'
    UserControl.Refresh
    lblCaption.Refresh
    cmdDummy.Refresh
    picEdit.Refresh
    picView.Refresh
    shpBorder.Refresh
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : RedrawControl
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Draw the graphical elements of the control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub RedrawControl()

'*' Fail through on errors.  This code is all drawing code.
'*'
On Error Resume Next

Dim lngPictureHeight                    As Long
Dim lngPicutreWidth                     As Long
Dim lngHorzOffset                       As Long
Dim lngVertOffset                       As Long

    '*' Blt the editible image if it belongs on the control.  Clear it, otherwise.
    '*'
    If m_bolEditable Then
        TransparentBlt picEdit.hDC, 2, 2, 14, 14, picTempEdit.hDC, 1, 1, 14, 14, RGB(192, 192, 192)
        picEdit.Refresh
    Else
        picEdit.Cls
    End If
    
    '*' Blt the viewable image if it belongs on the control.  Clear it, otherwise.
    '*'
    If m_bolViewable Then
        TransparentBlt picView.hDC, 2, 2, 14, 14, picTempView.hDC, 1, 1, 14, 14, RGB(192, 192, 192)
        picView.Refresh
    Else
        picView.Cls
    End If
    
    '*' Make sure that there is a thumbnail before trying to blt it.
    '*'
    If Not (m_spcThumbnail Is Nothing) Then
                
        '*' Set the buffer to hold the picture.
        '*'
        Set picBltThumb.Picture = m_spcThumbnail
        
        '*' Determine how the picture is oriented.
        '*'
        If picBltThumb.Height > picBltThumb.Width Then
        
            '*' Use the height to orient the scaling.
            '*'
            lngPictureHeight = picThumb.Height
            lngPicutreWidth = picThumb.Height * (picBltThumb.Width / picBltThumb.Width)
        
            '*' Determine to offset for centering vertically.
            '*'
            lngVertOffset = ((picThumb.Width - lngPicutreWidth) / 2) - 1
            
        ElseIf picBltThumb.Height < picBltThumb.Width Then
       
            '*' Use the width to orient the scaling.
            '*'
            lngPicutreWidth = picThumb.Height
            lngPictureHeight = (picBltThumb.Height / picBltThumb.Width) * lngPicutreWidth
        
            '*' Determine the offset for centering horizontally.
            '*'
            lngHorzOffset = ((picThumb.Height - lngPictureHeight) / 2) - 1
            
        Else
        
            '*' Treat them as equals.  No scaling.
            '*'
            lngPictureHeight = picThumb.Width
            lngPicutreWidth = picThumb.Width
            
        End If
       
        '*' Clear the current thumbnail.
        '*'
        picThumb.Cls
              
        '*' Scale and blt.
        '*'
        StretchBlt picThumb.hDC, lngVertOffset, lngHorzOffset, lngPicutreWidth, lngPictureHeight, _
                   picBltThumb.hDC, 0, 0, picBltThumb.Width, picBltThumb.Height, vbSrcCopy
  
        '*' Refresh the target.
        '*'
        picThumb.Refresh
        
    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdDummy_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdDummy_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdDummy_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdDummy_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdDummy_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdDummy_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblCaption_Click
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for Click()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblCaption_Click()
    RaiseEvent Click
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblCaption_DblClick
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for DblClick()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblCaption_DblClick()
    RaiseEvent DblClick
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblCaption_MouseDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseDown()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, lblCaption.Left + X - cmdDummy.Width, lblCaption.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblCaption_MouseMove
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseMove()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, lblCaption.Left + X - cmdDummy.Width, lblCaption.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : lblCaption_MouseUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseUp()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, lblCaption.Left + X - cmdDummy.Width, lblCaption.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picEdit_Click
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Toggle for layer editability.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picEdit_Click()
    Me.LayerEditable = Not (Me.LayerEditable)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picEdit_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picEdit_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress (KeyAscii)
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picEdit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picEdit_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picLine_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picLine_KeyDown(KeyCode As Integer, Shift As Integer)
'  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picLine_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picLine_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picLine_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picLine_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempEdit_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempEdit_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempEdit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempEdit_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempView_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempView_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempView_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempView_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picTempView_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picTempView_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_Click
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for Click()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_Click()
    RaiseEvent Click
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_DblClick
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for DblClick()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_DblClick()
    RaiseEvent DblClick
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_MouseDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseDown()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, picThumb.Left + X - cmdDummy.Width, picThumb.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_MouseMove
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseMove()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, picThumb.Left + X - cmdDummy.Width, picThumb.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picThumb_MouseUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseUp()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picThumb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, picThumb.Left + X - cmdDummy.Width, picThumb.Top + Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picView_Click
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Toggle the visibility state of the view icon.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picView_Click()
    Me.LayerViewable = Not (Me.LayerViewable)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picView_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picView_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picView_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyPress()
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picView_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : picView_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub picView_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Click
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for Click()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_DblClick
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for DblClick()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Initialize
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Load the images from the resource file.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Initialize()

'*' Fail through on local errors.
'*'
On Error Resume Next

    '*' Load the images from the resource file for the two different button types.
    '*'
    Set picTempView.Picture = LoadResPicture(101, vbResBitmap)
    Set picTempEdit.Picture = LoadResPicture(102, vbResBitmap)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_InitProperties
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for InitProperties()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_InitProperties()
    RaiseEvent InitProperties
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyDown()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyPress
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : KeyAscii (Integer)
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for KeyUp()
'*'
'*' Input     : KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseDown
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseDown()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X - cmdDummy.Width, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseMove
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseMove()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X - cmdDummy.Width, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseUp
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for MouseUp()
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X - cmdDummy.Width, Y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_ReadProperties
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for ReadProperties().
'*'
'*' Input     : PropBag (PropertyBag)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    RaiseEvent ReadProperties(PropBag)
    
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Caption")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    ThumbnailHeight = PropBag.ReadProperty("ThumbnailHeight", 16)
    ThumbnailWidth = PropBag.ReadProperty("ThumbnailWidth", 16)
    m_bolUseThumbnail = PropBag.ReadProperty("UseThumbnail", True)
    
    '*' Fire the resize event to make sure that the properties are in synch.
    '*'
    UserControl_Resize
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Resize
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for Resize().  Handles resize and display of local controls.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Resize()

'*' Fail through on local error.
'*'
On Error Resume Next

    '*' Set the size of the thumbnail to match the input or default sizes.
    '*'
    picThumb.Width = m_lngThumbnailWidth
    picThumb.Height = m_lngThumbnailHeight
    
    '*' Make sure that the panel is set to the full height.
    '*'
    cmdDummy.Height = UserControl.ScaleHeight
    
    '*' Position the dividing line in the panel.
    '*'
    picLine.Move picLine.Left, picLine.Top, picLine.Width, cmdDummy.Height - 6
    
    '*' Make sure that the images for the buttons are centered within the control.
    '*'
    picEdit.Top = (cmdDummy.Height - picEdit.Height) / 2
    picView.Top = picEdit.Top
    
    '*' Make sure that the caption is centered.
    '*'
    lblCaption.Top = (cmdDummy.Height - lblCaption.Height) / 2
    
    '*' Make sure that the thumbnail is centered.
    '*'
    picThumb.Top = (UserControl.ScaleHeight - picThumb.Height) / 2
    
    '*' Determine if the thumbnail is even visible.
    '*'
    If m_bolUseThumbnail Then
    
        '*' Show the thumbnail.
        '*'
        picThumb.Visible = True
                
        '*' Adjust the placement of the caption.
        '*'
        lblCaption.Left = picThumb.Left + picThumb.Width + 7
        
    Else
    
        '*' Hide the thumbnail.
        '*'
        picThumb.Visible = False
        
        '*' Adjust the placement of the caption.
        '*'
        lblCaption.Left = picThumb.Left
        
    End If
    
    '*' Adjust the border that would appear around the thumbnail.
    '*'
    shpBorder.Height = picThumb.Height + 2
    shpBorder.Width = picThumb.Width + 2
    shpBorder.Top = picThumb.Top - 1
    shpBorder.Left = picThumb.Left - 1
    
    '*' Adjust the seperator on the bottom of the control.
    '*'
    With linBottomSeperator
        .X1 = cmdDummy.Width
        .Y1 = UserControl.ScaleHeight - 1
        .X2 = UserControl.ScaleWidth
        .Y2 = .Y1
    End With
    
    '*' Adjust the seperator on the top of the control.
    '*'
    With linTopSeperator
        .X1 = cmdDummy.Width
        .Y1 = 0
        .X2 = UserControl.ScaleWidth
        .Y2 = .Y1
    End With
    
    '*' Adjust the position indicator on the bottom of the control.
    '*'
    With linBottomIndicator
        .X1 = cmdDummy.Width
        .Y1 = UserControl.ScaleHeight - 1
        .X2 = UserControl.ScaleWidth
        .Y2 = .Y1
    End With
    
    '*' Adjust the position indicator on the top of the control.
    '*'
    With linTopIndicator
        .X1 = cmdDummy.Width
        .Y1 = 0
        .X2 = UserControl.ScaleWidth
        .Y2 = .Y1
    End With
    
    '*' Adjust the top of the highlight based upon whether or not the top seperator is visible.  The goal is to
    '*' provide a single pixel of whitespace between any drawn borders and the highlight.
    '*'
    If linTopSeperator.Visible Then
        shpHighlight.Top = 2
    Else
        shpHighlight.Top = 1
    End If
    
    '*' Do the same for the bottom seperator.
    '*'
    If linBottomSeperator.Visible Then
        shpHighlight.Height = UserControl.ScaleHeight - shpHighlight.Top - 1
    Else
        shpHighlight.Height = UserControl.ScaleHeight - shpHighlight.Top
    End If
    
    '*' Make sure that the highlight is in the right position and extends to the edge of the control (with a single
    '*' pixel of whitespace).
    '*'
    shpHighlight.Left = cmdDummy.Width + 1
    shpHighlight.Width = UserControl.ScaleWidth - cmdDummy.Width - 1
    
    '*' Yield to any system-driven priorities.
    '*'
    DoEvents
    
    '*' Raise it to the caller.
    '*'
    RaiseEvent Resize
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Show
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for Show()
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_WriteProperties
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Event handler for WriteProperties()
'*'
'*' Input     : PropBag (PropertyBag)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ThumbnailHeight", m_lngThumbnailHeight, 16)
    Call PropBag.WriteProperty("ThumbnailWidth", m_lngThumbnailWidth, 16)
    Call PropBag.WriteProperty("UseThumbnail", m_bolUseThumbnail, True)
End Sub
