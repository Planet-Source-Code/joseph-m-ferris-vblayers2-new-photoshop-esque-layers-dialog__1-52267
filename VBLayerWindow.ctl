VERSION 5.00
Begin VB.UserControl VBLayerWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer timSynchCollections 
      Left            =   3960
      Top             =   750
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   0
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   294
      TabIndex        =   1
      Top             =   0
      Width           =   4410
      Begin VBLayers.LayerItem layLayerItem 
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   688
      End
   End
   Begin VB.VScrollBar vscMain 
      Height          =   1215
      Left            =   4515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "VBLayerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************************'
'*'
'*' Module    : VBLayerWindow
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.18.2004
'*'
'*' Depends   : LayerItem.ctl
'*'             FlatPicture.ctl
'*'
'*' Purpose   : Provides an ordered means for displaying and manipulating a collection of LayerItem controls
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' API Required Constant Declarations
'*'
'**********************************************************************************************************************'
Private Const HWND_TOPMOST              As Long = -1                '*' SetWindowPosition()
Private Const SM_CYVTHUMB               As Long = 9                 '*' GetSystemMetrics()
Private Const SWP_NOSIZE                As Long = &H1               '*' SetWindowPosition()

'**********************************************************************************************************************'
'*'
'*' API Required Type Declarations
'*'
'**********************************************************************************************************************'
Private Type POINTAPI                                               '*' GetCursorPos(), ScreenToClient()
    X                                   As Long
    Y                                   As Long
End Type

Private Type RECT                                                   '*' ClipCursor(), GetWindowRect(), InvalidateRect()
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

'**********************************************************************************************************************'
'*'
'*' API Declarations - kernel32.dll
'*'
'**********************************************************************************************************************'
Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)

'**********************************************************************************************************************'
'*'
'*' API Declarations - user32.dll
'*'
'**********************************************************************************************************************'
Private Declare Function ClipCursor Lib "user32" ( _
        lpRect As Any) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
        lpPoint As POINTAPI) As Long

Private Declare Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long

Private Declare Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As Long, _
        pRect As RECT) As Long

Private Declare Function InvalidateRect Lib "user32" ( _
        ByVal hwnd As Long, _
        lpRect As Any, _
        ByVal bErase As Long) As Long
    
Private Declare Function LockWindowUpdate Lib "user32" ( _
        ByVal hwndLock As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function ScreenToClient Lib "user32" ( _
        ByVal hwnd As Long, _
        lpPoint As Any) As Long
    
Private Declare Function SetCapture Lib "user32" ( _
        ByVal hwnd As Long) As Long
        
Private Declare Sub SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long)

'**********************************************************************************************************************'
'*'
'*' Private Constant Declarations
'*'
'**********************************************************************************************************************'
Private Const MIN_LAYER_HEIGHT          As Long = 26

'**********************************************************************************************************************'
'*'
'*' Private Constant Declarations (Error)
'*'
'**********************************************************************************************************************'
Private Const ERR_INVALIDKEY            As Long = vbObjectError + 25001
Private Const ERR_KEYNOTUNIQUE          As Long = vbObjectError + 25002

Private Const DES_INVALIDKEY            As String = "Invalid Key"
Private Const DES_KEYNOTUNIQUE          As String = "Key is not unique in collection"

'**********************************************************************************************************************'
'*'
'*' Private Member Declarations
'*'
'**********************************************************************************************************************'
Private m_bolCancelMove                 As Boolean
Private m_bolMoving                     As Boolean
Private m_bolPropChangeThroughCode      As Boolean
Private m_bolUseThumbnail               As Boolean
Private m_lngLayerItemHeight            As Long
Private m_lngLeftClip                   As Long
Private m_lngThumbnailHeight            As Long
Private m_lngThumbnailWidth             As Long
Private m_lngXOffset                    As Long
Private m_lngYOffset                    As Long
Private m_lyiLayerStack()               As LayerItems
Private m_pntOriginalLocation           As POINTAPI

'**********************************************************************************************************************'
'*'
'*' Public Member Declarations
'*'
'**********************************************************************************************************************'
Public LayerItems                       As clsLayerItems

'**********************************************************************************************************************'
'*'
'*' Private Type Declarations
'*'
'**********************************************************************************************************************'
Private Type LayerItems
    InternalIdentifier                  As String
    Caption                             As String
    Key                                 As String
    Tag                                 As String
    Picture                             As StdPicture
    Editable                            As Boolean
    Visible                             As Boolean
    Selected                            As Boolean
End Type

'**********************************************************************************************************************'
'*'
'*' Event Declarations
'*'
'**********************************************************************************************************************'
Event Click()
Event Selection(Index As Integer)
Event ViewableChange(Index As Integer, Value As Boolean)
Event EditableChange(Index As Integer, Value As Boolean)
Event CaptionChange(Index As Integer, Value As String)

'**********************************************************************************************************************'
'*'
'*' Procedure : LayerItemHeight
'*'
'*'
'*' Date      : 02.18.2004
'*'
'*' Purpose   : Set/Return the height of the layer items.
'*'
'*' Input     : Value (Long)
'*'
'*' Output    : LayerItemHeight (Long)
'*'
'**********************************************************************************************************************'
Public Property Get LayerItemHeight() As Long
    LayerItemHeight = m_lngLayerItemHeight
End Property

Public Property Let LayerItemHeight(Value As Long)
    
    '*' Make sure the minimum height is being used.
    '*'
    If Value < MIN_LAYER_HEIGHT Then
        Value = MIN_LAYER_HEIGHT
    End If
    
    '*' Set the value.
    '*'
    m_lngLayerItemHeight = Value
    PropertyChanged "LayerItemHeight"
    
    DistributeItems
    
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : ThumbnailHeight
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Set/Return the height of the thumbnail.
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
    
Dim lngCounter As Long

    m_lngThumbnailHeight = Value
    PropertyChanged "ThumbnailHeight"

    For lngCounter = 1 To layLayerItem.UBound
        layLayerItem(lngCounter).ThumbnailHeight = Value
    Next lngCounter

    DistributeItems
    
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : ThumbnailWidth
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Set/Return the width of the thumbnail.
'*'
'*' Input     : Value (Long)
'*'
'*' Output    : Thumbnail (Width)
'*'
'**********************************************************************************************************************'
Public Property Get ThumbnailWidth() As Long
       
    ThumbnailWidth = m_lngThumbnailWidth
      
End Property

Public Property Let ThumbnailWidth(Value As Long)

Dim lngCounter

    m_lngThumbnailWidth = Value
    PropertyChanged "ThumbnailWidth"
    
    For lngCounter = 1 To layLayerItem.UBound
        layLayerItem(lngCounter).ThumbnailWidth = Value
    Next lngCounter
    
    DistributeItems
    
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : UseThumbnail
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Get/Let visibility of the thumbnails of the layeritems.
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
    
Dim lngCounter As Long

    m_bolUseThumbnail = Value
    PropertyChanged "UseThumbnail"

    For lngCounter = 1 To layLayerItem.UBound
        layLayerItem(lngCounter).UseThumbnail = Value
    Next lngCounter

End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : DataChangedCallback
'*'
'*'
'*' Date      : 03.09.2004
'*'
'*' Purpose   : Provide a callback point where the child control can post changes to.
'*'
'*' Input     : m_strIdentifier (String)
'*'             ChangedProperty (String)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Public Sub DataChangedCallback(m_strIdentifier As String, ChangedProperty As String)

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim lngIndex                            As Long
    
    '*' Iterate through the layer items.
    '*'
    For lngIndex = 1 To LayerItems.Count
    
        '*' Check for a matching ID with the event signal.
        '*'
        If LayerItems(lngIndex).id = m_strIdentifier Then
        
            '*' Populate the appropriate property.
            '*'
            Select Case ChangedProperty
            
                Case "Caption"
                
                    m_lyiLayerStack(lngIndex - 1).Caption = LayerItems(lngIndex).Caption
                    layLayerItem(lngIndex).Caption = LayerItems(lngIndex).Caption
                
                    RaiseEvent CaptionChange(CInt(lngIndex), layLayerItem(lngIndex).Caption)
                    
                Case "LayerEditable"
                
                    m_lyiLayerStack(lngIndex - 1).Editable = LayerItems(lngIndex).LayerEditable
                    layLayerItem(lngIndex).LayerEditable = LayerItems(lngIndex).LayerEditable
                    
                Case "LayerViewable"
                
                    m_lyiLayerStack(lngIndex - 1).Visible = LayerItems(lngIndex).LayerViewable
                    layLayerItem(lngIndex).LayerViewable = LayerItems(lngIndex).LayerViewable
                    
                Case "Selected"
                
                    If Not m_bolPropChangeThroughCode Then
                
                        m_lyiLayerStack(lngIndex - 1).Selected = LayerItems(lngIndex).Selected
                        layLayerItem(lngIndex).Selected = LayerItems(lngIndex).Selected
                        
                    End If
                
            End Select
            
        Else
        
            Select Case ChangedProperty
            
                Case "Selected"
                
                    '*' Check to see if this property was changed via code in the usercontrol.  If it was, don't
                    '*' do anything.
                    '*'
                    If Not m_bolPropChangeThroughCode Then
                        
                        m_lyiLayerStack(lngIndex - 1).Selected = False
                        layLayerItem(lngIndex).Selected = False
                        
                    End If
                    
            End Select
            
        End If
        
    Next lngIndex
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : EventCallback
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Recieve callbacks from individual layeritem class objects.
'*'
'*' Input     : EventItem (clsLayerItem)
'*'             EventName (String)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub EventCallback(EventItem As clsLayerItem, EventName As String)
Attribute EventCallback.VB_UserMemId = -4
Attribute EventCallback.VB_MemberFlags = "40"

'*' Fail through on local errors.
'*'
On Error Resume Next

    '*' Check which event to perform on this clsLayerItem.
    '*'
    Select Case EventName
    
        Case "Add"
        
            '*' Add the class to the stack.
            '*"
            AddLayer EventItem
            
        Case "Remove"
        
            '*' Match and remove the item from the stack.
            '*'
            RemoveLayer EventItem
                    
    End Select
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_KeyDown
'*'
'*'
'*' Date      : 02.18.2004
'*'
'*' Purpose   : Watch for keyboard scrolling and selection movement.
'*'
'*' Input     : Index (Integer)
'*'             KeyCode (Integer)
'*'             Shift (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim lngIndex                            As Long
Dim lngFirst                            As Long
Dim lngLast                             As Long
    
    m_bolPropChangeThroughCode = True
    
    '*' Watch for specific key events (Arrows, Pages, Home, and End)
    '*'
    Select Case KeyCode
    
        Case 33                         '*' Page Up
        
            '*' Make sure that there is room for a full page up.
            '*'
            If Index > UserControl.ScaleHeight \ m_lngLayerItemHeight Then
            
                '*' Deselect the current item.
                '*'
                layLayerItem(Index).Selected = False
                LayerItems(Index).Selected = False
                m_lyiLayerStack(Index - 1).Selected = False
                
                '*' Select the new 'current' item.
                '*'
                layLayerItem(Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight)).Selected = True
                LayerItems(Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight)).Selected = True
                m_lyiLayerStack(Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight) - 1).Selected = True
                
                '*' Give focus to the same item.
                '*'
                layLayerItem(Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight)).SetFocus
                        
                '*' Move the scrollbar, if needed.
                '*'
                If ((Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight)) * m_lngLayerItemHeight) < Abs(vscMain.Value) + m_lngLayerItemHeight Then
                    vscMain.Value = -((Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight)) * m_lngLayerItemHeight) + m_lngLayerItemHeight
                End If

                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(Index - (UserControl.ScaleHeight \ m_lngLayerItemHeight))
                
            Else
            
                '*' If it can't move any further, bail.
                '*'
                If vscMain.Value = 0 Then
                    Exit Sub
                End If
                
                '*' Iterate from the second item to the end, deselecting all of the items.
                '*'
                For lngIndex = 2 To layLayerItem.UBound
                    layLayerItem(lngIndex).Selected = False
                    LayerItems(lngIndex).Selected = False
                    m_lyiLayerStack(lngIndex - 1).Selected = False
                Next lngIndex
                
                '*' Select the first item.
                '*'
                layLayerItem(1).Selected = True
                m_lyiLayerStack(0).Selected = True
                
                '*' Give it focus.
                '*'
                layLayerItem(1).SetFocus
                
                '*' Move to the top of the container.
                '*'
                vscMain.Value = 0
                
                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(1)
                
            End If
        
        Case 34                         '*' Page Down
        
            '*' Make sure that there is room to page down.
            '*'
            If Index < layLayerItem.UBound - (UserControl.ScaleHeight \ m_lngLayerItemHeight) Then
            
                '*' Deselect the current layer.
                '*'
                layLayerItem(Index).Selected = False
                LayerItems(Index).Selected = False
                m_lyiLayerStack(Index - 1).Selected = False
                
                '*' Select the new 'current' layer.
                '*'
                layLayerItem(Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight)).Selected = True
                LayerItems(Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight)).Selected = True
                m_lyiLayerStack(Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight) - 1).Selected = True
                
                '*' Give focus to the same item.
                '*'
                layLayerItem(Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight)).SetFocus
                
                '*' Move the scrollbar, if needed.
                '*'
                If UserControl.ScaleHeight - ((Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight)) * m_lngLayerItemHeight) < Abs(vscMain.Value) Then
                    vscMain.Value = ((UserControl.ScaleHeight - ((Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight)) * m_lngLayerItemHeight)))
                End If
                
                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(Index + (UserControl.ScaleHeight \ m_lngLayerItemHeight))
                
            Else
            
                '*' If it can't move any further, bail.
                '*'
                If vscMain.Value = vscMain.Max Then
                    Exit Sub
                End If
                    
                '*' Iterate from the first to the second to the last items.
                '*'
                For lngIndex = 1 To layLayerItem.UBound - 1
                
                    '*' Deselect the current layer.
                    '*'
                    layLayerItem(lngIndex).Selected = False
                    LayerItems(lngIndex).Selected = False
                    m_lyiLayerStack(lngIndex - 1).Selected = False
                    
                Next lngIndex
                
                '*' Select the last layer.
                '*'
                layLayerItem(layLayerItem.UBound).Selected = True
                LayerItems(LayerItems.Count).Selected = True
                m_lyiLayerStack(UBound(m_lyiLayerStack)).Selected = True
                
                '*' Give it focus.
                '*'
                layLayerItem(layLayerItem.UBound).SetFocus
                
                '*' Max out the scrollbar.
                '*'
                vscMain.Value = vscMain.Max
            
                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(LayerItems.Count)
                
            End If
            
        Case 35                         '*' End
        
            '*' Check to see if the control is already to the end.
            '*'
            If vscMain.Value = vscMain.Max Then
                Exit Sub
            End If
            
            '*' Iterate from the first to the second to last items.
            '*'
            For lngIndex = 1 To layLayerItem.UBound - 1
            
                '*' Deselect them.
                '*'
                layLayerItem(lngIndex).Selected = False
                LayerItems(lngIndex).Selected = False
                m_lyiLayerStack(lngIndex - 1).Selected = False
                
            Next lngIndex
            
            '*' Select the last item.
            '*'
            layLayerItem(layLayerItem.UBound).Selected = True
            LayerItems(LayerItems.Count).Selected = True
            m_lyiLayerStack(UBound(m_lyiLayerStack)).Selected = True
            
            '*' Give it focus.
            '*'
            layLayerItem(layLayerItem.UBound).SetFocus
            
            '*' Max out the scrollbar.
            '*'
            vscMain.Value = vscMain.Max
        
            '*' Trigger Selection.
            '*'
            RaiseEvent Selection(LayerItems.Count)
            
        Case 36                         '*' Home
        
            '*' Check if the user is already at the front.
            '*'
            If vscMain.Value = 0 Then
                Exit Sub
            End If
            
            '*' Iterate from the second through last items.
            '*'
            For lngIndex = 2 To layLayerItem.UBound
            
                '*' Deselect the current item.
                '*'
                layLayerItem(lngIndex).Selected = False
                LayerItems(lngIndex).Selected = False
                m_lyiLayerStack(lngIndex - 1).Selected = False
                
            Next lngIndex
            
            '*' Select the first item.
            '*'
            layLayerItem(1).Selected = True
            LayerItems(1).Selected = True
            m_lyiLayerStack(0).Selected = True
            
            '*' Give it focus.
            '*'
            layLayerItem(1).SetFocus
            
            '*' Minimize the value of the scrollbar.
            '*'
            vscMain.Value = 0
        
            '*' Trigger Selection.
            '*'
            RaiseEvent Selection(1)
            
        Case 38                         '*' Up Arrow
            
            '*' Make sure that it is not on the first item.
            '*'
            If Index > 1 Then
            
                '*' Deselect the current.
                '*'
                layLayerItem(Index).Selected = False
                LayerItems(Index).Selected = False
                m_lyiLayerStack(Index - 1).Selected = False
                
                '*' Select the previously sequential item.
                '*'
                layLayerItem(Index - 1).Selected = True
                LayerItems(Index - 1).Selected = True
                m_lyiLayerStack(Index - 2).Selected = True
                
                '*' Give it focus.
                '*'
                layLayerItem(Index - 1).SetFocus
                
                '*' Determine the visibility range.
                '*'
                CalculateVisible lngFirst, lngLast
                
                '*' Make sure that the current is within the visible range.
                '*'
                If Not (lngFirst < (Index - 1) And lngLast >= (Index - 1)) Then
                
                    '*' Out of range to the top.
                    '*'
                    If (Index - 1) <= lngFirst Then
                        vscMain.Value = -((Index - 2) * m_lngLayerItemHeight)
                    Else
                        
                        '*' Just set the scroll to match.
                        '*'
                        If lngFirst > 0 Then
                            vscMain.Value = UserControl.ScaleHeight - ((Index - 1) * m_lngLayerItemHeight)
                        Else

                            '*' Minimize the value, if it is not already.
                            '*'
                            If Not (vscMain.Value = 0) Then
                                vscMain.Value = 0
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(Index - 1)
                
            End If
            
        Case 40                         '*' Down Arrow
        
            '*' Make sure that it can move down a space.
            '*'
            If Index < layLayerItem.UBound Then
                
                '*' Deselect the current.
                '*'
                layLayerItem(Index).Selected = False
                LayerItems(Index).Selected = False
                m_lyiLayerStack(Index - 1).Selected = False
                
                '*' Select the next sequential item.
                '*'
                layLayerItem(Index + 1).Selected = True
                LayerItems(Index + 1).Selected = True
                m_lyiLayerStack(Index).Selected = True
                
                '*' Give it focus.
                '*'
                layLayerItem(Index + 1).SetFocus

                '*' Determine the visibility range of the control.
                '*'
                CalculateVisible lngFirst, lngLast
                
                '*' Check to see if it is outside of the range of the usercontrol.
                '*'
                If lngFirst > (Index + 1) Or lngLast < (Index + 1) Then
                    
                    '*' Scroll down.
                    '*'
                    If (Index + 1) >= lngLast Then
                        vscMain.Value = UserControl.ScaleHeight - ((Index + 1) * m_lngLayerItemHeight)
                    Else
                        vscMain.Value = -((Index) * m_lngLayerItemHeight)
                    End If
                End If
                
                '*' Trigger Selection.
                '*'
                RaiseEvent Selection(Index + 1)
                
            End If
                        
    End Select
        
    m_bolPropChangeThroughCode = False
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_MouseDown
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Begin tracking mouse movement for dragging and ensure selections.
'*'
'*' Input     : Index (Integer)
'*'             Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim pntPointStruct              As POINTAPI         '*' POINTAPI Structure for GetCursorPos()
Dim pntClip                     As POINTAPI
Dim rctWindow                   As RECT

Dim lngSubCounter               As Long
Dim liMain                      As clsLayerItem
        
    m_bolPropChangeThroughCode = True

    '*' Iterate through the stack.
    '*'
    For lngCounter = 0 To UBound(m_lyiLayerStack)
    
        '*' Deselect everything.
        '*'
        m_lyiLayerStack(lngCounter).Selected = False
        LayerItems(lngCounter + 1).Selected = False
        
    Next lngCounter
        
    GetCursorPos m_pntOriginalLocation
    
    '*' Iterate through all of the layer items (from a data aspect) that are known to exist.
    '*'
    For Each liMain In LayerItems
                        
        '*' Check for a match between the layeritem instance and the layeritem collection, via its unique id.
        '*'
        If liMain.id = layLayerItem(Index).InternalIdentifier Then
        
            '*' Iterate through the stack.
            '*'
            For lngSubCounter = 0 To UBound(m_lyiLayerStack)
                                
                '*' Check for a match to the layeritem.
                '*'
                If m_lyiLayerStack(lngSubCounter).InternalIdentifier = liMain.id Then
                    
                    '*' Select the current.
                    '*'
                    m_lyiLayerStack(lngSubCounter).Selected = True
                    liMain.Selected = True
                    Exit For
                    
                End If
            Next lngSubCounter
                                  
        Else
        
            '*' Make sure it is deselected.
            '*'
            liMain.Selected = False
            
        End If
        
    Next
    
    '*' Iterate through the layers.  (Note:  This is redundant.  LayerItems and the stack can be reconciled via
    '*' an index offset of -1.  This is to be removed at a later time.)
    '*'
    For lngCounter = 1 To layLayerItem.Count
        For lngSubCounter = 0 To UBound(m_lyiLayerStack)
            If layLayerItem(lngCounter).InternalIdentifier = m_lyiLayerStack(lngSubCounter).InternalIdentifier Then
                layLayerItem(lngCounter).Selected = m_lyiLayerStack(lngSubCounter).Selected
            End If
        Next lngSubCounter
    Next lngCounter
    
    '*' Flag the fact that the layer is going to begin moving.
    '*'
    m_bolMoving = True
        
    GetCursorPos m_pntOriginalLocation
    
    '*' Get the cursor position.
    '*'
    Call GetCursorPos(pntClip)
        
    '*' Set the point of clipping.
    '*'
    m_lngLeftClip = pntClip.X
    
    '*' Store the offset from the current mouse position.
    '*'
    m_lngXOffset = X
    m_lngYOffset = Y
    
    '*' Release capture from the OS.
    '*'
    ReleaseCapture
    
    '*' Set mouse capture to the ctlLayerItem.
    '*'
    SetCapture (layLayerItem(Index).hwnd)

    '*' Turn on the clipping.
    '*'
    EnableTrap
    
    '*' Make sure that the current item has focus.
    '*'
    layLayerItem(Index).SetFocus
    
    '*' Trigger selection.
    '*'
    RaiseEvent Selection(CLng(Index))
    
    m_bolPropChangeThroughCode = False
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_MouseMove
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Track dragging operations.
'*'
'*' Input     : Index (Integer)
'*'             Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' Fail through on local erros.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative loop counter.
Dim lngDropIndex                As Long             '*' Index that the user is over.
Dim pntPointStruct              As POINTAPI         '*' POINTAPI Structure for GetCursorPos
Dim rctWindowBounds             As RECT             '*' RECT Structure for GetWindowRect
Dim lngSubCounter               As Long

Static LastX                    As Long             '*' Track last known.
Static LastY                    As Long             '*' Track last known.

Dim lngX As Long

    '*' Turn off movement if there is not button pressed.  (Can't move if you don't drag.)
    '*'
    If Button = 0 Then
        m_bolMoving = False
        Exit Sub
    End If
    
    '*' Check to see if the user is in a drag state.
    '*'
    If m_bolMoving = True Then
        
        '*' Get the current position of the cursor.
        '*'
        Call GetCursorPos(pntPointStruct)
        
        '*' Assign current values and leave if the static is blank.
        '*'
        If LastY = 0 And LastX = 0 Then
            LastX = pntPointStruct.X
            LastY = pntPointStruct.Y
            Exit Sub
        End If
        
        '*' Check to see if there are statics and they are the same as last time.  Leave if they are.
        '*'
        If LastX = pntPointStruct.X And LastY = pntPointStruct.Y Then
            Exit Sub
        End If
            
        '*' Get the bounding region from the displayed picturebox that bounds the interal one.
        '*'
        Call GetWindowRect(picContainer.hwnd, rctWindowBounds)
        
        '*' Physically set the position, minding the offset also.
        '*'
        SetWindowPos layLayerItem(Index).hwnd, _
                     HWND_TOPMOST, _
                     pntPointStruct.X - rctWindowBounds.Left - m_lngXOffset, _
                     pntPointStruct.Y - rctWindowBounds.Top - m_lngYOffset, _
                     0, _
                     0, _
                     SWP_NOSIZE
                                          
        '*' Check to see if the user is in the five pixel hotspot at the top of the control.
        '*'
        If pntPointStruct.Y - 5 - Abs(vscMain.Value) < rctWindowBounds.Top Then
            
            '*' Attempt to move the scrollbar.
            '*'
            If vscMain.Value < vscMain.SmallChange Then
            
                '*' Implement a small change up.
                '*'
                vscMain.Value = vscMain.Value + 1
                            
                Call InvalidateRect(UserControl.hwnd, ByVal 0&, 1)
                UserControl.Refresh
                DoEvents
                Sleep 20
                
            End If
            
        '*' Check to see if the user is in the five pixel hotspot at the bottom of the control.
        '*'
        ElseIf pntPointStruct.Y + 5 > rctWindowBounds.Bottom - Abs(UserControl.ScaleHeight - picContainer.Height) Then
                    
            'LockWindowUpdate UserControl.hwnd
                        
            '*' Attempt to move the scrollbar.
            '*'
            If vscMain.Value > vscMain.Max - 1 Then
            
                '*' Implement a small change down.
                '*'
                vscMain.Value = vscMain.Value - 1
            
                '*' Force the usercontrol to redraw.
                '*'
                Call InvalidateRect(UserControl.hwnd, ByVal 0&, 1)
                UserControl.Refresh
                
                '*' Allow for the redraw to catch up.  (Prevents ghosting on drag scrolls.)
                '*'
                DoEvents
                Sleep 20
                
            End If
            
        End If
        
        '*' Get the item that the user is hovering over.
        '*'
        lngDropIndex = GetDropIndex(CLng(Index))
                                                                                        
        '*' This mess is the logic that determines the proper display of the drag indicators.  Don't play with this
        '*' code.  It works, and was conceptually the hardest to nail down.  ;-)
        '*'
        '*' Handle all exceptions first.  In order:
        '*'
        '*' 1.  Out of bounds (error)
        '*' 2.  Bottom half of previous item.
        '*' 3.  Current Item
        '*' 4.  Fist item on the second item.
        '*' 5.  Second item on the first item.
        '*' 6.  First item on the first item.
        '*'
        If (lngDropIndex = -2) Or _
           (Index > 1 And (lngDropIndex = Index - 1)) Or _
           (lngDropIndex = Index) Or _
           (lngDropIndex = 0 And Index = 2) Or _
           ((Index = 1) And (lngDropIndex = -1)) Or _
           ((Index = 1) And (lngDropIndex = 0)) Or _
           ((Index = 1) And (lngDropIndex = 1)) Then
        
            For lngCounter = 1 To layLayerItem.UBound
                layLayerItem(lngCounter).BottomIndicator = False
                layLayerItem(lngCounter).TopIndicator = False
            Next lngCounter
        
            Exit Sub
                    
        '*' Between first and second items.
        '*'
        ElseIf lngDropIndex = 0 Then
        
            layLayerItem(2).TopIndicator = True
            layLayerItem(1).BottomIndicator = True
            layLayerItem(2).BottomIndicator = False
            layLayerItem(1).TopIndicator = False
            
            For lngCounter = 3 To layLayerItem.UBound
                layLayerItem(lngCounter).BottomIndicator = False
                layLayerItem(lngCounter).TopIndicator = False
            Next lngCounter
            
        '*' Before first item.
        '*
        ElseIf lngDropIndex = -1 Then
                
            layLayerItem(1).TopIndicator = True
            layLayerItem(1).BottomIndicator = False
            
            For lngCounter = 2 To layLayerItem.UBound
                layLayerItem(lngCounter).BottomIndicator = False
                layLayerItem(lngCounter).TopIndicator = False
            Next lngCounter
            
        '*' After last item.
        '*'
        ElseIf lngDropIndex = UBound(m_lyiLayerStack) + 1 Then
                
            layLayerItem(layLayerItem.UBound).TopIndicator = False
            layLayerItem(layLayerItem.UBound).BottomIndicator = True
    
            For lngCounter = 1 To layLayerItem.UBound - 1
                layLayerItem(lngCounter).BottomIndicator = False
                layLayerItem(lngCounter).TopIndicator = False
            Next lngCounter
            
        '*' Everywhere else.
        '*'
        Else
            
            layLayerItem(lngDropIndex + 1).TopIndicator = True
            layLayerItem(lngDropIndex).BottomIndicator = True
            
            For lngCounter = 1 To layLayerItem.UBound
                If lngCounter = lngDropIndex Then
                    layLayerItem(lngDropIndex).TopIndicator = False
                ElseIf lngCounter = lngDropIndex + 1 Then
                    layLayerItem(lngDropIndex + 1).BottomIndicator = False
                Else
                    layLayerItem(lngCounter).TopIndicator = False
                    layLayerItem(lngCounter).BottomIndicator = False
                End If
            Next lngCounter
    
        End If
        
    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_MouseUp
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Update any drag movement of layer items.
'*'
'*' Input     : Index (Integer)
'*'             Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngCounter                  As Long
Dim lngNewIndex                 As Long             '*' Index of drop destination.

    m_bolPropChangeThroughCode = True
    
    '*' Reset the cursor watch.
    '*'
    m_pntOriginalLocation.X = -999
    m_pntOriginalLocation.Y = -999
    
    '*' Clear all of the highlight indicators.
    '*'
    For lngCounter = 1 To layLayerItem.UBound
        layLayerItem(lngCounter).TopIndicator = False
        layLayerItem(lngCounter).BottomIndicator = False
    Next lngCounter
        
    '*' Flag that there is not further dragging operation occurring.
    '*'
    m_bolMoving = False

    '*' Turn off the clipping rectangle.
    '*'
    Call ClipCursor(ByVal 0&)
           
    '*' Check for a cancelled move - 'Get Out of Jail Free'.
    '*'
    If m_bolCancelMove Then
        m_bolCancelMove = False
        Exit Sub
    End If
           
    '*' Hide the layer so that it doesn't look out of place.
    '*'
    layLayerItem(Index).Visible = False
                
    '*' Sync the VB .Top and .Left property of the control with its new position.
    '*'
    Call ForceUpdatePos(layLayerItem(Index))
    
    '*' Determine what the index of the new item will be, based upon where it was dropped.
    '*'
    lngNewIndex = GetDropIndex(CLng(Index))
        
    '*' Make sure that the new index and the item index are both valid indices.
    '*'
    If lngNewIndex > -2 And Index > -1 Then
            
        '*' Item moving from down to up, but not to front.
        '*'
        If (lngNewIndex < Index) And lngNewIndex > 0 Then
        
            '*' Offset by one for theoretical offset.
            '*'
            Call PopLayer(CLng(Index), lngNewIndex) ' + 1)
            
        '*' Moving from down to up to the front.
        '*'
        ElseIf (lngNewIndex < Index) And lngNewIndex = 0 Then
                
            If Not (Index = 1) Then
            
                '*' Swap layer indices.
                '*'
                Call PopLayer(CLng(Index), lngNewIndex + 1) ' - 1)
            
            End If
                        
        ElseIf lngNewIndex = -1 Then
        
            Call PopLayer(CLng(Index), lngNewIndex + 1)
                
        Else
        
            '*' Swap layer indices.
            '*'
            Call PopLayer(CLng(Index), lngNewIndex - 1)
            
        End If
            
    End If
                
    '*' Rebuild the display of the control.
    '*'
    DistributeItems
        
    m_bolPropChangeThroughCode = False
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_SelectionChange
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Update the selected item.
'*'
'*' Input     : Index (Integer)
'*'             Value (Boolean)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_SelectionChange(Index As Integer, Value As Boolean)

'*' Fail through on local errors.
'*'
On Error Resume Next

    '*' Check to see if the selection is out of range.
    '*'
    If Index > UBound(m_lyiLayerStack) Then
        Exit Sub
    End If
    
    '*' Set the value of the collection and display.
    '*'
    m_lyiLayerStack(Index - 1).Selected = Value
    LayerItems(Index).Selected = Value
    
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_SetEditable
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Set the value of the Editable Icon on the layer item.
'*'
'*' Input     : Index (Integer)
'*'             Value (Boolean)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_SetEditable(Index As Integer, Value As Boolean)

'*' Fail through on local errors.
'*'
On Error Resume Next

    'm_bolPropChangeThroughCode = True
    
    '*' Toggle the editable icon.
    '*'
    m_lyiLayerStack(Index - 1).Editable = Value
    
    RaiseEvent EditableChange(CLng(Index), Value)
    
    'm_bolPropChangeThroughCode = False
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : layLayerItem_SetVisible
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Set the value of the Visible Icon on the layer item.
'*'
'*' Input     : Index (Integer)
'*'             Value (Boolean)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub layLayerItem_SetVisible(Index As Integer, Value As Boolean)

'*' Fail through on local errors.
'*'
On Error Resume Next

    'm_bolPropChangeThroughCode = True

    '*' Toggle the visible icon.
    '*'
    m_lyiLayerStack(Index - 1).Visible = Value
    
    RaiseEvent ViewableChange(CLng(Index), Value)
    
    'm_bolPropChangeThroughCode = False
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Initialize
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Prepare the control for use.
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

    '*' Initialize the layer collection.
    '*'
    Set LayerItems = New clsLayerItems
    
    '*' Make sure that the layer collection is aware of the usercontrol.
    '*'
    Call LayerItems.Initialize(Me)
    
    '*' Set the mouse tracking coordinates to a generic default.
    '*'
    m_pntOriginalLocation.X = -999
    m_pntOriginalLocation.Y = -999
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_KeyPress
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Watch for keyboard events sent to the usercontrol.
'*'
'*' Input     : KeyAscii (Integer)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_KeyPress(KeyAscii As Integer)

'*' Fail through on local errors.
'*'
Dim lngIndex                            As Long
    
    '*' Check to see if the user has pressed the 'escape' key.
    '*'
    If KeyAscii = 27 Then
    
        '*' Cancel any movement caused by dragging layers.
        '*'
        m_bolMoving = False
        m_bolCancelMove = True
        
        '*' Iterate through the layer stack and turn off any drag indicators.
        '*'
        For lngIndex = 1 To layLayerItem.UBound
            layLayerItem(lngIndex).BottomIndicator = False
            layLayerItem(lngIndex).TopIndicator = False
        Next lngIndex
        
    End If

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_ReadProperties
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Pull persisted properties and try to clean up the ugly constituent scrollbar.
'*'
'*' Input     : PropBag (PropertyBag)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*' Fail through on local errors.
'*'
On Error Resume Next
        
    '*' Pull any persisted property values.
    '*'
    ThumbnailHeight = PropBag.ReadProperty("ThumbnailHeight", 16)
    ThumbnailWidth = PropBag.ReadProperty("ThumbnailWidth", 16)
    m_bolUseThumbnail = PropBag.ReadProperty("UseThumbnail", True)
    LayerItemHeight = PropBag.ReadProperty("LayerItemHeight", 26)
    
    '*' Check to see if the user is running in the IDE.
    '*'
    If UserControl.Ambient.UserMode Then
        
        '*' Set the handle to subclass.
        '*'
        g_lngTargetHwnd = UserControl.hwnd
        
        '*' ...line, and sinker.
        '*'
        modScrollFix.Hook
        
        '*' Make sure that the scrollbar is visible and disabled, by default.
        '*'
        vscMain.Visible = True
        vscMain.Enabled = False
        
    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Resize
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Resize child controls to the usercontrol constraints.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Resize()

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim lngCounter                          As Long

    '*' Make sure that the scrollbar uses Windows settings for its size.
    '*'
    vscMain.Move UserControl.ScaleWidth - GetSystemMetrics(SM_CYVTHUMB), 0, GetSystemMetrics(SM_CYVTHUMB), UserControl.ScaleHeight
    
    '*' Synch the container's width to the usercontrol.
    '*'
    picContainer.Move 0, picContainer.Top, UserControl.ScaleWidth - vscMain.Width, picContainer.Height
    
    '*' Synch the individual items witdth to the usercontrol.
    '*'
    For lngCounter = 0 To layLayerItem.UBound
        layLayerItem(lngCounter).Move 0, layLayerItem(lngCounter).Top, picContainer.Width
    Next lngCounter
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Terminate
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Clean up control on termination.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_Terminate()

'*' Fail through on local error.
'*'
On Error Resume Next

    '*' Check to see if the control is in the IDE.
    '*'
    If UserControl.Ambient.UserMode = True Then
        
        '*' Bail if there is an error.
        '*'
        If Not (Err.Number = 0) Then
            Exit Sub
        End If
        
        '*' Terminate Subsclassing.
        '*'
        modScrollFix.Unhook
        
    End If

    '*' Clear the layer stack.
    '*'
    Set LayerItems = Nothing
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_WriteProperties
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Persist control settings.
'*'
'*' Input     : PropBag (PropertyBag)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ThumbnailHeight", m_lngThumbnailHeight, 16)
    Call PropBag.WriteProperty("ThumbnailWidth", m_lngThumbnailWidth, 16)
    Call PropBag.WriteProperty("UseThumbnail", m_bolUseThumbnail, True)
    Call PropBag.WriteProperty("LayerItemHeight", m_lngLayerItemHeight, 26)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : vscMain_Change
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Synchronise the scroll bar and container.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub vscMain_Change()

'*' Fail through on local error.
'*'
On Error Resume Next

Dim lngIndex                            As Long

    '*' Synch the container position to the scrollbar.
    '*'
    picContainer.Top = vscMain.Value
                    
    '*' Find out if one of the items should have focus.
    '*'
    For lngIndex = 1 To layLayerItem.UBound
        If layLayerItem(lngIndex).Selected Then
            layLayerItem(lngIndex).SetFocus
            Exit Sub
        End If
    Next lngIndex
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : AddLayer
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Add a new layer to the layer stack.
'*'
'*' Input     : NewLayerItem (clsLayerItem)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub AddLayer( _
       NewLayerItem As clsLayerItem)

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim lngCounter                          As Long
Dim lngIndex                            As Long

    '*' Only validate the key if it is not a null string.
    '*'
    If Not NewLayerItem.Key = vbNullString Then
    
        '*' Verify the key uniqueness before allowing the addition.
        '*'
        If Not (-1 = Not m_lyiLayerStack) Then
        
            '*' Iterate through the existing keys.
            '*'
            For lngCounter = 0 To UBound(m_lyiLayerStack)
                
                '*' Check for a match.
                '*'
                If NewLayerItem.Key = m_lyiLayerStack(lngCounter).Key Then
                
                    '*' A duplicate key has been found.
                    '*'
                    Err.Raise ERR_KEYNOTUNIQUE, "AddLayer()", DES_KEYNOTUNIQUE
                    
                End If
                
            Next lngCounter
            
        End If
        
        '*' Verify that the key is not numeric.
        '*'
        If IsNumeric(NewLayerItem.Key) Then
        
            '*' Key is not numeric
            '*'
            Err.Raise ERR_INVALIDKEY, "AddLayer()", DES_INVALIDKEY
            
        End If
    
    End If
    
    '*' Check to see if the stack is initialized and size it to the correct bounds.
    '*'
    If -1 = Not m_lyiLayerStack Then
        ReDim m_lyiLayerStack(0)
    Else
        lngIndex = UBound(m_lyiLayerStack) + 1
        ReDim Preserve m_lyiLayerStack(lngIndex)
    End If
        
    '*' Load it up.
    '*'
    With m_lyiLayerStack(lngIndex)
        .Caption = NewLayerItem.Caption
        .Editable = NewLayerItem.LayerEditable
        .InternalIdentifier = NewLayerItem.id
        .Key = NewLayerItem.Key
        Set .Picture = NewLayerItem.Picture
        .Tag = NewLayerItem.Tag
        .Visible = NewLayerItem.LayerViewable
    End With
                
    '*' Make sure that the LayerItem controls are properly distributed.
    '*'
    DistributeItems
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : CalculateVisible
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Determine the first and last visible items.
'*'
'*' Input     : FirstVisible (Long)
'*'             LastVisible (Long)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub CalculateVisible(ByRef FirstVisible As Long, ByRef LastVisible As Long)

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim lngRange                            As Long

    '*' Calculate the first visible item by dividing the height of the layer into the size of the container.
    '*'
    FirstVisible = Abs(CLng(picContainer.Top / m_lngLayerItemHeight))
    
    '*' The range is determined by the clipping height of the usercontrol divided by the height of the layer.
    '*'
    lngRange = UserControl.ScaleHeight \ m_lngLayerItemHeight
    
    '*' Make sure that the last visible item is within bounds of the layer stack.
    '*'
    If FirstVisible + lngRange > layLayerItem.UBound Then
        LastVisible = layLayerItem.UBound
    Else
        LastVisible = FirstVisible + lngRange
    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : DistributeItems
'*'
'*'
'*' Date      : 03.04.2004
'*'
'*' Purpose   : Synchronise the layer stack with the display.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub DistributeItems()

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim lngCounter                          As Long
Dim lngIndexCount                       As Long
Dim sngTest As Single
Dim lngRetries  As Long
Dim lngSelected As Long
                              
    '*' Check to see if the stack has been initialized.
    '*'
    If -1 = Not m_lyiLayerStack Then
    
        '*' Clear any stragglers.
        '*'
        For lngCounter = 1 To layLayerItem.UBound
            layLayerItem(lngCounter).Visible = False
        Next lngCounter
        
        '*' Bail!
        '*'
        Exit Sub
    
    End If
            
    '*' Speed up display update.
    '*'
    LockWindowUpdate UserControl.hwnd
    
    '*' Get the stack count.
    '*'
    lngIndexCount = UBound(m_lyiLayerStack)
    
    '*' Remove after defaults...
    '*'
    If m_lngLayerItemHeight < 26 Then
        m_lngLayerItemHeight = 26
        LockWindowUpdate 0
    End If
                
    '*' Make sure that the container size matches the size needed for all visible items.
    '*'
    picContainer.Height = (lngIndexCount + 1) * m_lngLayerItemHeight
            
    '*' Iterate through the layer stack.
    '*'
    For lngCounter = 0 To UBound(m_lyiLayerStack)
    
        '*' Determine if there is a loaded item to use, or if a new one should be loaded.
        '*'
        If lngCounter >= layLayerItem.Count - 1 Then
            Load layLayerItem(lngCounter + 1)
        End If
        
        '*' Set the properties of the layer control.
        '*'
        With layLayerItem(lngCounter + 1)
        
            .UseThumbnail = m_bolUseThumbnail
            .ThumbnailHeight = m_lngThumbnailHeight
            .ThumbnailWidth = m_lngThumbnailWidth
            
            .Top = lngCounter * m_lngLayerItemHeight
            .Height = m_lngLayerItemHeight
            
            If Not .Visible Then
                .Visible = True
            End If
            
            .ShowBottomSeperator = lngCounter < UBound(m_lyiLayerStack)
                                    
            If Not .Caption = m_lyiLayerStack(lngCounter).Caption Then
                .Caption = m_lyiLayerStack(lngCounter).Caption
            End If
            
            If Not .LayerEditable = m_lyiLayerStack(lngCounter).Editable Then
                .LayerEditable = m_lyiLayerStack(lngCounter).Editable
            End If
            
            If Not .LayerViewable = m_lyiLayerStack(lngCounter).Visible Then
                .LayerViewable = m_lyiLayerStack(lngCounter).Visible
            End If
            
            .InternalIdentifier = m_lyiLayerStack(lngCounter).InternalIdentifier

            Set .Picture = m_lyiLayerStack(lngCounter).Picture
            
            If m_lyiLayerStack(lngCounter).Selected Then
                lngSelected = lngCounter + 1
            End If
            
            If Not .Selected = m_lyiLayerStack(lngCounter).Selected Then
                .Selected = m_lyiLayerStack(lngCounter).Selected
            End If
            
        End With
        
    Next lngCounter
    
    '*' Make sure that the items are visible.
    '*'
    For lngCounter = UBound(m_lyiLayerStack) + 2 To layLayerItem.UBound
        layLayerItem(lngCounter).Visible = False
    Next lngCounter
    
    '*' Set the initial scrollbar positions and visiblilty.
    '*'
    If picContainer.Height > UserControl.ScaleHeight Then
    
        vscMain.Max = UserControl.ScaleHeight - picContainer.Height
        vscMain.Enabled = True
        
        vscMain.LargeChange = UserControl.ScaleHeight
        vscMain.SmallChange = m_lngLayerItemHeight
        
    Else
    
        vscMain.Enabled = False
    
    End If
    
    '*' Select a layer item, if specified.
    '*'
    If lngSelected > 0 Then
        layLayerItem(lngSelected).SetFocus
    End If
    
    '*' Unlock display.
    '*'
    LockWindowUpdate 0
        
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : EnableTrap
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Ensure that the mouse is being clipped within relation to the current control.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub EnableTrap()

'*' Fail through on error.
'*'
On Error Resume Next

Dim lngResult                   As Long             '*' Function Return
Dim rctClipping                 As RECT             '*' RECT for ClipCursor
Dim rctResult                   As RECT             '*' RECT for GetWindowRect()
    
    '*' Get the rectangle of the clipping area.
    '*'
    Call GetWindowRect(picContainer.hwnd, rctResult)
    
    '*' Use the result RECT and the current mouse position (m_lngLeftClip) to set the rectangle to clip.
    '*'
    With rctClipping
      .Left = m_lngLeftClip
      .Top = rctResult.Top
      .Right = m_lngLeftClip + 1
      .Bottom = rctResult.Bottom
    End With
    
    '*' Clip it.
    '*'
    lngResult& = ClipCursor(rctClipping)
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : ForceUpdatePos
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Force a control to remain with the bounds of its parent
'*'
'*' Input     : ctlUnknown (Control)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub ForceUpdatePos(ctlUnknown As Control)

'*' Fail through on error.
'*'
On Error Resume Next

Dim rctControlBounds            As RECT             '*' Control Boundaries
    
    '*' Obtain the control's position in relation to its container.
    '*'
    Call GetWindowRect(ctlUnknown.hwnd, rctControlBounds)
    Call ScreenToClient(ctlUnknown.Container.hwnd, rctControlBounds.Left)
    
    '*' Call .Move() to physically "move" the control to its current location.  Fools VB into updating the property.
    '*'
    ctlUnknown.Move rctControlBounds.Left, rctControlBounds.Top
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : GetDropIndex
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Determine where a layer belongs, based upon the position of the mouse pointer.
'*'
'*' Input     : Index (Long)
'*'
'*' Output    : GetDropIndex (Long)
'*'
'**********************************************************************************************************************'
Private Function GetDropIndex(Index As Long) As Long

'*' Fail through on errors.
'*'
On Error Resume Next

Dim lngCounter                  As Long             '*' Iterative Loop Counter
Dim lngX                        As Long             '*' X Pos for Mouse
Dim lngY                        As Long             '*' Y Pos for Mouse
Dim pntCursorPos                As POINTAPI         '*' POINTAPI for GetCursorPos
Dim rctContainer                As RECT             '*' RECT for GetWindowRect

    '*' Get the mouse position and window rectangles.
    '*'
    Call GetWindowRect(picContainer.hwnd, rctContainer)
    Call GetCursorPos(pntCursorPos)
    
    '*' Check that the cursor is inside the horizontal constraints.
    '*'
    If rctContainer.Right > pntCursorPos.X And rctContainer.Left < pntCursorPos.X Then
    
        '*' Check that the cursor is inside the vertical constraints.
        '*'
        If rctContainer.Top - Abs(vscMain.Value) <= pntCursorPos.Y And rctContainer.Bottom >= pntCursorPos.Y Then
            
            '*' Get the X, Y coords of the cursor, relative to the control.
            '*'
            lngX = rctContainer.Right - pntCursorPos.X
            lngY = picContainer.Top + (pntCursorPos.Y - rctContainer.Top) + Abs(vscMain.Value)
                                    
            '*' Iterate through the layer items.
            '*'
            For lngCounter = 0 To UBound(m_lyiLayerStack) + 1
                
                '*' Check any index that isn't the dragged one.
                '*'
                If lngCounter <> Index Then
                    If layLayerItem(lngCounter).Top <= lngY And _
                       layLayerItem(lngCounter).Top + layLayerItem(lngCounter).Height >= lngY Then
                                                
                        '*' Determine if the user wants to drop before or after this object.
                        '*'
                        If layLayerItem(lngCounter).Top <= lngY And _
                           layLayerItem(lngCounter).Top + (0.5 * layLayerItem(lngCounter).Height) >= lngY Then
                        
                            '*' Top half drop.  Adjust it if it is above terminal (0).
                            '*'
                            GetDropIndex = lngCounter - 1
                            
                        Else
                                                                                
                            '*' Bottom half drop.
                            '*'
                            GetDropIndex = lngCounter
                    
                        End If
                                                
                        '*' Bail out.
                        '*'
                        Exit Function
                        
                    End If
                    
                End If
                
            Next lngCounter
            
       End If
        
    End If
    
    '*' Return a failing index.
    '*'
    GetDropIndex = -2
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : PopLayer
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Similar to the perl 'pop' function which will insert an item into an array at a specified location.
'*'
'*' Input     : CurrentIndex (Long)
'*'             NewIndex (Long)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Function PopLayer(CurrentIndex As Long, NewIndex As Long)

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim litTemp()                   As LayerItems       '*' Temp Data Store
Dim lngCounter                  As Long             '*' Iterative Counter
Dim lngIndex                    As Long             '*' Item Index
Dim lngIndexSort()              As Long             '*' Sorted Item Index
    
    '*' Adjust the index, due to the base one offset of the LayerItems collection.
    '*'
    CurrentIndex = CurrentIndex - 1
    
    '*' Initialize the Index
    '*'
    lngIndex = 0
    
    '*' Resize the Sorted Item Index to its final size.  (Minus one for the current element).
    '*'
    ReDim lngIndexSort(UBound(m_lyiLayerStack) - 1)
    
    '*' Iterate through the data.
    '*'
    For lngCounter = 0 To UBound(m_lyiLayerStack)
        
        '*' Store it if it isn't the current index.
        '*'
        If Not (CurrentIndex = lngCounter) Then
        
            lngIndexSort(lngIndex) = lngCounter
            lngIndex = lngIndex + 1
                    
        End If
        
    Next lngCounter
    
    '*' Resize the data store.
    '*'
    ReDim litTemp(UBound(m_lyiLayerStack))
    
    '*' Reset the index.
    '*'
    lngIndex = 0
    
    '*' Iterate through the data.
    '*'
    For lngCounter = 0 To UBound(m_lyiLayerStack)
    
        '*' Check for moved items.
        '*'
        If Not (lngCounter = NewIndex) Then
            
            '*' Write data to temp data store.
            '*'
            With litTemp(lngCounter)
                .Caption = m_lyiLayerStack(lngIndexSort(lngIndex)).Caption
                .Editable = m_lyiLayerStack(lngIndexSort(lngIndex)).Editable
                .Visible = m_lyiLayerStack(lngIndexSort(lngIndex)).Visible
                .Selected = m_lyiLayerStack(lngIndexSort(lngIndex)).Selected
                .InternalIdentifier = m_lyiLayerStack(lngIndexSort(lngIndex)).InternalIdentifier
                .Key = m_lyiLayerStack(lngIndexSort(lngIndex)).Key
                Set .Picture = m_lyiLayerStack(lngIndexSort(lngIndex)).Picture
                .Tag = m_lyiLayerStack(lngIndexSort(lngIndex)).Tag
            End With
        
            '*' Increment index.
            '*'
            lngIndex = lngIndex + 1
            
        Else
        
            '*' Write source index to its new home.
            '*'
            With litTemp(lngCounter)
                .Caption = m_lyiLayerStack(CurrentIndex).Caption
                .Editable = m_lyiLayerStack(CurrentIndex).Editable
                .Visible = m_lyiLayerStack(CurrentIndex).Visible
                .Selected = m_lyiLayerStack(CurrentIndex).Selected
                .InternalIdentifier = m_lyiLayerStack(CurrentIndex).InternalIdentifier
                .Key = m_lyiLayerStack(CurrentIndex).Key
                Set .Picture = m_lyiLayerStack(CurrentIndex).Picture
                .Tag = m_lyiLayerStack(CurrentIndex).Tag
                
            End With
        
        End If
            
    Next lngCounter
    
    '*' Set the buffer back to the stack.
    '*'
    m_lyiLayerStack = litTemp
            
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : RemoveLayer
'*'
'*'
'*' Date      : 02.27.2004
'*'
'*' Purpose   : Remove the currently specified layer from the stack and update the display.
'*'
'*' Input     : OldLayerItem (clsLayerItem)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub RemoveLayer( _
        OldLayerItem As clsLayerItem)
        
'*' Fail through on local error.
'*'
On Error Resume Next

Dim lngIndex                            As Long
Dim lngPosition                         As Long
Dim lyiLayerBuffer()                    As LayerItems

    '*' Check to see if there is even a reason to bother with redeclaring.
    '*'
    If Not UBound(m_lyiLayerStack) = 0 Then
            
        '*' Resize the buffer to match the size of the new stack.
        '*'
        ReDim lyiLayerBuffer(UBound(m_lyiLayerStack) - 1)
    
        '*' Iterate through the current stack.
        '*'
        For lngIndex = 0 To UBound(m_lyiLayerStack)
        
            '*' Check for a non-match for the removal.
            '*'
            If Not (m_lyiLayerStack(lngIndex).InternalIdentifier = OldLayerItem.id) Then
            
                '*' Transfer to the buffer.
                '*'
                lyiLayerBuffer(lngPosition) = m_lyiLayerStack(lngIndex)
                lngPosition = lngPosition + 1
                
            End If
            
        Next lngIndex
            
    End If
    
    '*' Assign the buffer back to the stack.
    '*'
    m_lyiLayerStack = lyiLayerBuffer
    
    '*' Make sure that the display is synched with the buffer.
    '*'
    DistributeItems
    
End Sub

