VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************************************************************************'
'*'
'*' Module    : LayerItem
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.23.2004
'*'
'*' Depends   : None.
'*'
'*' Purpose   : Provides a single object to be contained within the clsLayerItems collection class.  Meant to be a
'*'             static class-level representation of the objects within a LayerItem usercontrol array.
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' Private Member Declarations
'*'
'**********************************************************************************************************************'
Private m_bolBottomIndicator            As Boolean
Private m_bolLayerEditable              As Boolean
Private m_bolLayerViewable              As Boolean
Private m_bolSelected                   As Boolean
Private m_bolShowBottomSeperator        As Boolean
Private m_bolShowTopSeperator           As Boolean
Private m_bolTopIndicator               As Boolean
Private m_objVBLayerWindow              As Object
Private m_spcPicture                    As StdPicture
Private m_strCaption                    As String
Private m_strKey                        As String
Private m_strIdentifier                 As String
Private m_strTag                        As String

'**********************************************************************************************************************'
'*'
'*' Property  : BottomIndicator (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag which indicates whether a drag event is visible on the bottom of the LayerItem
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : BottomIndicator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get BottomIndicator() As Boolean
    BottomIndicator = m_bolBottomIndicator
End Property
Public Property Let BottomIndicator(Value As Boolean)
    m_bolBottomIndicator = Value
    SignalNotify "BottomIndicator"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Caption (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a caption for display on LayerItem
'*'
'*' Input     : Value (String)
'*'
'*' Output    : Caption (String)
'*'
'**********************************************************************************************************************'
Public Property Get Caption() As String
    Caption = m_strCaption
End Property
Public Property Let Caption(Value As String)
    m_strCaption = Value
    SignalNotify "Caption"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Key (Public Read-Only)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides protected Key storage that mirror the LayerItems collection key, if provided.
'*'
'*' Input     : Value (String)
'*'
'*' Output    : Key (String)
'*'
'**********************************************************************************************************************'
Public Property Get Key() As String
    Key = m_strKey
End Property
Friend Property Let Key(Value As String)
    m_strKey = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : ID (Protected)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides an internal identifier that can be used to verify the uniqueness of an individual item.
'*'
'*' Input     : Value (String)
'*'
'*' Output    : ID (String)
'*'
'**********************************************************************************************************************'
Friend Property Get id() As String
    id = m_strIdentifier
End Property
Friend Property Let id(Value As String)
    m_strIdentifier = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : LayerEditable (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag that indicates whether the layer is considered editable.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : LayerEditable (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get LayerEditable() As Boolean
    LayerEditable = m_bolLayerEditable
End Property
Public Property Let LayerEditable(Value As Boolean)
    m_bolLayerEditable = Value
    SignalNotify "LayerEditable"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : LayerViewable (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag that indicates whether the layer is considered visible/viewable.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : LayerViewable (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get LayerViewable() As Boolean
    LayerViewable = m_bolLayerViewable
End Property
Public Property Let LayerViewable(Value As Boolean)
    m_bolLayerViewable = Value
    SignalNotify "LayerViewable"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : LayerWindowCtl (Protected)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides an object reference to post back "events" through the SignalNotify() method of this class.  It
'*'             is readily assumed that the propery TypeName ('VBLayerWindow') has been checked prior to assignment.
'*'
'*' Input     : Value (Object)
'*'
'*' Output    : LayerWindowCtl (Object)
'*'
'**********************************************************************************************************************'
Friend Property Get LayerWindowCtl() As Object
    Set LayerWindowCtl = m_objVBLayerWindow
End Property
Friend Property Set LayerWindowCtl(Value As Object)
    Set m_objVBLayerWindow = Value
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Picture (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a stored StdPicture object that is used in rendering the thumbnail in the LayerItem control/
'*'
'*' Input     : Value (StdPicture)
'*'
'*' Output    : Picture (StdPicture)
'*'
'**********************************************************************************************************************'
Public Property Get Picture() As StdPicture
    Set Picture = m_spcPicture
End Property
Public Property Set Picture(Value As StdPicture)
    Set m_spcPicture = Value
    SignalNotify "LayerViewable"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Selected (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag that indicates whether the control believes that it is Selected.
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : Selected (Value)
'*'
'**********************************************************************************************************************'
Public Property Get Selected() As Boolean
    Selected = m_bolSelected
End Property
Public Property Let Selected(Value As Boolean)
    m_bolSelected = Value
    SignalNotify "Selected"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : ShowBottomSeperator (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag which indicates whether bottom line is displayed on LayerItem
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : ShowBottomSeperator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get ShowBottomSeperator() As Boolean
    ShowBottomSeperator = m_bolShowBottomSeperator
End Property
Public Property Let ShowBottomSeperator(Value As Boolean)
    m_bolShowBottomSeperator = Value
    SignalNotify "ShowBottomSeperator"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : ShowTopSeperator (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag which indicates whether bottom line is displayed on LayerItem
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : ShowTopSeperator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get ShowTopSeperator() As Boolean
    ShowTopSeperator = m_bolShowTopSeperator
End Property
Public Property Let ShowTopSeperator(Value As Boolean)
    m_bolShowTopSeperator = Value
    SignalNotify "ShowTopSeperator"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : Tag (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a storage space for miscellaneous data.
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
    SignalNotify "Tag"
End Property

'**********************************************************************************************************************'
'*'
'*' Property  : TopIndicator (Public)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a flag which indicates whether a drag event is visible on the top of the LayerItem
'*'
'*' Input     : Value (Boolean)
'*'
'*' Output    : TopIndicator (Boolean)
'*'
'**********************************************************************************************************************'
Public Property Get TopIndicator() As Boolean
    TopIndicator = m_bolTopIndicator
End Property
Public Property Let TopIndicator(Value As Boolean)
    m_bolTopIndicator = Value
    SignalNotify "TopIndicator"
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : SignalNotify
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Forcibly send the object Key and the Name of the ChangedProperty to the parent.
'*'
'*' Input     : ChangedProperty (String)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Private Sub SignalNotify(ChangedProperty As String)

'*' Fail through on local error.
'*'
On Error Resume Next

    '*' Only signal a change if the ID has been assigned.  An unassigned ID is equivalent of a non-finalized assignment
    '*' to the object.
    '*'
    If Not (id = vbNullString) Then
    
        '*' Take one down, pass it around...
        '*'
        Call m_objVBLayerWindow.DataChangedCallback(m_strIdentifier, ChangedProperty)
    
    End If
    
End Sub

