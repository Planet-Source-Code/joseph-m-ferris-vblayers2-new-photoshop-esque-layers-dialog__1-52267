VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayerItemsOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsLayerItems"
Attribute VB_Ext_KEY = "Member0" ,"clsLayerItems"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************************************************************************'
'*'
'*' Module    : LayerItems
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.23.2004
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
'*' Private Constant Declarations
'*'
'**********************************************************************************************************************'
Private Const ERR_INVALIDLAYERWINDOW    As Long = vbObjectError + 22001

'**********************************************************************************************************************'
'*'
'*' Private Member Declarations
'*'
'**********************************************************************************************************************'
Private m_colLayerStack                 As Collection
Private m_objUserControl                As Object

'**********************************************************************************************************************'
'*'
'*' Property  : Count (Read-Only)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Return a count of the individual LayerItems that are currently in the LayerItems collection.
'*'
'*' Input     : None.
'*'
'*' Output    : Count (Long)
'*'
'**********************************************************************************************************************'
Public Property Get Count() As Long
    Count = m_colLayerStack.Count
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : Item (Read-Only)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Provides a redirection to the individual LayerItem.  Variant declaration allows for retrieval either
'*'             via Key or via Index.
'*'
'*' Input     : vntIndexKey (Variant)
'*'
'*' Output    : Item (LayerItem)
'*'
'**********************************************************************************************************************'
Public Property Get Item(vntIndexKey As Variant) As LayerItem
  Set Item = m_colLayerStack(vntIndexKey)
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : NewEnum (Read-Only, Hidden)
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Allows For...Each syntax to be utilized through the IUnknown interface.
'*'
'*' Input     : None
'*'
'*' Output    : NewEnum (IUnknown)
'*'
'**********************************************************************************************************************'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -5
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colLayerStack.[_NewEnum]
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : Add
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Adds a new LayerItem to the LayerItems collection.
'*'
'*' Input     : In Flux.  ;-)
'*'
'*' Output    : LayerItem
'*'
'**********************************************************************************************************************'
Public Function Add(Caption As String, Optional Key As String, Optional Picture As StdPicture, _
       Optional LayerEditable As Boolean = False, Optional LayerViewable As Boolean = False, _
       Optional Tag As String = vbNullString) As clsLayerItem

Dim objNewMember                        As clsLayerItem
    
    '*' Instanciate the new class member.  Even though the variable is only scoped for this function, it is stored
    '*' in a collection with member level scope.  It will exist for the lifetime of the class.
    '*'
    Set objNewMember = New clsLayerItem

    '*' Set any Layer Properties directly to the LayerItem. as needed.  Note that the ID is set last.  This way, any
    '*' callback events will not be fired from the LayerItem class, since it requires the internal ID as a
    '*' 'finalization' of the intial assignments.
    '*'
    With objNewMember
        .Key = Key
        .Caption = Caption
        .LayerEditable = LayerEditable
        .LayerViewable = LayerViewable
        Set .LayerWindowCtl = m_objUserControl
        Set .Picture = Picture
        .Tag = Tag
        .id = CreateGUID
    End With
    
    '*' Check to see if a key has been provided by the user.
    '*'
    If Len(Key) = 0 Then
        m_colLayerStack.Add objNewMember
    Else
        m_colLayerStack.Add objNewMember, Key
    End If

    '*' Signal the event to the parent.
    '*'
    m_objUserControl.EventCallback objNewMember, "Add"
    
    '*' Return the newly created LayerItem and destroy the local reference.
    '*'
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : Initialize
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Attach an instance of the LayerWindow to the class for forward population to the LayerItem to implement
'*'             a quasi-callback ability so that multiple control items can "raise" the same "events".
'*'
'*' Input     : LayerWindowCtl (Object)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub Initialize(LayerWindowCtl As Object)

    '*' Make sure that a proper VBLayerWindow Object is passed.
    '*'
    If Not TypeName(LayerWindowCtl) = "VBLayerWindow" Then
    
        '*' Raise an error for an improper control type.
        '*'
        Err.Raise ERR_INVALIDLAYERWINDOW, "Initialize()"
    
    End If
        
    '*' Store the reference to the object for a quasi-callback interface.
    '*'
    Set m_objUserControl = LayerWindowCtl
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : Remove
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Remove an item from the stack with either the Key or the Index.
'*'
'*' Input     : vntIndexKey (Variant)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub Remove(vntIndexKey As Variant)
    m_colLayerStack.Remove vntIndexKey
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : Class_Initialize
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Initialize the collection that is used to hold the LayerItems.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub Class_Initialize()
    Set m_colLayerStack = New Collection
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : Class_Terminate
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Destroy the collection that is used to hold the LayerItems
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub Class_Terminate()
    Set m_colLayerStack = Nothing
End Sub


