VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayerItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsLayerItem"
Attribute VB_Ext_KEY = "Member0" ,"clsLayerItem"
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
'*' Depends   : None.
'*'
'*' Purpose   : Collection class.  Provides a LayerItem interface, similar to a ListItem for a Listview control.
'*'
'*' Notes     : 1.  Actions that require callback functionality are Add() and Remove().  Both execute the
'*'                 EventCallback() method of the VBLayerWindow instance.
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
'*' Procedure : Add
'*'
'*'
'*' Date      : 02.23.2004
'*'
'*' Purpose   : Adds a new LayerItem to the LayerItems collection.
'*'
'*' Input     : Caption (String)
'*'             Key (String)
'*'             Picture (StdPicture)
'*'             LayerEditable (Boolean)
'*'             LayerViewable (Boolean)
'*'             Tag (String)
'*'
'*' Output    : LayerItem
'*'
'**********************************************************************************************************************'
Public Function Add( _
       Caption As String, _
       Optional Key As String, _
       Optional Picture As StdPicture, _
       Optional LayerEditable As Boolean = False, _
       Optional LayerViewable As Boolean = False, _
       Optional Tag As String = vbNullString) As clsLayerItem

'*' Raise errors directly, via relay.  This will allow the collection errors to be directly raised to the calling
'*' source.
'*'
On Error GoTo LocalHandler

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

Exit Function

'*' Error handling stub.
'*'
LocalHandler:

    '*' Just pass it on.
    '*'
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'**********************************************************************************************************************'
'*'
'*' Procedure : Item
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Return a singular clsLayerItem object through either an Index reference or a Key reference.
'*'
'*' Input     : vntIndexKey (Variant)
'*'
'*' Output    : Item (clsLayerItem)
'*'
'**********************************************************************************************************************'
Public Property Get Item(vntIndexKey As Variant) As clsLayerItem
Attribute Item.VB_UserMemId = 0
  
'*' Handle error locally.
'*'
On Error GoTo LocalHandler

    '*' Pull it from the stack.
    '*'
    Set Item = m_colLayerStack(vntIndexKey)
    
Exit Property

'*' Error handling stub.
'*'
LocalHandler:

    '*' Pass it along to the caller.
    '*'
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : Count
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Returns a count of items in the collection
'*'
'*' Input     : None
'*'
'*' Output    : Count (Long)
'*'
'**********************************************************************************************************************'
Public Property Get Count() As Long
    
'*' Handle error locally.
'*'
On Error GoTo LocalHandler
    
    '*' Return the count of the collection.
    '*'
    Count = m_colLayerStack.Count

Exit Property

'*' Error handling stub.
'*'
LocalHandler:

    '*' Pass it along to the caller.
    '*'
    Err.Raise Err.Number, Err.Source, Err.Description

End Property

'**********************************************************************************************************************'
'*'
'*' Procedure : Remove
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Remove an item from the collection of clsLayerItem objects
'*'
'*' Input     : vntIndexKey (Variant)
'*'
'*' Output    : None
'*'
'**********************************************************************************************************************'
Public Sub Remove(vntIndexKey As Variant)

'*' Handle error locally.
'*'
On Error GoTo LocalHandler

    '*' Signal the event to the parent.
    '*'
    m_objUserControl.EventCallback m_colLayerStack.Item(vntIndexKey), "Remove"

    '*' Remove it from the array.
    '*'
    m_colLayerStack.Remove vntIndexKey

Exit Sub

'*' Error handling stub.
'*'
LocalHandler:

    '*' Pass it along to the caller.
    '*'
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : NewEnum
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Allows for enumeration in the For Each...Next syntax.
'*'
'*' Input     : None
'*'
'*' Output    : NewEnum (IUnknown)
'*'
'**********************************************************************************************************************'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

'*' Fail through on local error.
'*'
On Error Resume Next

    '*' Allow the callback to retrieve the implied item.
    '*'
    Set NewEnum = m_colLayerStack.[_NewEnum]

End Property

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

'*' Handle errors locally.  The sub purposely watches for an error to occur in the form of not having the VBLayerWindow
'*' attached to this class instance.
'*'
On Error GoTo LocalHandler

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
    
Exit Sub

'*' Error handling stub.
'*'
LocalHandler:

    '*' Pass it along to the caller.
    '*'
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : Class_Initialize
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Initialize the collection.
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
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Destroy the collection.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub Class_Terminate()
    Set m_colLayerStack = Nothing
End Sub
