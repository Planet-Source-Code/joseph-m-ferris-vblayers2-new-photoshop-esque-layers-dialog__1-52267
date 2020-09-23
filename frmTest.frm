VERSION 5.00
Object = "*\AVBLayers.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layer Properties"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   7350
      Top             =   2895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstEvents 
      Height          =   2985
      Left            =   3555
      TabIndex        =   10
      Top             =   45
      Width           =   4320
   End
   Begin VBLayers.VBLayerWindow VBLayerWindow1 
      Height          =   1995
      Left            =   45
      TabIndex        =   9
      Top             =   990
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3519
      ThumbnailHeight =   20
      ThumbnailWidth  =   20
   End
   Begin VB.CommandButton cmdRemoveSelected 
      Caption         =   "Remove"
      Height          =   435
      Left            =   1770
      TabIndex        =   8
      Top             =   3045
      Width           =   1710
   End
   Begin VB.CheckBox chkGeneric 
      Caption         =   "Preserve Transparency"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   750
      Width           =   3360
   End
   Begin VB.ComboBox cboGeneric2 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2430
      TabIndex        =   5
      Text            =   "100%"
      Top             =   315
      Width           =   765
   End
   Begin VB.ComboBox cboGeneric1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   4
      Text            =   "Normal"
      Top             =   315
      Width           =   1500
   End
   Begin VB.CommandButton cmdDummyDrop 
      Appearance      =   0  'Flat
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   225
   End
   Begin VB.CommandButton cmdAddTest 
      Caption         =   "Add"
      Height          =   435
      Left            =   15
      TabIndex        =   0
      Top             =   3045
      Width           =   1710
   End
   Begin VB.CheckBox chkLarge 
      Caption         =   "Use Large Layer Items"
      Height          =   285
      Left            =   3570
      TabIndex        =   11
      Top             =   3135
      Width           =   4260
   End
   Begin VB.Line linBreak4 
      BorderColor     =   &H80000014&
      X1              =   3495
      X2              =   3495
      Y1              =   0
      Y2              =   3495
   End
   Begin VB.Line linBreak3 
      BorderColor     =   &H80000010&
      X1              =   3480
      X2              =   3480
      Y1              =   0
      Y2              =   3495
   End
   Begin VB.Line linGenericShadow 
      BorderColor     =   &H80000010&
      X1              =   15
      X2              =   3450
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line linGenericHighlight 
      BorderColor     =   &H80000014&
      X1              =   15
      X2              =   3450
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label lblGenOpacity 
      Caption         =   "Opacity:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1785
      TabIndex        =   6
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paths"
      Height          =   180
      Left            =   1350
      TabIndex        =   3
      Top             =   30
      Width           =   1080
   End
   Begin VB.Label lblLayers 
      BackStyle       =   0  'Transparent
      Caption         =   "Layers"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   30
      Width           =   1080
   End
   Begin VB.Shape shpPaths 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   1245
      Top             =   15
      Width           =   1185
   End
   Begin VB.Shape shpLayersTab 
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   75
      Top             =   15
      Width           =   1185
   End
   Begin VB.Line linBreak2 
      X1              =   3240
      X2              =   3240
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line linBreak1 
      X1              =   3480
      X2              =   0
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************************'
'*'
'*' Module    : frmTest
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 03.09.2004
'*'
'*' Depends   : VBLayerWindow ActiveX Control
'*'
'*' Purpose   : Provides a test environment for the VBLayerWindow control.
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' Procedure : chkLarge_Click
'*'
'*'
'*' Date      : 03.09.2004
'*'
'*' Purpose   : Change the size of the thumbnail and item height.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub chkLarge_Click()

'*' Fail through on local error.
'*'
On Error Resume Next

    '*' Toggle between 40x32x32 and 26x16x16
    '*'
    If chkLarge.Value = Checked Then
        VBLayerWindow1.ThumbnailHeight = 32
        VBLayerWindow1.Thumbnailwidth = 32
        VBLayerWindow1.LayerItemHeight = 40
    Else
        VBLayerWindow1.ThumbnailHeight = 16
        VBLayerWindow1.Thumbnailwidth = 16
        VBLayerWindow1.LayerItemHeight = 26
    End If
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdAddTest_Click
'*'
'*'
'*' Date      : 03.09.2004
'*'
'*' Purpose   : Add an individual layer item that will consist of a prompted name and picture.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdAddTest_Click()

'*' Fail through on local error.
'*'
On Error Resume Next

Static s_lngCounter                     As Long

Dim strNewLayerName                     As String

    '*' Select a picture for the thumbnail.
    '*'
    cdgOpen.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
    cdgOpen.ShowOpen
   
    '*' Get a name for the new layer (dupes ok, goes off of id).
    '*'
    strNewLayerName = InputBox("Name for new layer:", "New Layer")
    
    '*' Add the item to the control.
    '*'
    VBLayerWindow1.LayerItems.Add strNewLayerName, "Item" & s_lngCounter, LoadPicture(cdgOpen.FileName)
        
    '*' Increment the local counter, which is used to populate the unique item key.
    '*'
    s_lngCounter = s_lngCounter + 1
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : cmdRemoveSelected_Click
'*'
'*'
'*' Date      : 03.09.2004
'*'
'*' Purpose   : Remove the selected item from the control, if one is selected.
'*'
'*' Input     : None.
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub cmdRemoveSelected_Click()

Dim layLocalItem As clsLayerItem

    '*' Iterate through the layers in the control.
    '*'
    For Each layLocalItem In VBLayerWindow1.LayerItems

        '*' Check to see if it is selected and remove it if it is.
        '*'
        If layLocalItem.Selected = True Then
            VBLayerWindow1.LayerItems.Remove (layLocalItem.Key)
            Exit For
        End If
    Next
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Event Tests
'*'
'**********************************************************************************************************************'

Private Sub VBLayerWindow1_CaptionChange(Index As Integer, Value As String)
    lstEvents.AddItem "Event:  CaptionChange(" & Index & ")"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_Click()
    lstEvents.AddItem "Event:  Click()"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_EditableChange(Index As Integer, Value As Boolean)
    lstEvents.AddItem "Event:  EditableChange(" & Index & ")"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_GotFocus()
    lstEvents.AddItem "Event:  GotFocus()"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_LostFocus()
    lstEvents.AddItem "Event:  LostFocus"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_Selection(Index As Integer)
    lstEvents.AddItem "Event:  Selection(" & Index & ")"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub

Private Sub VBLayerWindow1_ViewableChange(Index As Integer, Value As Boolean)
    lstEvents.AddItem "Event:  ViewableChange(" & Index & ")"
    lstEvents.Selected(lstEvents.ListCount - 1) = True
End Sub
