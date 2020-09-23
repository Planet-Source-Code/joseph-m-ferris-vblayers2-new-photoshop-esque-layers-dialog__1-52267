VERSION 5.00
Begin VB.UserControl FlatPicture 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   Begin VB.Line linBottom 
      BorderColor     =   &H80000014&
      X1              =   285
      X2              =   4140
      Y1              =   3825
      Y2              =   3825
   End
   Begin VB.Line linTop 
      BorderColor     =   &H80000010&
      X1              =   150
      X2              =   4440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linRight 
      BorderColor     =   &H80000014&
      X1              =   4665
      X2              =   4665
      Y1              =   0
      Y2              =   3825
   End
   Begin VB.Line linLeft 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3825
   End
End
Attribute VB_Name = "FlatPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************************************'
'*'
'*' Module    : FlatPicture
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.26.2004
'*'
'*' Depends   : None.
'*'
'*' Purpose   : Quick and dirty flat control that provides an 'inset' container.
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' Event Declarations
'*'
'**********************************************************************************************************************'
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event HitTest(x As Single, y As Single, HitResult As Integer)

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    RaiseEvent HitTest(x, y, HitResult)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseMove
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Map to the UserControl's MouseMove.
'*'
'*' Input     : Button
'*'
'*' Output    :
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_MouseUp
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Map to the UserControl's MouseUp event.
'*'
'*' Input     : Button (Integer)
'*'             Shift (Integer)
'*'             X (Single)
'*'             Y (Single)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_ReadProperties
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Map to the UserControl's ReadProperty method.
'*'
'*' Input     : Propbag (PropertyBag)
'*'
'*' Output    :
'*'
'**********************************************************************************************************************'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_Resize
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Resize the 'borders' on the control.
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

    '*' Top line.
    '*'
    With linTop
        .X1 = 0
        .X2 = UserControl.ScaleWidth
        .Y1 = 0
        .Y2 = 0
    End With
    
    '*' Bottom line.
    '*'
    With linBottom
        .X1 = 0
        .X2 = UserControl.ScaleWidth
        .Y1 = UserControl.ScaleHeight - 15
        .Y2 = .Y1
    End With
    
    '*' Left line.
    '*"
    With linLeft
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = UserControl.ScaleHeight - 15
    End With
    
    '*' Right line.
    '*'
    With linRight
        .X1 = UserControl.ScaleWidth - 15
        .Y1 = 0
        .X2 = .X1
        .Y2 = UserControl.ScaleHeight - 15
    End With
    
End Sub

'**********************************************************************************************************************'
'*'
'*' Procedure : UserControl_WriteProperties
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Map to the UserControl's WriteProperty method.
'*'
'*' Input     : PropBag (PropertyBag)
'*'
'*' Output    : None.
'*'
'**********************************************************************************************************************'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
End Sub
