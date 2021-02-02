VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as a handler for IControlEvents callbacks."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object that can be registered as a handler for IControlEvents callbacks."
'@Interface
'@Exposed
Option Explicit

'@Description "A callback that handles the 'Enter' event."
Public Sub HandleEnter()
Attribute HandleEnter.VB_Description = "A callback that handles the 'Enter' event."
End Sub

'@Description "A callback that handles the 'Exit' event."
Public Sub HandleExit(ByRef Cancel As Boolean)
Attribute HandleExit.VB_Description = "A callback that handles the 'Exit' event."
End Sub

'@Description "A callback that handles the 'AfterUpdate' event."
Public Sub HandleAfterUpdate()
Attribute HandleAfterUpdate.VB_Description = "A callback that handles the 'AfterUpdate' event."
End Sub

'@Description "A callback that handles the 'BeforeUpdate' event."
Public Sub HandleBeforeUpdate(ByRef Cancel As Boolean)
Attribute HandleBeforeUpdate.VB_Description = "A callback that handles the 'BeforeUpdate' event."
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
'@Description "A callback that handles the 'Change' event."
Public Sub HandleChange()
End Sub
