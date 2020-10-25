VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An observer that can notify registered handlers when applying a binding throws a validation error."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An observer that can notify registered handlers when applying a binding throws a validation error."
'@Interface
'@Exposed
Option Explicit

'@Description "Registers the specified handler."
Public Sub RegisterHandler(ByVal Handler As IHandleValidationError)
Attribute RegisterHandler.VB_Description = "Registers the specified handler."
End Sub

'@Description "Notifies registered handlers."
Public Sub OnValidationErrorsChanged(ByVal BindingContext As Object, Optional ByVal ValidationError As IValidationError)
Attribute OnValidationErrorsChanged.VB_Description = "Notifies registered handlers."
End Sub
