VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as a listener for data validation errors."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object that can be registered as a listener for data validation errors."
'@Interface
'@Exposed
Option Explicit

'@Description "A callback invoked when a validation error is added or removed from the validation manager."
Public Sub HandleValidationErrorsChanged(ByVal BindingPath As IBindingPath, Optional ByVal ValidationError As IValidationError)
Attribute HandleValidationErrorsChanged.VB_Description = "A callback invoked when a validation error is added or removed from the validation manager."
End Sub
