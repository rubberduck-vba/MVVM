VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a binding validation error."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object representing a binding validation error."
'@Interface
'@Exposed
Option Explicit

'@Description "The validation error message."
Public Property Get Message() As String
Attribute Message.VB_Description = "The validation error message."
End Property

'@Description "The binding that failed to validate."
Public Property Get Binding() As IPropertyBinding
Attribute Binding.VB_Description = "The binding that failed to validate."
End Property
