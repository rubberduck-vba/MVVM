VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for validating the value before a binding is applied."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object responsible for validating the value before a binding is applied."
'@Interface
'@Exposed
Option Explicit

'@Description "Using this validator in a property binding will coerce it into using this update trigger."
Public Property Get Trigger() As BindingUpdateSourceTrigger
Attribute Trigger.VB_Description = "Using this validator in a property binding will coerce it into using this update trigger."
End Property

'@Description "A user-friendly message describing the validation rule(s)."
Public Property Get Message() As String
Attribute Message.VB_Description = "A user-friendly message describing the validation rule(s)."
End Property

'@Description "True if the specified value is valid in itself or in the context of the binding source and target."
Public Function IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
Attribute IsValid.VB_Description = "True if the specified value is valid in itself or in the context of the binding source and target."
End Function
