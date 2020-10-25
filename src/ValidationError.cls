VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a binding validation error."
'@Folder MVVM.Infrastructure.Validation
'@ModuleDescription "Represents a binding validation error."
'@PredeclaredId
Option Explicit

Implements IValidationError

Private Type TState
    Binding As IPropertyBinding
    Message As String
End Type

Private This As TState

'@Description "Creates a new instance of a validation error object."
Public Function Create(ByVal Binding As IPropertyBinding, ByVal Message As String) As IValidationError
Attribute Create.VB_Description = "Creates a new instance of a validation error object."
    GuardClauses.GuardNonDefaultInstance Me, ValidationError, TypeName(Me)
    Dim Result As ValidationError
    Set Result = New ValidationError
    Set Result.Binding = Binding
    Result.Message = Message
    Set Create = Result
End Function

'@Description "The validation error message."
Public Property Get Message() As String
Attribute Message.VB_Description = "The validation error message."
    Message = This.Message
End Property

Public Property Let Message(ByVal RHS As String)
    This.Message = RHS
End Property

'@Description "Gets the binding responsible for this validation error."
Public Property Get Binding() As IPropertyBinding
Attribute Binding.VB_Description = "Gets the binding responsible for this validation error."
    Set Binding = This.Binding
End Property

Public Property Set Binding(ByVal RHS As IPropertyBinding)
    Set This.Binding = RHS
End Property

Private Property Get IValidationError_Message() As String
    IValidationError_Message = This.Message
End Property

Private Property Get IValidationError_Binding() As IPropertyBinding
    Set IValidationError_Binding = This.Binding
End Property
