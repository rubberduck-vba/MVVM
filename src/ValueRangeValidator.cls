VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueRangeValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A validator that accepts values between set min and max values. Uses Text comparison for strings."
'@Folder MVVM.Common.Validators
'@ModuleDescription "A validator that accepts values between set min and max values. Uses Text comparison for strings."
'@PredeclaredId
'@Exposed
Option Explicit
Option Compare Text
Implements IValueValidator

Private Type TState
    Min As Variant
    Max As Variant
End Type

Private This As TState

Public Function Create(ByVal Min As Variant, ByVal Max As Variant) As IValueValidator
    Dim Result As ValueRangeValidator
    Set Result = New ValueRangeValidator
    Result.Min = Min
    Result.Max = Max
    Set Create = Result
End Function

Public Property Get Min() As Variant
    Min = This.Min
End Property

Public Property Let Min(ByVal RHS As Variant)
    This.Min = RHS
End Property

Public Property Get Max() As Variant
    Max = This.Max
End Property

Public Property Let Max(ByVal RHS As Variant)
    This.Max = RHS
End Property

Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    If IsObject(Value) Or IsArray(Value) Or IsEmpty(Value) Or IsError(Value) Then Exit Function
    IValueValidator_IsValid = Value >= This.Min And Value <= This.Max
End Function

Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = StringBuilderNet.AppendFormat("A valid value is between {0} and {1}.", CStr(This.Min), CStr(This.Max)).ToString
End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
