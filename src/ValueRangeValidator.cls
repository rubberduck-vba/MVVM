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
'@ModuleDescription "A validator that accepts values between set min and max values."
'@PredeclaredId
'@Exposed
Option Explicit
Option Compare Text
Implements IValueValidator

Private Type TState
    Min As Variant
    Max As Variant
    Compare As ValueRangeType
End Type

Public Enum ValueRangeType
    VariantDefault
    NumericValueRange
    StringValueRange
    DateValueRange
End Enum

Private This As TState

'@Description "Creates a new validator that accepts the specified range of values. By default, strings that look like numbers/dates are validated like numbers/dates."
Public Function Create(ByVal Min As Variant, ByVal Max As Variant, Optional ByVal Compare As ValueRangeType) As IValueValidator
    Dim Result As ValueRangeValidator
    Set Result = New ValueRangeValidator
    Result.Min = Min
    Result.Max = Max
    Result.Compare = Compare
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

Public Property Get Compare() As ValueRangeType
    Compare = This.Compare
End Property

Public Property Let Compare(ByVal RHS As ValueRangeType)
    This.Compare = RHS
End Property

Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    If IsObject(Value) Or IsArray(Value) Or IsEmpty(Value) Or IsError(Value) Then Exit Function
    Dim Result As Boolean
    
    If (This.Compare = VariantDefault And IsNumeric(Value) And IsNumeric(This.Min) And IsNumeric(This.Max)) _
    Or This.Compare = NumericValueRange Then
    
        Result = IsValidNumber(CDbl(Value))
    
    ElseIf (This.Compare = VariantDefault And IsDate(Value) And IsDate(This.Min) And IsDate(This.Max)) _
    Or This.Compare = DateValueRange Then
    
        Result = IsValidDate(CDate(Value))
    
    ElseIf (This.Compare = VariantDefault And VarType(Value) = vbString And VarType(This.Min) = vbString And VarType(This.Max) = vbString) _
    Or This.Compare = StringValueRange Then
    
        Result = IsValidString(CStr(Value))
    
    Else
    
        Result = Value >= This.Min And Value <= This.Max
        
    End If
    
    IValueValidator_IsValid = Result
End Function

Private Function IsValidNumber(ByVal Value As Double) As Boolean
    IsValidNumber = Value >= CDbl(This.Min) And Value <= CDbl(This.Max)
End Function

Private Function IsValidString(ByVal Value As String) As Boolean
    IsValidString = Value >= CStr(This.Min) And Value <= CStr(This.Max)
End Function

Private Function IsValidDate(ByVal Value As Date) As Boolean
    IsValidDate = Value >= CDate(This.Min) And Value <= CDate(This.Max)
End Function

Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = StringBuilder.AppendFormat("A valid value is between {0} and {1}.", CStr(This.Min), CStr(This.Max)).ToString
End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
