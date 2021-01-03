VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatterVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a format string in a notation compatible for use with VBA.Strings.Format$()."
'@Folder MVVM.Infrastructure.StringFormatting
'@ModuleDescription "An object representing a format string in a notation compatible for use with VBA.Strings.Format$()."
'@PredeclaredId
'@Exposed
'NOTE: 'VF: renamed to StringFormatterVB from VBStringFormatter (in parallel to StringFormatterNet
Option Explicit
Implements IStringFormatter

Private Type TState
    FormatString As String
End Type

Private This As TState

Public Function Create(ByVal FormatString As String) As IStringFormatter
    Dim Result As StringFormatterVB
    Set Result = New StringFormatterVB
    Result.FormatString = FormatString
    Set Create = Result
End Function

Public Property Get FormatString() As String
    FormatString = This.FormatString
End Property

Friend Property Let FormatString(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, StringFormatterVB
    This.FormatString = RHS
End Property

Private Function IStringFormatter_Format(ParamArray Values() As Variant) As String
    GuardClauses.GuardExpression LBound(Values) <> UBound(Values), TypeName(Me), "This implementation only supports formatting a single value at a time."
    Dim LocalArgs() As Variant
    If IsArray(Values(LBound(Values))) Then
        LocalArgs = Values(LBound(Values))
    Else
        ReDim LocalArgs(LBound(Values) To UBound(Values))
        Dim Index As Long
        For Index = LBound(Values) To UBound(Values)
            LocalArgs(Index) = Values(Index)
        Next
    End If
    'NOTE: to myself ToDo: 'VF: try out formatting string array like this separately
    IStringFormatter_Format = VBA.Strings.Format$(LocalArgs(LBound(LocalArgs)), This.FormatString)
End Function

Private Property Get IStringFormatter_FormatString() As String
    IStringFormatter_FormatString = This.FormatString
End Property
