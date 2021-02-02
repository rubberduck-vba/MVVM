VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatterNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a format string in a notation compatible for use with .NET String.Format()."
'@Folder MVVM.Infrastructure.StringFormatting
'@ModuleDescription "An object representing a format string in a notation compatible for use with .NET String.Format()."
'@PredeclaredId
'@Exposed
'NOTE: 'VF: renamed to StringFormatterNet from StringFormatter in parallel to StringFormatterVB
Option Explicit
Implements IStringFormatter

Private Type TState
    FormatString As String
End Type

Private This As TState

Public Function Create(ByVal FormatString As String) As IStringFormatter
    Dim Result As StringFormatterNet
    Set Result = New StringFormatterNet
    Result.FormatString = FormatString
    Set Create = Result
End Function

Public Property Get FormatString() As String
    FormatString = This.FormatString
End Property

Friend Property Let FormatString(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, StringFormatterNet
    This.FormatString = RHS
End Property

Private Function IStringFormatter_Format(ParamArray Values() As Variant) As String
    Dim LocalArgs() As Variant
    'NOTE to myself: 'VF: rationale for taking lbound member as args? due use of .net stringbuilder?
    '?? where   Set Context = AppContext.Create(FormatterFactory:=New VBStringFormatterFactory)
    If IsArray(Values(LBound(Values))) Then
        LocalArgs = Values(LBound(Values))
    Else
        ReDim LocalArgs(LBound(Values) To UBound(Values))
        Dim Index As Long
        For Index = LBound(Values) To UBound(Values)
            LocalArgs(Index) = Values(Index)
        Next
    End If
    With StringBuilderNet.AppendFormat(This.FormatString, LocalArgs)
        IStringFormatter_Format = .ToString
    End With
End Function

Private Property Get IStringFormatter_FormatString() As String
    IStringFormatter_FormatString = This.FormatString
End Property
