VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToDateConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A value converter that can convert a String into a Date, or a Date into a String."
'@Folder MVVM.Common.Converters
'@ModuleDescription "A value converter that can convert a String into a Date, or a Date into a String."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValueConverter

Public Function Default() As IValueConverter
    GuardClauses.GuardNonDefaultInstance Me, StringToDateConverter
    Set Default = StringToDateConverter
End Function

Private Function IValueConverter_Convert(ByVal Value As Variant) As Variant
    IValueConverter_Convert = Convert(Value)
End Function

Private Function IValueConverter_ConvertBack(ByVal Value As Variant) As Variant
    IValueConverter_ConvertBack = Convert(Value)
End Function

Private Function Convert(ByVal Value As Variant) As Variant
    GuardClauses.GuardExpression Not IsDate(Value), TypeName(Me), "Value '" & Value & "' was not recognized as a valid date."
    If VarType(Value) = vbDate Then
        Convert = CStr(Value)
    Else
        Convert = CDate(Value)
    End If
End Function
