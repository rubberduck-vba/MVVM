VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object with the ability to format a string value."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Represents an object with the ability to format a string value."
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the format string used when formatting values."
Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Gets the format string used when formatting values."
End Property

'@Description "Formats the specified values as per the FormatString."
Public Function Format(ParamArray Values() As Variant) As String
Attribute Format.VB_Description = "Formats the specified values as per the FormatString."
End Function
