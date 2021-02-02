VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatterNetFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A factory that creates new .NET-syntax string formatters."
'@Folder MVVM.Infrastructure.StringFormatting
'@ModuleDescription "A factory that creates new .NET-syntax string formatters."
Option Explicit
Implements IStringFormatterFactory

Private Function IStringFormatterFactory_Create(ByVal FormatString As String) As IStringFormatter
    Set IStringFormatterFactory_Create = StringFormatterNet.Create(FormatString)
End Function
