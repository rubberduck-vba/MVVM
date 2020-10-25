VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringFormatterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A factory that creates instances of a particular IStringFormatter implementation."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "A factory that creates instances of a particular IStringFormatter implementation."
'@Interface
'@Exposed
Option Explicit

'@Description "Creates a new instance of a string formatter."
Public Function Create(ByVal FormatString As String) As IStringFormatter
Attribute Create.VB_Description = "Creates a new instance of a string formatter."
End Function
