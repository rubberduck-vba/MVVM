VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IContainerLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates the layout logic for dynamic UI components."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Encapsulates the layout logic for dynamic UI components."
'@Interface
'@Exposed
Option Explicit

'@Description "Adds a new dynamic control to the layout."
Public Function Add(ByVal ProgID As String, Optional ByVal Height As Variant, Optional ByVal Width As Variant, Optional ByVal ApplyMargin As Boolean = True) As Object
Attribute Add.VB_Description = "Adds a new dynamic control to the layout."
End Function

'@Description "Moves the position of the next control by a margin."
Public Sub SkipPosition(Optional ByVal Margin As Long)
Attribute SkipPosition.VB_Description = "Moves the position of the next control by a margin."
End Sub
