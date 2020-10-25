VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object that can decorate a binding target with implementation-defined components."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Describes an object that can decorate a binding target with implementation-defined components."
'@Interface
'@Exposed
Option Explicit

'@Description "Brings the adorner components into view."
Public Sub Show(ByVal Message As String)
Attribute Show.VB_Description = "Brings the adorner components into view."
End Sub

'@Description "Hides the adorner components from view."
Public Sub Hide()
Attribute Hide.VB_Description = "Hides the adorner components from view."
End Sub
