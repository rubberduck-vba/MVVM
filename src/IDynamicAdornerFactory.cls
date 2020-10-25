VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicAdornerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An abstract factory that creates dynamic validation error adorners."
'@Interface
Option Explicit

'@Description "Creates a new dynamic validation error adorner for the specified target."
Public Function Create(ByVal Target As Object) As IDynamicAdorner
End Function
