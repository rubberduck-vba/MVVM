VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes a View."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Describes a View."
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the ViewModel / binding Source."
Public Property Get ViewModel() As Object
Attribute ViewModel.VB_Description = "Gets the ViewModel / binding Source."
End Property

'@Description "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
Public Function ShowDialog() As Boolean
Attribute ShowDialog.VB_Description = "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
End Function

'@Description "Displays the View."
Public Sub Show()
Attribute Show.VB_Description = "Displays the View."
End Sub

'@Description "Hides the View."
Public Sub Hide()
Attribute Hide.VB_Description = "Hides the View."
End Sub
