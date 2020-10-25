VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object that describes how a control changes its appearance given a validation error."
'@Interface
'@Exposed
Option Explicit

'@Description "Applies error styling to the target given the specified validation error message."
Public Sub Apply(ByVal Target As Object, ByVal Message As String)
Attribute Apply.VB_Description = "Applies error styling to the target given the specified validation error message."
End Sub

'@Description "Removes error styling on the target."
Public Sub Restore(ByVal Target As Object)
Attribute Restore.VB_Description = "Removes error styling on the target."
End Sub
