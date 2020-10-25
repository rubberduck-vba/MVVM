VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An observer that provides the infrastructure to relay MSForms.Control events."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An observer that provides the infrastructure to relay MSForms.Control events."
'@Interface
'@Exposed
Option Explicit

Public Sub RegisterHandler(ByVal Handler As IHandleControlEvents)
End Sub

Public Sub OnEnter()
End Sub

Public Sub OnExit(ByVal Cancel As MSForms.IReturnBoolean)
End Sub

Public Sub OnAfterUpdate()
End Sub

Public Sub OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
End Sub
