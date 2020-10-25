VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcceptCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A command that closes (hides) a View, enabled when the ViewModel is valid."
'@Folder MVVM.Common.Commands
'@ModuleDescription "A command that closes (hides) a View, enabled when the ViewModel is valid."
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICommand

Private Type TState
    View As IView
    ValidationManager As IValidationManager
End Type

Private This As TState

'@Description "Creates a new instance of this command."
Public Function Create(ByVal View As IView, ByVal ValidationManager As IValidationManager) As ICommand
Attribute Create.VB_Description = "Creates a new instance of this command."
    GuardClauses.GuardNonDefaultInstance Me, AcceptCommand
    GuardClauses.GuardNullReference View, TypeName(Me), "View cannot be Nothing."
    GuardClauses.GuardNullReference ValidationManager, TypeName(Me), "ValidationManager cannot be Nothing."
    Dim Result As AcceptCommand
    Set Result = New AcceptCommand
    Set Result.View = View
    Set Result.ValidationManager = ValidationManager
    Set Create = Result
End Function

Public Property Get View() As IView
    Set View = This.View
End Property

Friend Property Set View(ByVal RHS As IView)
    GuardClauses.GuardDefaultInstance Me, AcceptCommand
    GuardClauses.GuardDoubleInitialization This.View, TypeName(Me)
    Set This.View = RHS
End Property

Public Property Get ValidationManager() As IValidationManager
    Set ValidationManager = This.ValidationManager
End Property

Friend Property Set ValidationManager(ByVal RHS As IValidationManager)
    GuardClauses.GuardDefaultInstance Me, AcceptCommand
    GuardClauses.GuardDoubleInitialization This.ValidationManager, TypeName(Me)
    Set This.ValidationManager = RHS
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = This.ValidationManager.IsValid(Context)
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Accept changes and close."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    This.View.Hide
End Sub
