VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object responsible for binding a command to a UI element."
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the event source object bound to the command."
Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets the event source object bound to the command."
End Property

'@Description "Gets the command bound to the event source."
Public Property Get Command() As ICommand
Attribute Command.VB_Description = "Gets the command bound to the event source."
End Property

'@Description "Evaluates whether the command can execute given the binding context."
Public Sub EvaluateCanExecute(ByVal Context As Object)
Attribute EvaluateCanExecute.VB_Description = "Evaluates whether the command can execute given the binding context."
End Sub
