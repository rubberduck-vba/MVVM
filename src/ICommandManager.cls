VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing command bindings in an application context."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Represents an object responsible for managing command bindings in an application context."
'@Interface
'@Exposed
Option Explicit

'@Description "Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control)."
Public Function BindCommand(ByVal BindingContext As Object, ByVal Target As Object, ByVal Command As ICommand) As ICommandBinding
Attribute BindCommand.VB_Description = "Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control)."
End Function

'@Description "Causes all commands to evaluate whether they can be executed."
Public Sub EvaluateCanExecute(ByVal BindingContext As Object)
Attribute EvaluateCanExecute.VB_Description = "Causes all commands to evaluate whether they can be executed."
End Sub
