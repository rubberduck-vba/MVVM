VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing command bindings."
'@Folder MVVM.Infrastructure.Bindings.CommandBindings
'@ModuleDescription "An object responsible for managing command bindings."
'@Exposed
Option Explicit
Implements ICommandManager
Implements IDisposable

Private Type TState
    CommandBindings As Collection
    DebugOutput As Boolean
End Type

Private This As TState

'@Description "Gets/sets a value indicating whether the object produces debugger output."
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value indicating whether the object produces debugger output."
    DebugOutput = This.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    This.DebugOutput = RHS
End Property

'@Description "Releases all held ICommandBinding references, prepares the object for proper destruction."
Public Sub Dispose()
Attribute Dispose.VB_Description = "Releases all held ICommandBinding references, prepares the object for proper destruction."
    '@Ignore VariableNotUsed
    Dim Index As Long
    For Index = 1 To This.CommandBindings.Count
        Disposable.TryDispose This.CommandBindings.Item(This.CommandBindings.Count)
        This.CommandBindings.Remove This.CommandBindings.Count
    Next
    Set This.CommandBindings = Nothing
End Sub

Private Sub EvaluateCanExecute(ByVal BindingContext As Object)
    Dim Binding As ICommandBinding
    For Each Binding In This.CommandBindings
        Binding.EvaluateCanExecute BindingContext
    Next
End Sub

Private Sub DebugMessage(ByVal Message As String)
    If This.DebugOutput Then Debug.Print TypeName(Me) & ": " & Message
End Sub

Private Sub Class_Initialize()
    Set This.CommandBindings = New Collection
End Sub

Private Function ICommandManager_BindCommand(ByVal BindingContext As Object, ByVal Target As Object, ByVal Command As ICommand) As ICommandBinding
    GuardClauses.GuardNullReference BindingContext, TypeName(Me)
    GuardClauses.GuardNullReference Target, TypeName(Me)
    GuardClauses.GuardNullReference Command, TypeName(Me)
    'GuardClauses.GuardDefaultInstance Me, CommandManager
    
    Dim Binding As ICommandBinding
    
    Select Case True
    
        Case TypeOf Target Is MSForms.CommandButton
            Set Binding = CommandBinding.ForCommandButton(Target, Command, BindingContext)
            
        Case TypeOf Target Is MSForms.Image
            Set Binding = CommandBinding.ForImage(Target, Command, BindingContext)
            
        Case TypeOf Target Is MSForms.Label
            Set Binding = CommandBinding.ForLabel(Target, Command, BindingContext)
            
        Case TypeOf Target Is MSForms.CheckBox
            Set Binding = CommandBinding.ForCheckBox(Target, Command, BindingContext)
            
        Case Else
            GuardClauses.GuardExpression True, TypeName(BindingContext), "Target type '" & TypeName(Target) & "' does not currently support command bindings."
            
    End Select
    
    This.CommandBindings.Add Binding
    Set ICommandManager_BindCommand = Binding
    
    DebugMessage "Created CommandBinding for " & TypeName(Command) & "."
    
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal BindingContext As Object)
    EvaluateCanExecute BindingContext
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

