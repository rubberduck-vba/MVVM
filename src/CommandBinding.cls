VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
'@Folder MVVM.Infrastructure.Bindings.CommandBindings
'@ModuleDescription "An object responsible for binding a command to a UI element."
'@PredeclaredId
'@Exposed
Implements ICommandBinding
Implements IDisposable
Option Explicit

Private Type TCommandBinding
    ViewModel As Object
    Target As Object
    Command As ICommand
End Type

Private WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
Private WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
Private WithEvents ImageEvents As MSForms.Image
Attribute ImageEvents.VB_VarHelpID = -1
Private WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private This As TCommandBinding

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
Public Function ForCommandButton(ByVal Target As MSForms.CommandButton, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCommandButton.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCommandButton = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
Public Function ForCheckBox(ByVal Target As MSForms.CheckBox, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCheckBox.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCheckBox = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
Public Function ForImage(ByVal Target As MSForms.Image, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForImage.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForImage = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
Public Function ForLabel(ByVal Target As MSForms.Label, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForLabel.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForLabel = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
Public Function Create(ByVal Target As Object, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute Create.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    
    Dim Result As CommandBinding
    Set Result = New CommandBinding
    
    Set Result.ViewModel = ViewModel
    Set Result.Target = Target
    Set Result.Command = Command
    
    Set Create = Result
    
End Function

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set This.ViewModel = RHS
End Property

Public Property Get Target() As Object
    Set Target = This.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    
    GuardClauses.GuardDoubleInitialization This.Target, TypeName(Me)
    Set This.Target = RHS
    
    Select Case True
        Case TypeOf RHS Is MSForms.CommandButton
            Set CommandButtonEvents = RHS
            
        Case TypeOf RHS Is MSForms.CheckBox
            Set CheckBoxEvents = RHS
            
        Case TypeOf RHS Is MSForms.Image
            Set ImageEvents = RHS
            
        Case TypeOf RHS Is MSForms.Label
            Set LabelEvents = RHS
            
        Case Else
            GuardClauses.GuardExpression _
                Throw:=True, _
                Source:=TypeName(Me), _
                Message:="Type '" & TypeName(RHS) & "' does not support command bindings at the moment."
    End Select
    
End Property

Public Property Get Command() As ICommand
    Set Command = This.Command
End Property

Public Property Set Command(ByVal RHS As ICommand)
    GuardClauses.GuardDoubleInitialization This.Command, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set This.Command = RHS
End Property

Private Property Get AsInterface() As ICommandBinding
    Set AsInterface = Me
End Property

Private Sub OnExecute()
    If Not This.Command Is Nothing Then
        This.Command.Execute This.ViewModel
    Else
        Debug.Print "BUG in " & TypeName(Me) & ": Command is 'Nothing', cannot execute."
        Debug.Assert False ' should not happen, break here if it does.
    End If
End Sub

Private Sub EvaluateCanExecute(ByVal Source As Object)
    If This.Target Is Nothing Then Exit Sub
    If This.Command Is Nothing Then
        This.Target.Enabled = False
    Else
        On Error Resume Next
        This.Target.Enabled = This.Command.CanExecute(Source)
        'a validation error formatter might have taken over the tooltip?
        This.Target.ControlTipText = IIf(This.Target.ControlTipText = vbNullString, This.Command.Description, This.Target.ControlTipText)
        On Error GoTo 0
    End If
End Sub

Public Function ToString() As String
    ToString = TypeName(This.Target) & " -> " & TypeName(This.Command)
End Function

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = This.Target
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = This.Command
End Property

Private Sub ICommandBinding_EvaluateCanExecute(ByVal Context As Object)
    EvaluateCanExecute Context
End Sub

Private Sub CheckBoxEvents_Click()
    GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.CheckBox, TypeName(Me)
    OnExecute
End Sub

Private Sub CommandButtonEvents_Click()
    GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.CommandButton, TypeName(Me)
    OnExecute
End Sub

Private Sub IDisposable_Dispose()
    Set This.Command = Nothing
    Set This.Target = Nothing
    Set This.ViewModel = Nothing
End Sub

Private Sub ImageEvents_Click()
    GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.Image, TypeName(Me)
    OnExecute
End Sub

Private Sub LabelEvents_Click()
    GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.Label, TypeName(Me)
    OnExecute
End Sub
