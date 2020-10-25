# Model-View-ViewModel Infrastructure for VBA

### Quick Start

0. Not *required*, but consider installing [Rubberduck](https://rubberduckvba.com). You'll get to organize these dozens of modules into a clean folder hierarchy, enhanced navigation tooling, unit tests, and so much more.

1. Import all modules into your VBA project (current version requires a reference to the Microsoft Scripting Runtime library).

2. Add a `UserForm` to your project, get to its code-behind and add this boilerplate:

```vb
'@Folder MVVM.Example
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As MVVM.IAppContext
    ViewModel As ExampleViewModel
    IsCancelled As Boolean
End Type

Private This As TState

'@Description "Creates a new instance of this form."
Public Function Create(ByVal Context As MVVM.IAppContext, ByVal ViewModel As ExampleViewModel) As IView
    Dim Result As ExampleDynamicView
    Set Result = New ExampleDynamicView
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel
    Set Create = Result
End Function

Public Property Get Context() As MVVM.IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As MVVM.IAppContext)
    Set This.Context = RHS
End Property

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set This.ViewModel = RHS
End Property

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()
    
    Dim Layout As IContainerLayout
    Set Layout = ContainerLayout.Create(Me.Controls, TopToBottom)
    
    With DynamicControls.Create(This.Context, Layout)
        'TODO create controls for ViewModel properties
    End With
    
    This.Context.Bindings.Apply This.ViewModel
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeView
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    Me.Show vbModal
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
```
