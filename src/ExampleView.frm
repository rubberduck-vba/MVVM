VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleView 
   Caption         =   "ExampleView"
   ClientHeight    =   4716
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6564
   OleObjectBlob   =   "ExampleView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An example implementation of a View."

'@Folder MVVM.Example
'@ModuleDescription "An example implementation of a View."
Implements IView
Implements ICancellable
Option Explicit

Private Type TView
    Context As MVVM.IAppContext
    
    'IView state:
    ViewModel As ExampleViewModel
    
    'ICancellable state:
    IsCancelled As Boolean
    
End Type

Private This As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As ExampleViewModel) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, ExampleView, TypeName(Me)
    GuardClauses.GuardNullReference ViewModel, TypeName(Me)
    GuardClauses.GuardNullReference Context, TypeName(Me)
    
    Dim Result As ExampleView
    Set Result = New ExampleView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel
    
    Set Create = Result
    
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ExampleView
End Property

'@Description "Gets/sets the ViewModel to use as a context for property and command bindings."
Public Property Get ViewModel() As ExampleViewModel
Attribute ViewModel.VB_Description = "Gets/sets the ViewModel to use as a context for property and command bindings."
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As ExampleViewModel)
    GuardClauses.GuardDefaultInstance Me, ExampleView, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.ViewModel = RHS
End Property

'@Description "Gets/sets the MVVM application context."
Public Property Get Context() As MVVM.IAppContext
Attribute Context.VB_Description = "Gets/sets the MVVM application context."
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As MVVM.IAppContext)
    GuardClauses.GuardDefaultInstance Me, ExampleView, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Context, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Context = RHS
End Property

Private Sub BindViewModelCommands()
    With Context.Commands
        .BindCommand ViewModel, Me.OkButton, AcceptCommand.Create(Me, This.Context.Validation)
        .BindCommand ViewModel, Me.CancelButton, CancelCommand.Create(Me)
        .BindCommand ViewModel, Me.BrowseButton, ViewModel.SomeCommand
        '...
    End With
End Sub

Private Sub BindViewModelProperties()
    With Context.Bindings
        
        'Binding to a Label control without a target property like this creates a CaptionPropertyBinding.
        'This type of binding defaults to a one-way binding (source property -> target property)
        'that will update the target caption if the source changes (source property setter must invoke OnPropertyChanged):
        .BindPropertyPath ViewModel, "Instructions", Me.InstructionsLabel
        
        'If we know we're not going to change the caption at any point,
        'we can always make the binding one-time (source property -> target property).
        'Binding to an OptionButton control without specifying a target property creates an OptionButtonBinding,
        'and because we're binding to a String property the target property is inferred to be the Caption:
        .BindPropertyPath ViewModel, "SomeOptionName", Me.OptionButton1, Mode:=OneTimeBinding
        .BindPropertyPath ViewModel, "SomeOtherOptionName", Me.OptionButton2, Mode:=OneTimeBinding
        
        'Binding to a TextBox control we can specify a format string to format the control when
        'it loses focus, and by setting the binding's UpdateTrigger to OnKeyPress we get to
        'use a KeyValidator that can prevent invalid (here, non-numeric) inputs.
        'Without specifying a target property, we're binding to TextBox.Text regardless of the data type of the source property:
        .BindPropertyPath ViewModel, "SomeAmount", Me.AmountBox, _
            StringFormat:="{0:C2}", _
            UpdateTrigger:=OnKeyPress, _
            Validator:=New DecimalKeyValidator
        
        'Binding a Date property on the ViewModel to a TextBox control works best with a value converter.
        'The converter must be able to convert the specified format string into a Date, and back.
        'We can handle validation errors by providing a ValidationErrorAdorner instance:
        .BindPropertyPath ViewModel, "SomeDate", Me.TextBox1, _
            StringFormat:="{0:MMMM dd, yyyy}", _
            Converter:=StringToDateConverter.Default, _
            Validator:=New RequiredStringValidator, _
            ValidationAdorner:=ValidationErrorAdorner.Create( _
                Target:=Me.TextBox1, _
                TargetFormatter:=ValidationErrorFormatter.WithErrorBorderColor.WithErrorBackgroundColor)
        
        'OptionButton controls automatically bind their value to a Boolean property:
        .BindPropertyPath ViewModel, "SomeOption", Me.OptionButton1
        .BindPropertyPath ViewModel, "SomeOtherOption", Me.OptionButton2
        
        'When binding an array property to a ComboBox target, the List property is the implicit target:
        .BindPropertyPath ViewModel, "SomeItems", Me.ComboBox1
        
        'If we want we can bind a String property to automatically bind to ComboBox.Text:
        .BindPropertyPath ViewModel, "SelectedItemText", Me.ComboBox1
        
        'Or (and?) we want we can bind a Long property to automatically bind to ComboBox.ListIndex:
        .BindPropertyPath ViewModel, "SelectedItemIndex", Me.ComboBox1
        
        'Binding to any other source property data type binds to ComboBox.Value;
        'that's especially useful when the List has multiple columns and the first (the Value!) contains some hidden unique ID.
        
        'So MVVM works for a MSForms UI.
        'What if the binding target was something else?
        
        'a worksheet cell's value?
        '.BindPropertyPath ViewModel, "SelectedItemText", Sheet1.Range("A1"), "Value"
        
        'a...chart's title?
        '.BindPropertyPath ViewModel, "Instructions", Sheet1.ChartObjects("Chart 1"), "Chart.ChartTitle.Text"
        
        '...I've created a monster, haven't I?
        
    End With
End Sub

Private Sub InitializeBindings()
    If ViewModel Is Nothing Then Exit Sub
    BindViewModelProperties
    BindViewModelCommands
    This.Context.Bindings.Apply ViewModel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
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
    InitializeBindings
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeBindings
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
