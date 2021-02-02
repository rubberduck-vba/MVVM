VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object responsible for binding a ViewModel property path to a UI element."
'@Interface
'@Exposed
Option Explicit

Public Enum BindingMode
    TwoWayBinding
    OneWayBinding
    OneWayToSource
    OneTimeBinding
End Enum

Public Enum BindingUpdateSourceTrigger
    'VF: added NotSetYet  and renumbered enums
    NotSetYet = 0 ''avoid being called twice, for runtime optimization
    Never = 1
    OnPropertyChanged = 2 '
    OnKeyPress = 3 '
    OnExit = 4 '
    
    'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
    OnChange = 5
    
End Enum

'@Description "Gets a value indicating the binding mode/direction."
Public Property Get Mode() As BindingMode
Attribute Mode.VB_Description = "Gets a value indicating the binding mode/direction."
End Property

'@Description "Gets a value indicating the binding update trigger."
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
Attribute UpdateSourceTrigger.VB_Description = "Gets a value indicating the binding update trigger."
End Property

'@Description "Gets the binding path resolving to the binding source."
Public Property Get Source() As IBindingPath
Attribute Source.VB_Description = "Gets the binding path resolving to the binding source."
End Property

'@Description "Gets the binding path resolving to the binding target."
Public Property Get Target() As IBindingPath
Attribute Target.VB_Description = "Gets the binding path resolving to the binding target."
End Property

'@Description "Gets the name of the target property implicitly bound for the Target class type."
Public Property Get DefaultTargetProperty() As String
Attribute DefaultTargetProperty.VB_Description = "Gets the name of the target property implicitly bound for the Target class type."
End Property

'@Description "Gets the converter (if any) used when applying the binding."
Public Property Get Converter() As IValueConverter
Attribute Converter.VB_Description = "Gets the converter (if any) used when applying the binding."
End Property

'@Description "Gets the value validator (if any) used when applying the binding."
Public Property Get Validator() As IValueValidator
Attribute Validator.VB_Description = "Gets the value validator (if any) used when applying the binding."
End Property

'@Description "Gets the string formatter (if any) used when applying the binding to the target."
Public Property Get StringFormat() As IStringFormatter
Attribute StringFormat.VB_Description = "Gets the string formatter (if any) used when applying the binding to the target."
End Property

'@Description "Determines whether the binding prevents focus leaving the control with invalid data."
Public Property Get CancelExitOnValidationError() As Boolean
Attribute CancelExitOnValidationError.VB_Description = "Determines whether the binding prevents focus leaving the control with invalid data."
End Property

'@Description "Applies the binding."
Public Sub Apply()
Attribute Apply.VB_Description = "Applies the binding."
End Sub
