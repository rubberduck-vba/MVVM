VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequiredStringValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A validator that requires a string value to be non-empty."
'@Folder MVVM.Common.Validators
'@ModuleDescription "A validator that requires a string value to be non-empty."
Option Explicit
Implements IValueValidator

'VF: failed space bar/blank spaces between words entered in TextBox that's why TextBox event (IValueValidator_Trigger) switched to onChange from onKeyPress
'but OnChange was also needed to capture Cut and Paste in TextBox
Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    IValueValidator_IsValid = Len(Trim$(Value)) > 0
End Function

Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "Value cannot be empty."
End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    'VF: must use OnChange (tried OnKeyPres first) despite sucking some performance
    'IValueValidator_Trigger = OnExit
    'IValueValidator_Trigger = BindingUpdateSourceTrigger.OnKeyPress 'does not catch BackSpace and Delete so empty textbox is skipped
    'tried blindly first:
    'IValueValidator_Trigger = BindingUpdateSourceTrigger.OnPropertyChanged 'does not catch BackSpace and Delete either
    'catches cut/paste and backspace/delete
    IValueValidator_Trigger = BindingUpdateSourceTrigger.OnChange
    
    'OnExit and OnPropertyChanged are still somewhat ambiguously defined, _
    but OnExit gets to keep the target control focused if there’s a validation error, _
    and OnPropertyChanged is [currently] firing for every keypress in a TextBox, after the control’s Text property udpates. _
    Disambiguating the terms would mean breaking with MSForms’ meaning of “updated”… which may actually be a good thing: _
    OnPropertyChanged would basically fire on exit but without a possible navigation-blocking validation error, and then _
    OnKeyPress mode would still need to behave like OnPropertyChanged as far as updating the source goes.
    
    'VF: description on webpage does not match code
    'What actually triggers the propagation of the new TextBox.Text value to the ViewModel property, is the _
    BeforeUpdate and Exit events, both of which are initially handled in the ControlEventsPunk class, an object that leverages the Win32 IConnectionPoint API to hook up event handlers for the MSForms.Controls interface of our TextBox control
    
End Property
