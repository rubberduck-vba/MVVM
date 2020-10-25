VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicControlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Builds dynamic MSForms UI components from a binding source."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Builds dynamic MSForms UI components from a binding source."
'@Interface
'@Exposed
Option Explicit

'@Description "Creates a multiline MSForms.TextBox control for the spercified String property binding path."
Public Function TextAreaFor(ByVal SourceValue As IBindingPath, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
Attribute TextAreaFor.VB_Description = "Creates a multiline MSForms.TextBox control for the spercified String property binding path."
End Function

'@Description "Creates a MSForms.TextBox control for the specified String property binding path."
Public Function TextBoxFor(ByVal SourceValue As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
Attribute TextBoxFor.VB_Description = "Creates a MSForms.TextBox control for the specified String property binding path."
End Function

'@Description "Creates a MSForms.Label control for the specified Caption string or String property binding path."
Public Function LabelFor(ByVal SourceCaption As Variant, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter) As MSForms.Label
Attribute LabelFor.VB_Description = "Creates a MSForms.Label control for the specified Caption string or String property binding path."
End Function

'@Description "Creates a MSForms.ComboBox control for the specified Value property binding path; SourceItems should be an array property."
Public Function ComboBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.ComboBox
Attribute ComboBoxFor.VB_Description = "Creates a MSForms.ComboBox control for the specified Value property binding path; SourceItems should be an array property."
End Function

'@Description "Creates a MSForms.ListBox control for the specified Value property binding path; SourceItems should be an array property."
Public Function ListBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal TitleSource As Variant) As MSForms.ListBox
Attribute ListBoxFor.VB_Description = "Creates a MSForms.ListBox control for the specified Value property binding path; SourceItems should be an array property."
End Function

'@Description "Creates a MSForms.OptionButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
Public Function OptionButtonFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.OptionButton
Attribute OptionButtonFor.VB_Description = "Creates a MSForms.OptionButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
End Function

'@Description "Creates a MSForms.CheckBoxButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
Public Function CheckBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.CheckBox
Attribute CheckBoxFor.VB_Description = "Creates a MSForms.CheckBoxButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
End Function

'@Description "Creates a MSForms.CommandButton control for the specified ICommand, using the specified ViewModel context and Caption string or String property binding path."
Public Function CommandButtonFor(ByVal Command As ICommand, ByVal BindingContext As Object, ByVal SourceCaption As Variant) As MSForms.CommandButton
Attribute CommandButtonFor.VB_Description = "Creates a MSForms.CommandButton control for the specified ICommand, using the specified ViewModel context and Caption string or String property binding path."
End Function
