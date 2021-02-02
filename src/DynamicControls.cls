VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Builds dynamic MSForms UI components from a binding source."
'@Folder "MVVM.Infrastructure.StringFormatting.View.Dynamic"
'@ModuleDescription "Builds dynamic MSForms UI components from a binding source."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicControlBuilder

Private Type TState
    Context As MVVM.IAppContext
    Container As MVVM.IContainerLayout
End Type

Private This As TState

Public Function Create(ByVal Context As MVVM.IAppContext, ByVal Container As IContainerLayout) As IDynamicControlBuilder
    Dim Result As DynamicControls
    Set Result = New DynamicControls
    Set Result.Context = Context
    Set Result.Container = Container
    Set Create = Result
End Function

Public Property Get Container() As MVVM.IContainerLayout
    Set Container = This.Container
End Property

Friend Property Set Container(ByVal RHS As MVVM.IContainerLayout)
    Set This.Container = RHS
End Property

Public Property Get Context() As MVVM.IAppContext
    Set Context = This.Context
End Property

Friend Property Set Context(ByVal RHS As MVVM.IAppContext)
    Set This.Context = RHS
End Property

Private Sub BindCaptionSource(ByVal Source As Variant, ByVal Target As Object, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter)
    If VarType(Source) = vbString Then
        'late-bound
        Target.Caption = Source
        
    ElseIf TypeOf Source Is MVVM.IBindingPath Then
        Dim SourcePath As IBindingPath
        Set SourcePath = Source
        This.Context.Bindings.BindPropertyPath SourcePath.Context, SourcePath.Path, Target, _
            Converter:=Converter, _
            StringFormat:=FormatString
        
    Else
        GuardClauses.GuardExpression True, TypeName(Me), "Unsupported Source type."
    End If
End Sub

Private Function CreateLabel(ByVal SourceCaption As Variant, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal ApplyMargin As Boolean = True) As MSForms.Label
    Dim Result As MSForms.Label
    Set Result = This.Container.Add(MVVM.FormsProgID.LabelProgId, ApplyMargin:=ApplyMargin)
    Result.AutoSize = True
    Result.WordWrap = False
    BindCaptionSource SourceCaption, Result, FormatString, Converter
    Set CreateLabel = Result
End Function

Private Function ErrorAdornerOrDefault(ByVal Validator As IValueValidator, ByVal ErrorAdorner As IDynamicAdorner, ByVal Target As Object) As IDynamicAdorner
    Dim Result As IDynamicAdorner
    If Not Validator Is Nothing Then
        If ErrorAdorner Is Nothing Then
            Set Result = This.Context.Validation.AdornerFactory.Create(Target)
        Else
            Set Result = ErrorAdorner
        End If
    End If
    Set ErrorAdornerOrDefault = Result
End Function

Private Function IDynamicControlBuilder_CheckBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.CheckBox
    
    Dim Result As MSForms.CheckBox
    Set Result = This.Container.Add(MVVM.FormsProgID.CheckBoxProgId)
    
    BindCaptionSource SourceCaption, Result
    This.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.Path, Result
    
    Set IDynamicControlBuilder_CheckBoxFor = Result
    
End Function

Private Function IDynamicControlBuilder_ComboBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.ComboBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim Result As MSForms.ComboBox
    Set Result = This.Container.Add(MVVM.FormsProgID.ComboBoxProgId)
        
    This.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.Path, Result, _
        StringFormat:=FormatString, _
        Converter:=Converter, _
        Validator:=Validator, _
        ValidationAdorner:=ErrorAdornerOrDefault(Validator, ErrorAdorner, Result)

    This.Context.Bindings.BindPropertyPath SourceItems.Context, SourceItems.Path, Result
    
    Set IDynamicControlBuilder_ComboBoxFor = Result
    
End Function

Private Function IDynamicControlBuilder_CommandButtonFor(ByVal Command As ICommand, ByVal BindingContext As Object, ByVal SourceCaption As Variant) As MSForms.CommandButton
    
    Dim Result As MSForms.CommandButton
    Set Result = This.Container.Add(MVVM.FormsProgID.CommandButtonProgId)
    
    Dim Button As MSForms.Control
    Set Button = Result
    
    If TypeOf Command Is CancelCommand Then
        Button.Cancel = True
    ElseIf TypeOf Command Is AcceptCommand Then
        Button.Default = True
    End If
    
    BindCaptionSource SourceCaption, Result
    This.Context.Commands.BindCommand BindingContext, Result, Command
    
    Set IDynamicControlBuilder_CommandButtonFor = Result
    
End Function

Private Function IDynamicControlBuilder_LabelFor(ByVal SourceCaption As Variant, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter) As MSForms.Label
    Set IDynamicControlBuilder_LabelFor = CreateLabel(SourceCaption, FormatString, Converter)
End Function

Private Function IDynamicControlBuilder_ListBoxFor(ByVal SourceValue As IBindingPath, ByVal SourceItems As IBindingPath, Optional ByVal TitleSource As Variant) As MSForms.ListBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim Result As MSForms.ListBox
    Set Result = This.Container.Add(MVVM.FormsProgID.ListBoxProgId)
    
    This.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.Path, Result
    This.Context.Bindings.BindPropertyPath SourceItems.Context, SourceItems.Path, Result
    
    Set IDynamicControlBuilder_ListBoxFor = Result
    
End Function

Private Function IDynamicControlBuilder_OptionButtonFor(ByVal SourceValue As IBindingPath, ByVal SourceCaption As Variant) As MSForms.OptionButton
    
    Dim Result As MSForms.OptionButton
    Set Result = This.Container.Add(MVVM.FormsProgID.OptionButtonProgId)
    
    BindCaptionSource SourceCaption, Result
    This.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.Path, Result
    
    Set IDynamicControlBuilder_OptionButtonFor = Result
    
End Function

Private Function IDynamicControlBuilder_TextAreaFor(ByVal SourceValue As IBindingPath, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim Result As MSForms.TextBox
    Set Result = IDynamicControlBuilder_TextBoxFor(SourceValue, Converter:=Converter, Validator:=Validator, ErrorAdorner:=ErrorAdorner)
    
    Result.MultiLine = True
    Result.WordWrap = True
    Result.ScrollBars = fmScrollBarsVertical
    
    Set IDynamicControlBuilder_TextAreaFor = Result
    
End Function

Private Function IDynamicControlBuilder_TextBoxFor(ByVal SourceValue As IBindingPath, Optional ByVal FormatString As String, Optional ByVal Converter As IValueConverter, Optional ByVal Validator As IValueValidator, Optional ByVal ErrorAdorner As IDynamicAdorner, Optional ByVal TitleSource As Variant) As MSForms.TextBox
    
    If Not IsEmpty(TitleSource) Then
        CreateLabel TitleSource, ApplyMargin:=False
    End If
    
    Dim Result As MSForms.TextBox
    Set Result = This.Container.Add(MVVM.FormsProgID.TextBoxProgId)
    
    Dim Trigger As BindingUpdateSourceTrigger
'    'VF: Logic error?: would always be zero and that means never // missing: if not validator is nothing then ??? trigger not picked up from validator
'    'is updated later. Why not set here and skip update later?
'    If Validator Is Nothing Then
'        Trigger = Validator.Trigger '= 0 that is never
''    '    'VF:
''    '    Else ' set value of validator
''    '        Trigger = Validator.Trigger
''    '    End If
'    End If
    'VF: simplifies to unconditional
    'but causes DecimalValuator::IValueValidator_Trigger() to be called twice
    Trigger = Validator.Trigger
    
    
    This.Context.Bindings.BindPropertyPath SourceValue.Context, SourceValue.Path, Result, _
        StringFormat:=FormatString, _
        Converter:=Converter, _
        Validator:=Validator, _
        ValidationAdorner:=ErrorAdornerOrDefault(Validator, ErrorAdorner, Result), _
        UpdateTrigger:=Trigger
    
    Set IDynamicControlBuilder_TextBoxFor = Result
    
End Function

