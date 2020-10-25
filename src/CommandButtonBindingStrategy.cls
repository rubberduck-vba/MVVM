VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandButtonBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A target binding strategy that creates a CaptionPropertyBinding by default, or a OneWayPropertyBinding given a target property path."
'@Folder MVVM.Infrastructure.Bindings.PropertyBindings.Strategies
'@ModuleDescription "A target binding strategy that creates a CaptionPropertyBinding by default, or a OneWayPropertyBinding given a target property path."
Option Explicit
Implements IBindingTargetStrategy

Private Function IBindingTargetStrategy_DefaultPropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Set IBindingTargetStrategy_DefaultPropertyBindingFor = CaptionPropertyBinding _
        .Create(Context, Source, Target, _
            Validator:=Validator, _
            StringFormat:=StringFormat, _
            Converter:=Converter, _
            ValidationAdorner:=ValidationAdorner)

End Function

Private Function IBindingTargetStrategy_PropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As IBindingPath, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_PropertyBindingFor = OneWayPropertyBinding _
        .Create(Context, Source, Target, _
            Validator:=Validator, _
            StringFormat:=StringFormat, _
            Converter:=Converter, _
            ValidationAdorner:=ValidationAdorner)

End Function
