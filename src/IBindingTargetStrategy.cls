VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingTargetStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A strategy for creating an IPropertyBinding "
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "A strategy for creating an IPropertyBinding "
'@Interface
Option Explicit

'@Description "Infers the target property path from the runtime type of the specified Target."
Public Function DefaultPropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute DefaultPropertyBindingFor.VB_Description = "Infers the target property path from the runtime type of the specified Target."
End Function

'@Description "Creates a property binding as specified."
Public Function PropertyBindingFor(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As IBindingPath, _
ByVal Mode As BindingMode, _
ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
ByVal Validator As IValueValidator, _
ByVal Converter As IValueConverter, _
ByVal StringFormat As IStringFormatter, _
ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute PropertyBindingFor.VB_Description = "Creates a property binding as specified."
End Function
