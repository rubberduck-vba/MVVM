VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing property bindings in an application context."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Represents an object responsible for managing property bindings in an application context."
'@Interface
'@Exposed
Option Explicit

'@Description "Creates a new property binding for the specified path, source, and target."
Public Function BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As String, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute BindPropertyPath.VB_Description = "Creates a new property binding for the specified path, source, and target."
End Function

'@Description "Gets a reference to an object that creates notification helper object instances."
Public Property Get NotifierFactory() As INotifierFactory
Attribute NotifierFactory.VB_Description = "Gets a reference to an object that creates notification helper object instances."
End Property

'@Description "Gets a reference to an object that creates string-formatting helper object instances."
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets a reference to an object that creates string-formatting helper object instances."
End Property

'@Description "Gets a value indicating whether bindings produce debugger output."
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets a value indicating whether bindings produce debugger output."
End Property

'@Description "Applies all property bindings."
Public Sub Apply(ByVal Context As Object)
Attribute Apply.VB_Description = "Applies all property bindings."
End Sub
