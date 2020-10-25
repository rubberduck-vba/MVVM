VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an IPropertyBinding dependency in a unit test."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub an IPropertyBinding dependency in a unit test."
'@PredeclaredId
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
Optional ByVal TargetProperty As String = "Value", _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnExit, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As ITestStub
    Dim Result As TestPropertyBinding
    Set Result = New TestPropertyBinding
    
    Dim Base As PropertyBindingBase
    Set Base = PropertyBindingBase.Create(Context, Source, Target, TargetProperty, Mode, UpdateSource, Converter, StringFormat, Validator, ValidationAdorner)
    Result.InjectBindingInfo Base
    Set Create = Result
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, TestPropertyBinding, TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Handler, TypeName(Me)
    Set This.Base = BindingInfo
    Set This.Handler = BindingInfo
End Sub

Public Property Get AsIHandlePropertyChanged() As IHandlePropertyChanged
    Set AsIHandlePropertyChanged = Me
End Property

Public Property Get Source() As IBindingPath
    Set Source = This.Base.Source
End Property

Public Property Set Source(ByVal RHS As IBindingPath)
    Set This.Base.Source = RHS
End Property

Public Property Get Target() As IBindingPath
    Set Target = This.Base.Target
End Property

Public Property Set Target(ByVal RHS As IBindingPath)
    Set This.Base.Target = RHS
End Property

Public Property Get Mode() As BindingMode
    Mode = This.Base.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    This.Base.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    This.Base.UpdateSourceTrigger = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = This.Base.Validator
End Property

Public Property Set Validator(ByVal RHS As IValueValidator)
    Set This.Base.Validator = RHS
End Property

Public Property Get ValidationAdorner() As IDynamicAdorner
    Set ValidationAdorner = This.Base.ValidationAdorner
End Property

Public Property Set ValidationAdorner(ByVal RHS As IDynamicAdorner)
    Set This.Base.ValidationAdorner = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = This.Base.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    Set This.Base.Converter = RHS
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Stub.OnInvoke "HandlePropertyChanged"
    This.Handler.HandlePropertyChanged Source, PropertyName
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    This.Stub.OnInvoke "DefaultTargetProperty.Get"
    IPropertyBinding_DefaultTargetProperty = vbNullString
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    This.Stub.OnInvoke "StringFormat.Get"
    Set IPropertyBinding_StringFormat = This.Base.StringFormat
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    This.Stub.OnInvoke "Validator.Get"
    Set IPropertyBinding_Validator = This.Base.Validator
End Property

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = This.Stub.MemberInvokes(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = This.Stub.ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    This.Stub.Verify Assert, MemberName, ExpectedInvokes
End Sub

Private Property Get IPropertyBinding_Mode() As BindingMode
    This.Stub.OnInvoke "Mode.Get"
    IPropertyBinding_Mode = This.Base.Mode
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    This.Stub.OnInvoke "UpdateSourceTrigger.Get"
    IPropertyBinding_UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    This.Stub.OnInvoke "Source.Get"
    Set IPropertyBinding_Source = This.Base.Source
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    This.Stub.OnInvoke "Target.Get"
    Set IPropertyBinding_Target = This.Base.Target
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    This.Stub.OnInvoke "Converter.Get"
    Set IPropertyBinding_Converter = This.Base.Converter
End Property

Private Sub IPropertyBinding_Apply()
    This.Stub.OnInvoke "Apply"
End Sub

