VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit
Implements ITestStub
Implements IBindingManager
Implements IHandlePropertyChanged

Private Type TState
    Stub As TestStubBase
    Context As IAppContext
    PropertyBindings As Collection
    NotifierFactory As INotifierFactory
    StringFormatterFactory As IStringFormatterFactory
    DebugOutput As Boolean
End Type

Private This As TState

Public Property Get NotifierFactory() As INotifierFactory
    Set NotifierFactory = This.NotifierFactory
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
    'Set This.Context = TestAppContext.Create
    Set This.NotifierFactory = New TestNotifierFactory
    Set This.PropertyBindings = New Collection
End Sub

Private Sub IBindingManager_Apply(ByVal Source As Object)
    This.Stub.OnInvoke "Apply"
End Sub

'Private Function IBindingManager_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String, _
Optional ByVal Mode As BindingMode = 0&, _
Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger = 1&, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal FormatString As String, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
'VF: changed default UpdateTrigger due to revised enum BindingUpdateSourceTrigger
Private Function IBindingManager_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String, _
Optional ByVal Mode As BindingMode = 0&, _
Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger = 2&, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal FormatString As String, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
        
    This.Stub.OnInvoke "BindPropertyPath"
    
    Dim Binding As IPropertyBinding
    Set Binding = TestPropertyBinding.Create(This.Context, BindingPath.Create(Source, PropertyPath), Target, TargetProperty, Mode, UpdateTrigger, Validator, Converter, This.StringFormatterFactory.Create(FormatString), ValidationAdorner)
    This.PropertyBindings.Add Binding
    
    Set IBindingManager_BindPropertyPath = Binding
End Function

Private Property Get IBindingManager_DebugOutput() As Boolean
    This.Stub.OnInvoke "DebugOutput.Get"
    IBindingManager_DebugOutput = This.DebugOutput
End Property

Private Property Get IBindingManager_NotifierFactory() As INotifierFactory
    This.Stub.OnInvoke "NotifierFactory.Get"
    Set IBindingManager_NotifierFactory = This.NotifierFactory
End Property

Private Property Get IBindingManager_StringFormatterFactory() As IStringFormatterFactory
    This.Stub.OnInvoke "StringFormatterFactory.Get"
    Set IBindingManager_StringFormatterFactory = This.StringFormatterFactory
End Property

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Stub.OnInvoke "OnPropertyChanged"
End Sub

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = This.Stub.MemberInvokes(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = This.Stub.ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    This.Stub.Verify Assert, MemberName, ExpectedInvokes
End Sub

