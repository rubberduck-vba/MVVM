VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBindingObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub a binding source or target for unit tests."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub a binding source or target for unit tests."
'@PredeclaredId
Option Explicit
Implements ITestStub
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged ' intercept events handled in PropertyBindingBase

Private Type TState
    Stub As TestStubBase
    Handlers As Collection
    TestStringProperty As String
    TestNumericProperty As Long
    TestBindingObjectProperty As TestBindingObject
End Type

Private This As TState

Public Function Create() As TestBindingObject
    GuardClauses.GuardNonDefaultInstance Me, TestBindingObject, TypeName(Me)
    Dim Result As TestBindingObject
    Set Result = New TestBindingObject
    Set Create = Result
End Function

Public Property Get TestStringProperty() As String
    This.Stub.OnInvoke "TestStringProperty.Get"
    TestStringProperty = This.TestStringProperty
End Property

Public Property Let TestStringProperty(ByVal RHS As String)
    This.Stub.OnInvoke "TestStringProperty.Let"
    If This.TestStringProperty <> RHS Then
        This.TestStringProperty = RHS
        OnPropertyChanged Me, "TestStringProperty"
    End If
End Property

Public Property Get TestNumericProperty() As Long
    This.Stub.OnInvoke "TestNumericProperty.Get"
    TestNumericProperty = This.TestNumericProperty
End Property

Public Property Let TestNumericProperty(ByVal RHS As Long)
    This.Stub.OnInvoke "TestNumericProperty.Let"
    If This.TestNumericProperty <> RHS Then
        This.TestNumericProperty = RHS
        OnPropertyChanged Me, "TestNumericProperty"
    End If
End Property

Public Property Get TestBindingObjectProperty() As TestBindingObject
    This.Stub.OnInvoke "TestBindingObjectProperty.Get"
    Set TestBindingObjectProperty = This.TestBindingObjectProperty
End Property

Public Property Set TestBindingObjectProperty(ByVal RHS As TestBindingObject)
    This.Stub.OnInvoke "TestBindingObjectProperty.Set"
    If Not This.TestBindingObjectProperty Is RHS Then
        Set This.TestBindingObjectProperty = RHS
        OnPropertyChanged Me, "TestBindingObjectProperty"
    End If
End Property

Private Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
    Set This.Handlers = New Collection
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Handlers.Add Handler
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
