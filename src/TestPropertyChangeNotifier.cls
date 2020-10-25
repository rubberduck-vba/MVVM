VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyChangeNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Stubs")
Option Explicit
Implements INotifyPropertyChanged
Implements ITestStub

Private Type TState
    Handlers As Collection
    Stub As TestStubBase
End Type
Private This As TState

Public Property Get Handlers() As Collection
    Set Handlers = This.Handlers
End Property

Public Property Get Stub() As ITestStub
    Set Stub = This.Stub
End Property

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
    Set This.Stub = New TestStubBase
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Stub.OnInvoke "OnPropertyChanged"
    Dim Handler As IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Stub.OnInvoke "RegisterHandler"
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

