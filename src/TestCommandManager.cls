VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit
Implements ICommandManager
Implements ITestStub

Private Type TState
    Stub As TestStubBase
End Type

Private This As TState

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Function ICommandManager_BindCommand(ByVal BindingContext As Object, ByVal Target As Object, ByVal Command As ICommand) As ICommandBinding
    This.Stub.OnInvoke "BindCommand"
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal BindingContext As Object)
    This.Stub.OnInvoke "EvaluateCanExecute"
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


