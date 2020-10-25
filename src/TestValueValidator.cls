VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder Tests.Stubs
Option Explicit
Implements IValueValidator
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    IsValid As Boolean
    Message As String
End Type

Private This As TState

'@Description "Gets/sets a value that determines whether this validator evaluates to True or False."
Public Property Get IsValid() As Boolean
Attribute IsValid.VB_Description = "Gets/sets a value that determines whether this validator evaluates to True or False."
    IsValid = This.IsValid
End Property

Public Property Let IsValid(ByVal RHS As Boolean)
    This.IsValid = RHS
End Property

'@Description "Gets/sets the validation error message for this validator."
Public Property Get Message() As String
Attribute Message.VB_Description = "Gets/sets the validation error message for this validator."
    Message = This.Message
End Property

Public Property Let Message(ByVal RHS As String)
    This.Message = RHS
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    This.Stub.OnInvoke "IsValid"
    IValueValidator_IsValid = This.IsValid
End Function

Private Property Get IValueValidator_Message() As String
    This.Stub.OnInvoke "Message"
    IValueValidator_Message = This.Message
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

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
