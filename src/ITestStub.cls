VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that stubs an interface for testing purposes."
'@Exposed
'@Folder Tests.Stubs
'@ModuleDescription "An object that stubs an interface for testing purposes."
'@Interface
Option Explicit

'@Description "Gets the number of times the specified member was invoked in the lifetime of the object."
Public Property Get MemberInvokes(ByVal MemberName As String) As Long
Attribute MemberInvokes.VB_Description = "Gets the number of times the specified member was invoked in the lifetime of the object."
End Property

'@Description "Gets a string representation of the object's internal state, for debugging purposes (not intended for asserts!)."
Public Function ToString() As String
Attribute ToString.VB_Description = "Gets a string representation of the object's internal state, for debugging purposes (not intended for asserts!)."
End Function

'@Description "Fails a test when the member invokes for the specified member name don't match the expected count."
Public Sub Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
Attribute Verify.VB_Description = "Fails a test when the member invokes for the specified member name don't match the expected count."
End Sub
