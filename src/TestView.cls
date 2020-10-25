VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an IView dependency in a unit test."
'@Folder Tests.Stubs
'@ModuleDescription "An object that can stub an IView dependency in a unit test."
Option Explicit
Implements ITestStub
Implements IView
Implements ICancellable

Private Type TState
    Cancel As Boolean
    Stub As TestStubBase
    Shown As Boolean
    ViewModel As Object
End Type

Private This As TState

'@Description "Gets a value indicating whether the stub IView is in a visible (shown) state or not."
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Gets a value indicating whether the stub IView is in a visible (shown) state or not."
    Visible = This.Shown
End Property

'@Description "Gets or sets the ViewModel object this stub is working with."
Public Property Get ViewModel() As Object
Attribute ViewModel.VB_Description = "Gets or sets the ViewModel object this stub is working with."
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set This.ViewModel = RHS
End Property

'@Description "Makes IView.ShowDialog return False (simulates a cancelled dialog)."
Public Sub OnCancel()
Attribute OnCancel.VB_Description = "Makes IView.ShowDialog return False (simulates a cancelled dialog)."
    This.Cancel = True
    This.Shown = False
End Sub

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.Cancel
End Property

Private Sub ICancellable_OnCancel()
    This.Stub.OnInvoke "OnCancel"
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

Private Sub IView_Hide()
    This.Stub.OnInvoke "Hide"
    This.Shown = False
End Sub

Private Sub IView_Show()
    This.Stub.OnInvoke "Show"
    This.Shown = True
End Sub

Private Function IView_ShowDialog() As Boolean
    This.Stub.OnInvoke "ShowDialog"
    This.Shown = Not This.Cancel
    IView_ShowDialog = Not This.Cancel
End Function

Private Property Get IView_ViewModel() As Object
    This.Stub.OnInvoke "ViewModel.Get"
    Set IView_ViewModel = This.ViewModel
End Property
