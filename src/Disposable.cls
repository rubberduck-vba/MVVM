VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Disposable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates the necessary casting operation to invoke IDisposable.Dispose."
'@Folder MVVM.Infrastructure
'@ModuleDescription "An object that encapsulates the necessary casting operation to invoke IDisposable.Dispose."
'@PredeclaredId
'@Exposed
Option Explicit

'@Description "True if a Dispose method was successfully invoked (early-bound IDisposable.Dispose unless specified otherwise)."
Public Function TryDispose(ByVal Target As Object, Optional ByVal TryLateBound As Boolean = False) As Boolean
Attribute TryDispose.VB_Description = "True if a Dispose method was successfully invoked (early-bound IDisposable.Dispose unless specified otherwise)."
    Dim DisposableTarget As IDisposable
    If TryCastToDisposable(Target, outResult:=DisposableTarget) Then
        DisposableTarget.Dispose
        TryDispose = True
    ElseIf TryLateBound Then
        On Error Resume Next
        Target.Dispose
        TryDispose = (Err.Number = 0)
        On Error GoTo 0
    End If
End Function

Private Function TryCastToDisposable(ByVal Target As Object, ByRef outResult As IDisposable) As Boolean
    If TypeOf Target Is IDisposable Then
        Set outResult = Target
        TryCastToDisposable = True
    End If
End Function
