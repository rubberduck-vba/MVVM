VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotifierBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides a base implementation for the INotifyPropertyChanged observer interface."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "Provides a base implementation for the INotifyPropertyChanged observer interface."
Option Explicit
Implements INotifyPropertyChanged
Implements IDisposable

Private Type TState
    Handlers As Collection
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
End Sub

Private Sub IDisposable_Dispose()
    Set This.Handlers = Nothing
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Handlers.Add Handler
End Sub
