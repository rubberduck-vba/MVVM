VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowseCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Example
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    If TypeOf Context Is ExampleViewModel Then
        Dim ViewModel As ExampleViewModel
        Set ViewModel = Context
        ICommand_CanExecute = ViewModel.SomeOtherOption
    End If
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Browse for source file..."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    Dim ViewModel As ExampleViewModel
    Set ViewModel = Context
    
    ViewModel.SomeFilePath = "TEST"
    
    MsgBox "Implement me!", vbInformation, "Hello from " & TypeName(Me)
End Sub
