Attribute VB_Name = "Example"
'@Folder MVVM.Example
Option Explicit

'@Description "Runs the MVVM example UI."
Public Sub Run()
Attribute Run.VB_Description = "Runs the MVVM example UI."
'here a more elaborate application would wire-up dependencies for complex commands,
'and then property-inject them into the ViewModel via a factory method e.g. SomeViewModel.Create(args).

    Dim ViewModel As ExampleViewModel
    Set ViewModel = ExampleViewModel.Create
    
    'ViewModel properties can be set before or after it's wired to the View.
    'ViewModel.SourcePath = "TEST"
    ViewModel.SomeOption = True
    
    Set ViewModel.SomeCommand = New BrowseCommand
    
    Dim App As AppContext
    Set App = AppContext.Create(DebugOutput:=True)
    
    ViewModel.BooleanProperty = False
    ViewModel.ByteProperty = 240
    ViewModel.DateProperty = VBA.DateTime.Now + 2
    ViewModel.DoubleProperty = 85
    ViewModel.StringProperty = "Beta"
    ViewModel.LongProperty = -42
    
    Dim View As IView
    Set View = ExampleView.Create(App, ViewModel)
    
    If View.ShowDialog Then
        Debug.Print ViewModel.SomeFilePath, ViewModel.SomeOption, ViewModel.SomeOtherOption
    Else
        Debug.Print "Dialog was cancelled."
    End If
    
    Disposable.TryDispose App
    
End Sub

Public Sub DynamicRun()
    Dim Context As IAppContext
    Set Context = AppContext.Create
    
    Dim ViewModel As ExampleViewModel
    Set ViewModel = ExampleViewModel.Create
    
    Dim View As IView
    Set View = ExampleDynamicView.Create(Context, ViewModel)
    
    Debug.Print View.ShowDialog
    
End Sub
