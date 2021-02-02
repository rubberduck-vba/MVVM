Attribute VB_Name = "Example"
'@Folder MVVM.Example
Option Explicit

'VF: Windows 10 is having a hard time handling multiple monitors, especially if different resolutions and more so if legacy applications like the VBE
    'keeps shrinking the userform in the VBE and thus showing the shrunk form <- must counteract this ugly Windows bug by specifying Height and Width of IView
    'rendering engine was changed from 2010 to 2013
    'should go into IView, shouldn't it?
Public Type TViewDims
    Height As Long
    Width As Long
End Type

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
    Dim ViewDims As TViewDims
    'VF: in non-dynamic userforms like ExampleView the controls stay put so I use the right bottom most controls as anchor point like Me.Width = LastControl.left+LastControl.width + OffsetWidthPerOfficeVersion <- yes, userform are rendered differently depending on the version of Office 2007, ....
        'if sizing dynamically I would proceed likewise somehow with the (right bottom most) container <- is going to take quite an amount of code :-(
    With ViewDims
        .Height = 180 'some value that work in 2019, and somehow in 2010, too
        .Width = 230
    End With
    Set View = ExampleDynamicView.Create(Context, ViewModel, ViewDims)
    'or keep factory .Create 'clean'?
'    With ExampleDynamicView.Create(Context, ViewModel, ViewDims)
'        .SizeView 'not implemented
'        .ShowDialog
'        'payload DoSomething if not cancelled
'    End With
        
    Debug.Print View.ShowDialog
    
End Sub

