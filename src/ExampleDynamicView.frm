VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleDynamicView 
   Caption         =   "ExampleDynamicView"
   ClientHeight    =   288
   ClientLeft      =   -456
   ClientTop       =   -1512
   ClientWidth     =   84
   OleObjectBlob   =   "ExampleDynamicView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleDynamicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder MVVM.Example
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As MVVM.IAppContext
    ViewModel As ExampleViewModel
    IsCancelled As Boolean
End Type

Private This As TState

'@Description "Creates a new instance of this form."
Public Function Create(ByVal Context As MVVM.IAppContext, ByVal ViewModel As ExampleViewModel, ViewDims As TViewDims) As IView
Attribute Create.VB_Description = "Creates a new instance of this form."
    Dim Result As ExampleDynamicView
    Set Result = New ExampleDynamicView
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel
    With Result
        .Height = ViewDims.Height
        .Width = ViewDims.Width
    End With
    Set Create = Result
End Function

Public Property Get Context() As MVVM.IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As MVVM.IAppContext)
    Set This.Context = RHS
End Property

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set This.ViewModel = RHS
End Property

Public Sub SizeView(Height As Long, Width As Long)
    With Me
        .Height = Height
        .Width = Width
    End With
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()
    
    Dim Layout As IContainerLayout
    Set Layout = ContainerLayout.Create(Me.Controls, TopToBottom)
    
    With DynamicControls.Create(This.Context, Layout)
        
        With .LabelFor("All controls on this form are created at run-time.")
            .Font.Bold = True
        End With
        
        .LabelFor BindingPath.Create(This.ViewModel, "Instructions")
        
        'VF: refactor free string to some enum PropertyName ("StringProperty", "CurrencyProperty") throughout (?) [when I frame a question mark in parentheses is not really a question but a rhetorical question, meaning I am pretty sure of the correct answer]
        .TextBoxFor BindingPath.Create(This.ViewModel, "StringProperty"), _
                    Validator:=New RequiredStringValidator, _
                    TitleSource:="Some String:"
                    
        .TextBoxFor BindingPath.Create(This.ViewModel, "CurrencyProperty"), _
                    FormatString:="{0:C2}", _
                    Validator:=New DecimalKeyValidator, _
                    TitleSource:="Some Amount:"
        
        'ToDo: 'VF: needs validation .CanExecute(This.Context) before .Show
        '(as textbox1 has focus and is empty and when moving to this close button, tb1 is validated and OnClick is disabled leaving the user out in the rain)
        .CommandButtonFor AcceptCommand.Create(Me, This.Context.Validation), This.ViewModel, "Close"
        
    End With
    
    This.Context.Bindings.Apply This.ViewModel
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeView
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    Me.Show vbModal
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
