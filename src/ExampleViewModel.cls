VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An example ViewModel implementation for some dialog."
'@Folder MVVM.Example
'@ModuleDescription "An example ViewModel implementation for some dialog."
'@PredeclaredId
'@Exposed
Implements INotifyPropertyChanged
Option Explicit

Private Type TViewModel
    Notifier As INotifyPropertyChanged
    
    SomeCommand As ICommand
    SomeFilePath As String
    SomeAmount As Currency
    SomeDate As Date
    SomeProperty As String
    SomeOption As Boolean
    SomeOtherOption As Boolean
    SomeItems As Variant
    SelectedItemText As String
    SelectedItemIndex As Long
    BooleanProperty As Boolean
    ByteProperty As Byte
    CurrencyProperty As Currency
    DateProperty As Date
    DoubleProperty As Double
    LongProperty As Long
    StringProperty As String
End Type

Private This As TViewModel

Public Function Create() As ExampleViewModel
    GuardClauses.GuardNonDefaultInstance Me, ExampleViewModel, TypeName(Me)
    
    Dim Result As ExampleViewModel
    Set Result = New ExampleViewModel
    
    Set Create = Result
End Function

Public Property Get BooleanProperty() As Boolean
    BooleanProperty = This.BooleanProperty
End Property

Public Property Let BooleanProperty(ByVal RHS As Boolean)
    If This.BooleanProperty <> RHS Then
        This.BooleanProperty = RHS
        OnPropertyChanged "BooleanProperty"
    End If
End Property

Public Property Get ByteProperty() As Byte
    ByteProperty = This.ByteProperty
End Property

Public Property Let ByteProperty(ByVal RHS As Byte)
    If This.ByteProperty <> RHS Then
        This.ByteProperty = RHS
        OnPropertyChanged "ByteProperty"
    End If
End Property

Public Property Get CurrencyProperty() As Currency
    CurrencyProperty = This.CurrencyProperty
End Property

Public Property Let CurrencyProperty(ByVal RHS As Currency)
    If This.CurrencyProperty <> RHS Then
        This.CurrencyProperty = RHS
        OnPropertyChanged "CurrencyProperty"
    End If
End Property

Public Property Get DateProperty() As Date
    DateProperty = This.DateProperty
End Property

Public Property Let DateProperty(ByVal RHS As Date)
    If This.DateProperty <> RHS Then
        This.DateProperty = RHS
        OnPropertyChanged "DateProperty"
    End If
End Property

Public Property Get DoubleProperty() As Double
    DoubleProperty = This.DoubleProperty
End Property

Public Property Let DoubleProperty(ByVal RHS As Double)
    If This.DoubleProperty <> RHS Then
        This.DoubleProperty = RHS
        OnPropertyChanged "DoubleProperty"
    End If
End Property

Public Property Get LongProperty() As Long
    LongProperty = This.LongProperty
End Property

Public Property Let LongProperty(ByVal RHS As Long)
    If This.LongProperty <> RHS Then
        This.LongProperty = RHS
        OnPropertyChanged "LongProperty"
    End If
End Property

Public Property Get StringProperty() As String
    StringProperty = This.StringProperty
End Property

Public Property Let StringProperty(ByVal RHS As String)
    If This.StringProperty <> RHS Then
        This.StringProperty = RHS
        OnPropertyChanged "StringProperty"
    End If
End Property

Public Property Get SomeFilePath() As String
    SomeFilePath = This.SomeFilePath
End Property

Public Property Let SomeFilePath(ByVal RHS As String)
    If This.SomeFilePath <> RHS Then
        This.SomeFilePath = RHS
        OnPropertyChanged "SomeFilePath"
    End If
End Property

Public Property Get SomeItems() As Variant
    SomeItems = This.SomeItems
End Property

Public Property Get SelectedItemText() As String
    SelectedItemText = This.SelectedItemText
End Property

Public Property Let SelectedItemText(ByVal RHS As String)
    If This.SelectedItemText <> RHS Then
        This.SelectedItemText = RHS
        OnPropertyChanged "SelectedItemText"
    End If
End Property

Public Property Get SelectedItemIndex() As Long
    SelectedItemIndex = This.SelectedItemIndex
End Property

Public Property Let SelectedItemIndex(ByVal RHS As Long)
    If This.SelectedItemIndex <> RHS Then
        This.SelectedItemIndex = RHS
        OnPropertyChanged "SelectedItemIndex"
    End If
End Property

Public Property Get SomeAmount() As Currency
    SomeAmount = This.SomeAmount
End Property

Public Property Let SomeAmount(ByVal RHS As Currency)
    If This.SomeAmount <> RHS Then
        This.SomeAmount = RHS
        OnPropertyChanged "SomeAmount"
    End If
End Property

Public Property Get SomeDate() As Date
    SomeDate = This.SomeDate
End Property

Public Property Let SomeDate(ByVal RHS As Date)
    If This.SomeDate <> RHS Then
        This.SomeDate = RHS
        OnPropertyChanged "SomeDate"
    End If
End Property

Public Property Get SomeProperty() As String
    SomeProperty = This.SomeProperty
End Property

Public Property Let SomeProperty(ByVal RHS As String)
    If This.SomeProperty <> RHS Then
        This.SomeProperty = RHS
        OnPropertyChanged "SomeProperty"
    End If
End Property

Public Property Get SomeOption() As Boolean
    SomeOption = This.SomeOption
End Property

Public Property Let SomeOption(ByVal RHS As Boolean)
    If This.SomeOption <> RHS Then
        This.SomeOption = RHS
        OnPropertyChanged "SomeOption"
    End If
End Property

Public Property Get SomeOtherOption() As Boolean
    SomeOtherOption = This.SomeOtherOption
End Property

Public Property Let SomeOtherOption(ByVal RHS As Boolean)
    If This.SomeOtherOption <> RHS Then
        This.SomeOtherOption = RHS
        OnPropertyChanged "SomeOtherOption"
    End If
End Property

Public Property Get SomeCommand() As ICommand
    Set SomeCommand = This.SomeCommand
End Property

Public Property Set SomeCommand(ByVal RHS As ICommand)
    Set This.SomeCommand = RHS
End Property

Public Property Get SomeOptionName() As String
    SomeOptionName = "Auto"
End Property

Public Property Get SomeOtherOptionName() As String
    SomeOtherOptionName = "Manual/Browse"
End Property

Public Property Get Instructions() As String
    Instructions = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Me, PropertyName
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifierBase
    ReDim This.SomeItems(0 To 4)
    Dim i As Long
    For i = LBound(This.SomeItems) To UBound(This.SomeItems)
        This.SomeItems(i) = "Selection " & i
    Next
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub


