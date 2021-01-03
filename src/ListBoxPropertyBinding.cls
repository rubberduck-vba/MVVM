VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder MVVM.Infrastructure.Bindings.PropertyBindings
'@PredeclaredId
'@Exposed
Option Explicit
Implements IPropertyBinding
Implements IDisposable
Implements IHandlePropertyChanged
Implements IHandleControlEvents

Private WithEvents TargetEventSource As MSForms.ListBox
Attribute TargetEventSource.VB_VarHelpID = -1
Private Const DefaultTargetControlProperty As String = "Value"
Private Const FormsControlProgId As String = "Forms.ListBox.1"

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Public Property Get ProgID() As String
    ProgID = FormsControlProgId
End Property

Public Function Create(ByVal Context As IAppContext, Source As IBindingPath, ByVal Target As MSForms.ListBox, _
Optional ByVal TargetProperty As Variant, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim LocalTargetProperty As String
    If IsMissing(TargetProperty) Then
        LocalTargetProperty = InferTargetPropertyFromSource(Source, outMode:=Mode)
    Else
        LocalTargetProperty = TargetProperty
    End If
    
    Dim BindingBase As PropertyBindingBase
    Set BindingBase = PropertyBindingBase _
        .Create(Context, Source, Target, LocalTargetProperty, _
            Mode:=Mode, _
            UpdateSource:=UpdateSource, _
            Converter:=Converter, _
            StringFormat:=StringFormat, _
            Validator:=Validator, _
            ValidationAdorner:=ValidationAdorner)
    
    Dim Result As ListBoxPropertyBinding
    Set Result = New ListBoxPropertyBinding
    
    If Not Validator Is Nothing Then
        BindingBase.AsINotifyValidationError.RegisterHandler ValidationManager
    End If
    
    Result.InjectBindingInfo BindingBase
    Set Create = Result
    
End Function

Private Function InferTargetPropertyFromSource(ByVal Source As IBindingPath, ByRef outMode As MVVM.BindingMode) As String
    Dim Result As String
    Dim SourceValue As Variant
    If Source.TryReadPropertyValue(outValue:=SourceValue) Then
        Select Case True
            Case VarType(SourceValue) = VbVarType.vbLong, _
                 VarType(SourceValue) = VbVarType.vbInteger, _
                 VarType(SourceValue) = VbVarType.vbDouble, _
                 VarType(SourceValue) = VbVarType.vbSingle, _
                 VarType(SourceValue) = VbVarType.vbByte, _
                Result = "ListIndex"
            Case VarType(SourceValue) = VbVarType.vbString
                Result = "Text"
            Case IsArray(SourceValue) Or VarType(SourceValue) = VbVarType.vbArray
                Debug.Print TypeName(Me) & ": binding source [" & Source.ToString & "] one-way."
                outMode = MVVM.BindingMode.OneWayBinding
                Result = "List"
            Case Else
                Debug.Print TypeName(Me) & ": binding source [" & Source.ToString & "] one-way to source."
                outMode = MVVM.BindingMode.OneWayToSource
                Result = DefaultTargetControlProperty
        End Select
    Else
        Debug.Print TypeName(Me) & ": Could not read source property value; binding to default target property."
        Result = DefaultTargetControlProperty
    End If
    InferTargetPropertyFromSource = Result
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, ListBoxPropertyBinding, TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Handler, TypeName(Me)
    Set This.Base = BindingInfo
    Set This.Handler = BindingInfo
    This.Base.AsIControlEvents.RegisterHandler Me
End Sub

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is TextBoxPropertyBinding
End Property

Private Sub IDisposable_Dispose()
    Set This.Handler = Nothing
    Disposable.TryDispose This.Base
    Set This.Base = Nothing
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef Cancel As Boolean)
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IHandleControlEvents_HandleChange()
End Sub

Private Sub IHandleControlEvents_HandleEnter()
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef Cancel As Boolean)
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Handler.HandlePropertyChanged Source, PropertyName
End Sub

Private Sub IPropertyBinding_Apply()
    This.Base.Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = This.Base.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = DefaultTargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.Base.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = This.Base.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = This.Base.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = This.Base.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = This.Base.Validator
End Property

Private Sub TargetEventSource_Change()
    If This.Base.UpdateSourceTrigger = OnPropertyChanged Then This.Base.ApplyToSource
End Sub

