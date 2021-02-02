VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBindingBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder MVVM.Infrastructure.Bindings
'@PredeclaredId
'@Exposed
Option Explicit
Implements IPropertyBinding
Implements IDisposable

Implements IControlEvents   '<~ relays MSForms.Control events to more specific IPropertyBinding implementations
Implements IHandleControlEvents '<~ ControlEventPunk callbacks
Implements IHandlePropertyChanged '<~ handles INotifyPropertyChange events from ViewModel (i.e. "applies" the binding)

Private Const DefaultTargetControlProperty As String = "Value"

Public Enum ApplyResult
    BindingSkipped
    BindingSuccess
    BindingValidationError
    BindingConversionError
    BindingFailed
End Enum

Private Type TState
    Applied As Boolean
    Applying As Boolean
    
    Punk As ControlEventsPunk
    Handlers As Collection
    
    Context As IAppContext
    Source As IBindingPath
    Target As IBindingPath
    
    Mode As BindingMode
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    
    Converter As IValueConverter
    StringFormat As IStringFormatter
    
    Validator As IValueValidator
    ValidationAdorner As IDynamicAdorner ' data validations (might be unset)
    CancelExitOnValidationError As Boolean
    
End Type

Private This As TState

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal TargetContext As Object, _
Optional ByVal TargetPropertyPath As String = DefaultTargetControlProperty, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnExit, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal ValidationAdorner As IDynamicAdorner, _
Optional ByVal CancelExitOnValidationError As Boolean = True) As IPropertyBinding

    GuardClauses.GuardNonDefaultInstance Me, PropertyBindingBase, TypeName(Me)
    
    Dim Result As PropertyBindingBase
    Set Result = New PropertyBindingBase
    
    Result.CancelExitOnValidationError = CancelExitOnValidationError
    Result.Mode = Mode
    Result.UpdateSourceTrigger = UpdateSource
    
    Set Result.Source = Source
    Set Result.Target = BindingPath.Create(TargetContext, TargetPropertyPath)
    
    Set Result.Context = Context
    
    Set Result.Converter = Converter
    Set Result.StringFormat = StringFormat
    
    Set Result.Validator = Validator
    Set Result.ValidationAdorner = ValidationAdorner
    
    Set Create = Result
    
End Function

Public Property Get AsIControlEvents() As IControlEvents
    Set AsIControlEvents = Me
End Property

Public Property Get AsINotifyValidationError() As INotifyValidationError
    Set AsINotifyValidationError = Me
End Property

Public Property Get Source() As IBindingPath
    Set Source = This.Source
End Property

Friend Property Set Source(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization This.Source, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set This.Source = RHS
End Property

Public Property Get Target() As IBindingPath
    Set Target = This.Target
End Property

Friend Property Set Target(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization This.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set This.Target = RHS
    If TypeOf This.Target.Context Is MSForms.Control Then
        On Error Resume Next
        Set This.Punk.Target = This.Target.Context
        If This.Punk.Connect Then
            This.Punk.RegisterHandler Me
        Else
            DebugMessage "Unable to connect target IConnectionPoint."
        End If
        If Err.Number <> 0 Then
            DebugMessage "Unable to handle MSForms.Control events. Binding is effectively one-way. " & Err.Description
        End If
        On Error GoTo 0
    Else
        Set This.Punk = Nothing
    End If
End Property

Public Property Get Mode() As BindingMode
    Mode = This.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    This.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = This.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    This.UpdateSourceTrigger = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = This.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    GuardClauses.GuardDoubleInitialization This.Converter, TypeName(Me)
    Set This.Converter = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = This.Validator
End Property

Friend Property Set Validator(ByVal RHS As IValueValidator)
    GuardClauses.GuardDoubleInitialization This.Validator, TypeName(Me)
    Set This.Validator = RHS
    If Not This.Validator Is Nothing Then This.UpdateSourceTrigger = This.Validator.Trigger
End Property

Public Property Get StringFormat() As IStringFormatter
    Set StringFormat = This.StringFormat
End Property

Friend Property Set StringFormat(ByVal RHS As IStringFormatter)
    GuardClauses.GuardDoubleInitialization This.StringFormat, TypeName(Me)
    Set This.StringFormat = RHS
End Property

Public Property Get ValidationAdorner() As IDynamicAdorner
    Set ValidationAdorner = This.ValidationAdorner
End Property

Friend Property Set ValidationAdorner(ByVal RHS As IDynamicAdorner)
    GuardClauses.GuardDoubleInitialization This.ValidationAdorner, TypeName(Me)
    Set This.ValidationAdorner = RHS
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Friend Property Set Context(ByVal RHS As IAppContext)
    GuardClauses.GuardDoubleInitialization This.Context, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set This.Context = RHS
End Property

Public Property Get CancelExitOnValidationError() As Boolean
    CancelExitOnValidationError = This.CancelExitOnValidationError
End Property

Friend Property Let CancelExitOnValidationError(ByVal RHS As Boolean)
    This.CancelExitOnValidationError = RHS
End Property

Private Function TryConvert(ByRef Value As Variant) As Boolean
    
    If This.Converter Is Nothing Then
        TryConvert = True
        Exit Function
    End If
    
    On Error Resume Next
    Value = This.Converter.Convert(Value)
    If Err.Number = 0 Then
        DebugMessage "Value was successfully converted."
        TryConvert = True
    Else
        DebugMessage "IValueConverter.Convert raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

Private Function TryConvertBack(ByRef Value As Variant) As Boolean
    
    If This.Converter Is Nothing Then
        TryConvertBack = True
        Exit Function
    End If
    
    On Error Resume Next
    Value = This.Converter.ConvertBack(Value)
    If Err.Number = 0 Then
        DebugMessage "Value was successfully converted back."
        TryConvertBack = True
    Else
        DebugMessage "IValueConverter.ConvertBack raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

Public Sub Apply()
    If This.Applying Then Exit Sub '<~ could break accidental recursion... leave this here...
    
    If This.Mode = OneTimeBinding Then
        If Not This.Applied Then ApplyToTarget
        
    ElseIf This.Mode = OneWayBinding Or This.Mode = TwoWayBinding Then
        ApplyToTarget
        
    ElseIf This.Mode = OneWayToSource Then
        ApplyToSource
        
    End If

End Sub

Private Property Get CanApplyToTarget() As Boolean
    'one-way to source and already-applied one-time binding modes do not apply to target
    CanApplyToTarget = Not This.Applying And This.Mode <> OneWayToSource And (This.Mode <> OneTimeBinding Or Not This.Applied)
End Property

Public Sub ApplyToTarget()
'reads from the source and writes to the target.
    
    If Not CanApplyToTarget Then Exit Sub
    This.Applying = True
    
    Dim SourceValue As Variant
    Dim UseFallbackDefault As Boolean
    
    If This.Source.Object Is Nothing Then
        UseFallbackDefault = TryGetDefaultBindingValue(outValue:=SourceValue)
    End If
    
    This.Source.Resolve
    
    If This.Source.Object Is Nothing And Not UseFallbackDefault Then
        DebugMessage "ApplyToTarget aborted: source object in path '" & This.Source.Path & "' is Nothing and there is no fallback default value for this binding."
        This.Applying = False
        Exit Sub
        
    ElseIf This.Source.Object Is Nothing Then
        DebugMessage "ApplyToTarget: source object in path '" & This.Source.Path & "' is Nothing; binding target property '" & This.Target.PropertyName & "' to default/fallback value."
        
    Else
        If Not This.Source.TryReadPropertyValue(outValue:=SourceValue) Then
            DebugMessage "ApplyToTarget failed to read current source property value."
            This.Applying = False
            Exit Sub
        End If
    End If
    
    Validate SourceValue, Propagate:=False
    SourceValue = Convert(SourceValue)
    
    This.Target.Resolve
    Dim CurrentValue As Variant
    If Not This.Target.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToTarget failed to read current target property value."
        This.Applying = False
        Exit Sub
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(SourceValue) Then
        Proceed = SourceValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        If Not This.StringFormat Is Nothing Then
            SourceValue = This.StringFormat.Format(SourceValue)
        End If
        
        If Not This.Target.TryWritePropertyValue(SourceValue) Then
            DebugMessage "ApplyToTarget failed to write to target property."
        Else
            DebugMessage "ApplyToTarget successfully applied to target."
            This.Applied = True
        End If
    End If
    This.Applying = False
End Sub

Private Property Get CanApplyToSource() As Boolean
    'one-time and one-way binding modes do not write values to the binding source (ViewModel).
    CanApplyToSource = Not This.Applying And This.Mode <> OneTimeBinding And This.Mode <> OneWayBinding
End Property

Public Function ApplyToSource(Optional ByVal TargetValidationValue As Variant) As ApplyResult
'reads from the target and writes to the source.
'when This.Mode=KeyPress, the TargetValidationValue contains the KeyAscii value.
    On Error GoTo CleanFail

    If Not CanApplyToSource Then Exit Function
    This.Applying = True
    
    This.Target.Resolve
    Dim TargetValue As Variant
    If Not This.Target.TryReadPropertyValue(outValue:=TargetValue) Then
        DebugMessage "ApplyToSource failed to read target property value."
        ApplyToSource = ApplyResult.BindingFailed
        This.Applying = False
        Exit Function
    End If
    
    If IsMissing(TargetValidationValue) Then
        TargetValidationValue = TargetValue
        
    ElseIf This.UpdateSourceTrigger = BindingUpdateSourceTrigger.OnKeyPress Then
        'keypress hasn't reached the control yet:
        TargetValue = InsertKeypress(TargetValue, TargetValidationValue)
        
    End If
    
    If Not TryConvertBack(TargetValue) Then
        ApplyToSource = ApplyResult.BindingConversionError
        DebugMessage TypeName(This.Converter) & " ConvertBack failed to convert the target value."
        OnEnter
        This.Applying = False
        Exit Function
    End If
    
    If Not Validate(TargetValidationValue) Then
        ApplyToSource = ApplyResult.BindingValidationError
        DebugMessage TypeName(This.Validator) & " failed to validate the target value."
        This.Applying = False
        Exit Function
    End If
    
    This.Source.Resolve
    Dim CurrentValue As Variant
    If Not This.Source.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToSource failed to read current source property value."
        ApplyToSource = ApplyResult.BindingFailed
        This.Applying = False
        Exit Function
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(CurrentValue) Then
        Proceed = TargetValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        ClearValidationErrors
        
        If Not This.Source.TryWritePropertyValue(TargetValue) Then
            DebugMessage "ApplyToSource failed to write to source property."
            ApplyToSource = ApplyResult.BindingFailed
            Exit Function
        End If
        
        DebugMessage "Binding was successfully applied to source."
        ApplyToSource = ApplyResult.BindingSuccess
        
    End If
    
CleanExit:
    This.Applying = False
    Exit Function
    
CleanFail:
    DebugMessage "ApplyToSource failed unexpectedly."
    ApplyToSource = ApplyResult.BindingFailed
    Resume CleanExit
    Resume
End Function

Private Function InsertKeypress(ByVal InitialValue As String, ByVal Char As String) As String
'FIXME this function is at the wrong abstraction level.
    
    If Char = vbNullString Then
        InsertKeypress = InitialValue
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim SelectionStart As Long
    SelectionStart = This.Target.Object.SelStart
    
    Dim SelectionLength As Long
    SelectionLength = This.Target.Object.SelLength
    
    On Error GoTo 0
    
    Select Case True
        
        Case SelectionStart > 0 And SelectionLength = 0
            
            'insert at SelStart
            With StringBuilderNet.Append(InitialValue) _
                              .Insert(SelectionStart, Char)
                InsertKeypress = .ToString
            End With
        
        Case SelectionStart > 0 And SelectionLength > 0
            
            'replace selection
            With StringBuilderNet.Append(InitialValue) _
                              .Remove(SelectionStart, SelectionLength) _
                              .Insert(SelectionStart, Char)
                     
                InsertKeypress = .ToString
            End With
            
        Case Else
            
            'replace initial value
            InsertKeypress = Char
    
    End Select
    
End Function

Private Function IsValid() As Boolean
    IsValid = This.Context.Validation.IsValid(This.Source.Object, This.Source.PropertyName)
End Function

Private Sub ClearValidationErrors()
    With This.Context
        .Validation.ClearValidationError This.Source
        .Commands.EvaluateCanExecute This.Source
    End With
End Sub

Private Function ToString() As String
    ToString = TypeName(This.Source.Context) & "." & This.Source.Path & " -> " & TypeName(This.Target.Context) & "." & This.Target.Path
End Function

Private Function ResolvePropertyPath(ByVal Source As Object, ByVal PropertyPath As String) As Object
    
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    
    If UBound(Parts) = LBound(Parts) Then
        Set ResolvePropertyPath = Source
    Else
        Dim RecursiveProperty As Object
        Set RecursiveProperty = CallByName(Source, Parts(0), VbGet)
        If RecursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(RecursiveProperty, Right$(PropertyPath, Len(PropertyPath) - Len(Parts(0)) - 1))
    End If
    
End Function

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    ResolvePropertyName = Parts(UBound(Parts))
End Function

Private Function TryGetDefaultBindingValue(ByRef outValue As Variant) As Boolean
'Gets a default value for certain specific target properties, used when source path cannot be fully resolved,
'e.g. when target binds to "SomeObjectProperty.SomeProperty" and "SomeObjectProperty" is Nothing.

    Select Case This.Target.PropertyName
    
        Case "Text", "Caption"
            outValue = vbNullString
            TryGetDefaultBindingValue = True
            
        Case "Enabled", "Visible"
            outValue = False
            TryGetDefaultBindingValue = True
            
        Case "Value"
            If TypeOf This.Target Is MSForms.CheckBox _
                Or TypeOf This.Target Is MSForms.OptionButton _
            Then
                outValue = False
                TryGetDefaultBindingValue = True
            End If
            
    End Select
End Function

Private Function Convert(ByVal Value As Variant) As Variant
    
    On Error GoTo CleanFail
    If Not This.Converter Is Nothing Then
        Convert = This.Converter.Convert(Value)
    Else
        Convert = Value
    End If
    
CleanExit:
    Exit Function
    
CleanFail:
    DebugMessage "IValueConverter failed with error " & Err.Number & ": " & Err.Description
    OnEnter
    CustomErrors.RethrowOnError
    Resume CleanExit 'NOTE: 'VF: is never reached
End Function

Public Function Validate(ByVal TargetValue As Variant, Optional ByVal Propagate As Boolean = True) As Boolean

    If This.Validator Is Nothing Then
        Validate = True
        Exit Function
    End If
    
    On Error GoTo CleanFail
    If This.Validator.IsValid(TargetValue, This.Source, This.Target) Then
        
        If Propagate Then
            This.Context.Validation.ClearValidationError This.Source
            If Not This.ValidationAdorner Is Nothing Then This.ValidationAdorner.Hide
        End If
        
        Validate = True
        
    Else
        
        If Propagate Then
            Dim Message As String
            Message = This.Validator.Message
            
            This.Context.Validation.OnValidationError _
                Context:=This.Source.Context, _
                ValidationError:=ValidationError.Create(Me, Message)
            
            If Not This.ValidationAdorner Is Nothing Then This.ValidationAdorner.Show Message
            'VF: commented out as it leads to select all, which is not desired in text and numeric textboxes even though there may be more than one eventhandler
            'OnEnter
        End If
    End If
        
CleanExit:
    If Propagate Then
        This.Context.Commands.EvaluateCanExecute This.Source.Context
    End If
    Exit Function
    
CleanFail:
    DebugMessage "IValueValidator failed with error " & Err.Number & ": " & Err.Description
    OnEnter
    Resume CleanExit
    Resume
End Function

Private Sub Format()
    If This.StringFormat Is Nothing Then Exit Sub

    Dim BindingValue As Variant
    If Not This.Source.TryReadPropertyValue(outValue:=BindingValue) Then Exit Sub
        
    On Error GoTo CleanFail
    Dim FormattedValue As String
    FormattedValue = This.StringFormat.Format(BindingValue)
    On Error GoTo 0
    
    This.Applying = True
    If Not This.Target.TryWritePropertyValue(FormattedValue) Then
        DebugMessage "Could not format the binding value."
    End If
    This.Applying = False
    
CleanExit:
    Exit Sub
CleanFail:
    DebugMessage "IStringFormatter.Format() raised an error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub OnAfterUpdate()
    Format
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleAfterUpdate
    Next
End Sub

Private Sub OnBeforeUpdate(ByRef Cancel As Boolean)
    OnBindingUpdateSourceOpportunity Cancel, BindingUpdateSourceTrigger.OnPropertyChanged
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleBeforeUpdate Cancel
        If Cancel Then Exit For
    Next
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub OnChange()
    Format
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleChange
    Next
End Sub

Private Sub OnEnter()
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleEnter
    Next
End Sub

Private Sub OnExit(ByRef Cancel As Boolean)
    OnBindingUpdateSourceOpportunity Cancel, BindingUpdateSourceTrigger.OnExit
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleExit Cancel
    Next
End Sub

Public Sub OnBindingUpdateSourceOpportunity(ByRef Cancel As Boolean, ByVal Trigger As MVVM.BindingUpdateSourceTrigger, Optional ByVal TargetValidationValue As Variant)
    If Not CanApplyToSource Then Exit Sub
    
    Dim Result As MVVM.ApplyResult
    Select Case Trigger
        
        'VF: ? remove OnKeyPress here?, at least in TextBox no longer used
        'VF: changed: required to capture Cut/Paste and Backspace/Delete in TextBox
        'Case BindingUpdateSourceTrigger.OnKeyPress
        Case BindingUpdateSourceTrigger.OnKeyPress, BindingUpdateSourceTrigger.OnChange
        
            'If This.UpdateSourceTrigger = OnKeyPress And Not IsMissing(TargetValidationValue) Then
            If (This.UpdateSourceTrigger = OnKeyPress Or This.UpdateSourceTrigger = MVVM.BindingUpdateSourceTrigger.OnChange) _
                And Not IsMissing(TargetValidationValue) Then
                'just validate the keypress 'VF: or string pasted
                Result = IIf(Validate(TargetValidationValue), _
                    ApplyResult.BindingSuccess, _
                    ApplyResult.BindingValidationError)
                            
            End If
            
        Case BindingUpdateSourceTrigger.OnPropertyChanged
            
            If This.UpdateSourceTrigger = OnPropertyChanged Or This.UpdateSourceTrigger = OnKeyPress Then
                Result = ApplyToSource
            End If
        
        Case BindingUpdateSourceTrigger.OnExit
            
            If This.UpdateSourceTrigger <> Never Then
                
                Dim outValue As Variant
                If This.Target.TryReadPropertyValue(outValue) Then
                
                    Result = IIf(Validate(outValue), _
                        ApplyResult.BindingSuccess, _
                        ApplyResult.BindingValidationError)
                        
                End If
            End If
            
    End Select
    
    Cancel = This.CancelExitOnValidationError And (Cancel Or (Result = BindingValidationError) Or (Result = BindingConversionError))
    
End Sub

Private Sub OnValidationError()
    This.Context.Validation.OnValidationError This.Source.Context, ValidationError.Create(Me, This.Validator.Message)
    This.ValidationAdorner.Show This.Validator.Message
End Sub

Private Sub DebugMessage(ByVal Message As String)
    If This.Context.Bindings.DebugOutput Then Debug.Print "PropertyBinding: [" & ToString & "] " & Message
End Sub

Private Sub Class_Initialize()
    If Not Me Is PropertyBindingBase Then
        Set This.Handlers = New Collection
        Set This.Punk = New ControlEventsPunk
    End If
End Sub

Private Sub Class_Terminate()
    Set This.Handlers = Nothing
    If Not This.Punk Is Nothing Then
        This.Punk.Disconnect
        Set This.Punk = Nothing
    End If
End Sub

Private Sub IControlEvents_OnAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IControlEvents_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    OnBeforeUpdate LocalCancel
    Cancel.Value = LocalCancel
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IControlEvents_OnChange()
    OnChange
End Sub

Private Sub IControlEvents_OnEnter()
    OnEnter
End Sub

Private Sub IControlEvents_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    OnExit LocalCancel
    Cancel = LocalCancel
End Sub

Private Sub IControlEvents_RegisterHandler(ByVal Handler As IHandleControlEvents)
    This.Handlers.Add Handler
End Sub

Private Sub IDisposable_Dispose()
    
    Set This.Context = Nothing
    
    This.Punk.Disconnect
    Set This.Punk = Nothing
    Disposable.TryDispose This.ValidationAdorner
    Set This.ValidationAdorner = Nothing
    
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef Cancel As Boolean)
    OnBeforeUpdate Cancel
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IHandleControlEvents_HandleChange()
    OnChange
End Sub

Private Sub IHandleControlEvents_HandleEnter()
    OnEnter
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef Cancel As Boolean)
    OnExit Cancel
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    'IHandlePropertyChanged -> INotifyPropertyChanged: source is the ViewModel, so we're propagating to Target here:
    If Not This.Applying And Source Is This.Source.Object And PropertyName = This.Source.PropertyName Then
        ApplyToTarget
    End If
End Sub

Private Sub IPropertyBinding_Apply()
    Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = This.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = DefaultTargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = This.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = This.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = This.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = This.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = This.Validator
End Property


