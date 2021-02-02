VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxPropertyBinding"
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
Implements IHandleControlEvents
Implements IHandlePropertyChanged

Private WithEvents TargetEventSource As MSForms.TextBox
Attribute TargetEventSource.VB_VarHelpID = -1

Private Const DefaultTargetControlProperty As String = "Text"
Private Const FormsControlProgId As String = "Forms.TextBox.1"

Private Type TState
    Base As PropertyBindingBase
    Notifier As INotifyPropertyChanged
    Handler As IHandlePropertyChanged
End Type

Private This As TState

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Public Property Get ProgID() As String
    ProgID = FormsControlProgId
End Property

Public Function Create(ByVal Context As IAppContext, ByVal Source As IBindingPath, ByVal Target As Object, _
Optional ByVal TargetProperty As String = DefaultTargetControlProperty, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnExit, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal StringFormat As IStringFormatter, _
Optional ByVal ValidationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim BindingBase As PropertyBindingBase
    Set BindingBase = PropertyBindingBase _
        .Create(Context, Source, Target, _
            TargetPropertyPath:=TargetProperty, _
            Mode:=Mode, _
            UpdateSource:=UpdateSource, _
            Validator:=Validator, _
            Converter:=Converter, _
            StringFormat:=StringFormat, _
            ValidationAdorner:=ValidationAdorner)
    
    Dim Result As TextBoxPropertyBinding
    Set Result = New TextBoxPropertyBinding
    
    Result.InjectBindingInfo BindingBase
    Set Create = Result
    
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, TextBoxPropertyBinding, TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, TypeName(Me)
    GuardClauses.GuardExpression Not TypeOf BindingInfo.Target.Context Is MSForms.TextBox
    GuardClauses.GuardDoubleInitialization This.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Handler, TypeName(Me)
    Set This.Base = BindingInfo
    Set This.Handler = BindingInfo
    Set TargetEventSource = BindingInfo.Target.Context
    This.Base.AsIControlEvents.RegisterHandler Me
End Sub

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is TextBoxPropertyBinding
End Property

'BUG: fixed 'VF: logical error with the validation if space bar pressed <- fixed by switch to _Change/OnChange from KeyPress
'BUG: fixed 'VF: bug mouse click triggered selectall but select all should not be executed on mouse events so like a user clicks in the entered string to correct a spelling error
Private Sub SelectAll()
    TargetEventSource.SelStart = 0
    TargetEventSource.SelLength = TargetEventSource.TextLength
End Sub

Private Sub IDisposable_Dispose()
    Set This.Handler = Nothing
    Disposable.TryDispose This.Base
    Set This.Base = Nothing
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef Cancel As Boolean)
End Sub
'VF: added as replacement for KeyPress to capture cut/paste and backspace/delete
Private Sub IHandleControlEvents_HandleChange()
    '? dont't create an extra Private Sub OnChange()
    IHandlePropertyChanged_HandlePropertyChanged TargetEventSource, "StringProperty"
End Sub
'Note: 'VF: not sure of added value of selecting all upon enter,  probably a matter of every users' habit/preferences
Private Sub IHandleControlEvents_HandleEnter()
    SelectAll
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
'VF: event added to capture cut/paste and backspace/delete
Private Sub TargetEventSource_Change()
    If This.Base.UpdateSourceTrigger <> MVVM.BindingUpdateSourceTrigger.OnChange Then Exit Sub
    
    Dim Cancel As Boolean
    'VF: pass in entire textbox content for verification
    Dim outText As String
    If Not This.Base.Target.TryReadPropertyValue(outText) Then 'pass invalid string for further validation
        This.Base.OnBindingUpdateSourceOpportunity Cancel, MVVM.BindingUpdateSourceTrigger.OnChange, vbNullString
    Else
        Debug.Print "TargetEventSource_Change"; outText
        This.Base.OnBindingUpdateSourceOpportunity Cancel, MVVM.BindingUpdateSourceTrigger.OnChange, outText ', Chr$(KeyAscii.Value)
    End If
    
    'not needed if text and if numeric textbox leave correction up to user as instructed by validation error adorner
'    If Cancel Then
'        'VF: delete invalid text?
'        'more testing needed with all sorts of textbox usages
'        'check return value? and raise error or do something else if false?
'        This.Base.Target.TryWritePropertyValue (vbNullString)
'    End If
    
End Sub
'VF: required by numeric textbox, text textbox must use _Change
Private Sub TargetEventSource_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If This.Base.UpdateSourceTrigger <> OnKeyPress Then Exit Sub

    Dim Cancel As Boolean
    This.Base.OnBindingUpdateSourceOpportunity Cancel, MVVM.BindingUpdateSourceTrigger.OnKeyPress, Chr$(KeyAscii.Value)

    If Cancel Then
        'swallow the invalid keypress:
        KeyAscii.Value = 0
        'VF: but don't select all, just swallow invalid keypress <- was caused be calling OnEnter, which is commented out now
        'VF: must re-run
        'ToDo: 'VF: refactor to have clean structure
        '   (a) logic error non-numeric key is swallowed so error no longer exists but validation error adorner sticks
        'rerun with numeric char to remove validation error adorner
        This.Base.OnBindingUpdateSourceOpportunity Cancel, MVVM.BindingUpdateSourceTrigger.OnKeyPress, 0 'Chr$(KeyAscii.Value)
        'either add delay for user to see validation error adorner or inform by msgbox why his keystrokes don't show
        MsgBox "Only numeric entries accepted.", vbInformation, "'ToDo: 'VF: use caption of control or userform or something"
        '   (b) cut/paste and backspace/delete not captured so maybe have two triggers: KeyPress and Change or just funnel through Change and let user delete non-numeric chars?
    End If

End Sub
'VF: commented out as dysfunctional from (my) user and developer perspective: click into string in textbox to add text or to correct a spelling error must be possible and this sub would select all
'if user wants to select all, he can do so himself
'Private Sub TargetEventSource_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    'https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/mousedown-mouseup-events
'    Const fmButtonLeft As Long = 1
'    If Button = fmButtonLeft And TargetEventSource.SelLength = 0 Then SelectAll
'End Sub

