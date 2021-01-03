VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalKeyValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A key validator that allows numeric and decimal separator inputs."
'@Folder MVVM.Common.Validators
'@ModuleDescription "A key validator that allows numeric and decimal separator inputs."
'@Exposed
Option Explicit
Implements IValueValidator

Private SeparatorChar As String
Private UpdtSourceTrigger As MVVM.BindingUpdateSourceTrigger 'avoid being called twice

Private Sub Class_Initialize()
    SeparatorChar = VBA.Strings.Format$(0, ".")
End Sub

Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As IBindingPath, ByVal Target As IBindingPath) As Boolean
    Dim Result As Boolean
    Result = IsNumeric(Value)
    
    If Value = SeparatorChar Then
        
        Dim SourceValue As String
        If Source.TryReadPropertyValue(outValue:=SourceValue) Then
            'decimal separator character is legal if there aren't any already
            Dim TargetValue As String
            If Target.TryReadPropertyValue(outValue:=TargetValue) Then
                If SourceValue = 0 Or IsBindingTargetEmpty(Target) Then
                    TargetValue = vbNullString
                End If
                Dim Separators As Long
                Separators = Len(TargetValue) - Len(Replace(TargetValue, SeparatorChar, vbNullString))
                Result = Separators = 0
            End If
        End If
    End If
    
    IValueValidator_IsValid = Result
End Function

Private Function IsBindingTargetEmpty(ByVal Target As IBindingPath) As Boolean
    'TargetValue may be string-formatted; if all the content is selected, treat it as empty.
    On Error Resume Next
    'late-bound member calls should work against TextBox-like controls
    IsBindingTargetEmpty = Target.Context.SelLength = Target.Context.TextLength
    On Error GoTo 0
End Function

Private Property Get IValueValidator_Message() As String
    'IValueValidator_Message = "Value must be numeric."
    'VF: reworded to reflect different logic/trigger
    IValueValidator_Message = "Numeric value required."
End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    'VF: funnel through OnChange to capture cut/paste and backspace/delete
    'avoid being called twice
    If UpdtSourceTrigger = NotSetYet Then
        Select Case MsgBox("Trigger validation of numeric textboxes 'OnChange'?" & vbCr & vbCr & "No = Trigger 'OnKeyPress' (as originally, which means not capturing cut/paste and backspace/delete in the textboxes)", vbQuestion + vbYesNo, TypeName(Me))
            Case vbYes
                'but deletes entire string ...as of now
                'IValueValidator_Trigger = OnChange
                UpdtSourceTrigger = OnChange
            Case Else
                'IValueValidator_Trigger = OnKeyPress
                UpdtSourceTrigger = OnKeyPress
        End Select
    Else
        IValueValidator_Trigger = UpdtSourceTrigger
    End If
End Property
