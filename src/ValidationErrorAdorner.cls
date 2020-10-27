VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target MSForms.Control object."
'@Folder MVVM.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "An object that dynamically decorates a target MSForms.Control object."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicAdorner
Implements IDisposable

Private Type TState
    Target As MSForms.Control
    TargetFormatter As MVVM.IValidationErrorFormatter
    
    AdornerLabel As MSForms.Label
    LabelFormatter As MVVM.IValidationErrorFormatter
    LabelControl As MSForms.Control
    
    AdornerIcon As MSForms.Image
    IconFormatter As MVVM.IValidationErrorFormatter
    IconControl As MSForms.Control
    
    LabelPosition As IDynamicPosition
    IconPosition As IDynamicPosition
    
    LabelPositionSet As Boolean
    IconPositionSet As Boolean
    Visible As Boolean
    
End Type

Private This As TState

Public Function Create(ByVal Target As MSForms.Control, _
Optional ByVal TargetFormatter As MVVM.IValidationErrorFormatter = Nothing, _
Optional ByVal AdornerLabel As Variant, _
Optional ByVal LabelFormatter As MVVM.IValidationErrorFormatter = Nothing, _
Optional ByVal LabelPosition As MVVM.IDynamicPosition = Nothing, _
Optional ByVal AdornerIcon As Variant, _
Optional ByVal IconFormatter As MVVM.IValidationErrorFormatter = Nothing, _
Optional ByVal IconPosition As MVVM.IDynamicPosition = Nothing) As IDynamicAdorner
    
    GuardClauses.GuardNonDefaultInstance Me, ValidationErrorAdorner
    GuardClauses.GuardNullReference Target, TypeName(Me)
    
    Dim Result As ValidationErrorAdorner
    Set Result = New ValidationErrorAdorner
    
    InitTarget Result, Target, TargetFormatter
    If IsMissing(AdornerLabel) Then
        Set AdornerLabel = Nothing
        InitLabel Result, AdornerLabel, LabelFormatter, LabelPosition
    ElseIf IsObject(AdornerLabel) Then
        If Not AdornerLabel Is Nothing Then
            InitLabel Result, AdornerLabel, LabelFormatter, LabelPosition
        End If
    End If
    
    If IsMissing(AdornerIcon) Then
        Set AdornerIcon = Nothing
        InitIcon Result, AdornerIcon, IconFormatter, IconPosition
    ElseIf IsObject(AdornerIcon) Then
        If Not AdornerIcon Is Nothing Then
            Set AdornerIcon = AdornerIcon
            InitIcon Result, AdornerIcon, IconFormatter, IconPosition
        End If
    End If
    
    Set Create = Result
    
End Function

Private Sub InitTarget(ByVal Result As ValidationErrorAdorner, ByVal Target As Object, ByVal Formatter As IValidationErrorFormatter)
    Set Result.Target = Target
    If Formatter Is Nothing Then
        Set Result.TargetFormatter = DefaultTargetFormatter
    Else
        Set Result.TargetFormatter = Formatter
    End If
End Sub

Private Sub InitLabel(ByVal Result As ValidationErrorAdorner, ByRef AdornerLabel As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef Position As IDynamicPosition)
    If Position Is Nothing Then Set Position = DefaultLabelPosition
    
    If AdornerLabel Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = Result.Target.Parent.Controls
        
        Set AdornerLabel = Parent.Add(MVVM.FormsProgID.LabelProgId, Visible:=False)
        
        AdornerLabel.BackStyle = fmBackStyleTransparent
        AdornerLabel.WordWrap = False
        AdornerLabel.AutoSize = True
        AdornerLabel.Font.Size = Result.Target.Parent.Font.Size - 2
                        
    End If
    
    PositionAdornerControlX Result.Target, AdornerLabel, Position
    PositionAdornerControlY Result.Target, AdornerLabel, Position
    
    Set Result.AdornerLabel = AdornerLabel
    Set Result.AdornerLabelPosition = Position
    If Formatter Is Nothing Then
        Set Result.AdornerLabelFormatter = DefaultLabelFormatter
    Else
        Set Result.AdornerLabelFormatter = Formatter
    End If
End Sub

Private Sub InitIcon(ByVal Result As ValidationErrorAdorner, ByRef AdornerIcon As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef Position As IDynamicPosition)
    If Position Is Nothing Then Set Position = DefaultIconPosition
    
    If AdornerIcon Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = Result.Target.Parent.Controls
        
        Set AdornerIcon = Parent.Add(MVVM.FormsProgID.ImageProgId, Visible:=False)
        Set AdornerIcon.Picture = Resources.ValidationErrorIcon.Picture
        
        AdornerIcon.BackStyle = fmBackStyleOpaque
        AdornerIcon.BorderStyle = fmBorderStyleSingle
        AdornerIcon.PictureSizeMode = fmPictureSizeModeClip
        AdornerIcon.PictureAlignment = fmPictureAlignmentCenter
        AdornerIcon.PictureTiling = False
        AdornerIcon.AutoSize = True
        
    End If
    
    PositionAdornerControlX Result.Target, AdornerIcon, Position
    PositionAdornerControlY Result.Target, AdornerIcon, Position
    
    Set Result.AdornerIcon = AdornerIcon
    Set Result.AdornerIconPosition = Position
    If Formatter Is Nothing And Not AdornerIcon Is Nothing Then
        Set Result.AdornerIconFormatter = DefaultIconFormatter
    Else
        Set Result.AdornerIconFormatter = Formatter
    End If
End Sub

Private Property Get DefaultTargetFormatter() As MVVM.IValidationErrorFormatter
    Set DefaultTargetFormatter = ValidationErrorFormatter _
        .WithErrorBackgroundColor _
        .WithErrorBoldFont _
        .WithErrorBorderColor _
        .WithErrorForeColor
End Property

Private Property Get DefaultLabelFormatter() As MVVM.IValidationErrorFormatter
    Set DefaultLabelFormatter = ValidationErrorFormatter _
        .WithErrorBoldFont _
        .WithErrorForeColor _
        .WithTargetOnlyVisibleOnError
End Property

Private Property Get DefaultLabelPosition() As MVVM.IDynamicPosition
    Set DefaultLabelPosition = DynamicControlPosition.Create(RelativePosition.BelowRight, 3, 0)
End Property

Private Property Get DefaultIconFormatter() As MVVM.IValidationErrorFormatter
    Set DefaultIconFormatter = ValidationErrorFormatter _
        .WithErrorBorderColor _
        .WithTargetOnlyVisibleOnError
End Property

Private Property Get DefaultIconPosition() As MVVM.IDynamicPosition
    Set DefaultIconPosition = DynamicControlPosition.Create(RelativePosition.InsideRight, 3)
End Property

Friend Property Get Target() As Object
    Set Target = This.Target
End Property

Friend Property Set Target(ByVal RHS As Object)
    GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    GuardClauses.GuardDoubleInitialization This.Target, TypeName(Me)
    Set This.Target = RHS
End Property

Friend Property Get TargetFormatter() As IValidationErrorFormatter
    Set TargetFormatter = This.TargetFormatter
End Property

Friend Property Set TargetFormatter(ByVal RHS As IValidationErrorFormatter)
    Set This.TargetFormatter = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = This.Visible
End Property

Private Sub PositionAdornerControlX(ByVal Target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal Position As IDynamicPosition)
    Select Case Position.Position
    
        Case RelativePosition.AboveRight, _
             RelativePosition.BelowRight, _
             RelativePosition.InsideRight, _
             RelativePosition.Default
            
            TrySetTextAlign Adorner, fmTextAlignRight
            Adorner.Left = Target.Left + Target.Width - Adorner.Width - Position.Margin(RightSide)
            
        Case RelativePosition.AboveLeft, _
             RelativePosition.BelowLeft, _
             RelativePosition.InsideLeft
            
            TrySetTextAlign Adorner, fmTextAlignLeft
            Adorner.Left = Target.Left + Position.Margin(LeftSide)
            
    End Select
End Sub

Private Sub TrySetTextAlign(ByVal Target As Object, ByVal Value As MSForms.fmTextAlign)
    'late-bound assignment will fail if Target doesn't have a TextAlign property.
    On Error Resume Next
    'if execution stops here, go Tools ~> Options ~> General ~> Error Trapping,
    'and make sure "Break on unhandled errors" is selected.
    'only use "Break on all errors" for debugging & diagnostics.
    Target.TextAlign = Value
    On Error GoTo 0
End Sub

Private Sub PositionAdornerControlY(ByVal Target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal Position As IDynamicPosition)
    Select Case Position.Position
    
        Case RelativePosition.AboveLeft, _
             RelativePosition.AboveRight
            
            Adorner.Top = Target.Top - Adorner.Height - Position.Margin(BottomSide)
            
        Case RelativePosition.InsideLeft, _
             RelativePosition.InsideRight
            
            Adorner.Top = Target.Top + Position.Margin(TopSide)
            
        Case RelativePosition.BelowLeft, _
             RelativePosition.BelowRight, _
             RelativePosition.Default
            
            Adorner.Top = Target.Top + Target.Height + Position.Margin(TopSide)
            
    End Select
End Sub

Friend Property Get AdornerLabel() As MSForms.Label
    Set AdornerLabel = This.AdornerLabel
End Property

Friend Property Set AdornerLabel(ByVal RHS As MSForms.Label)
    Set This.AdornerLabel = RHS
    Set This.LabelControl = RHS
End Property

Friend Property Get AdornerLabelFormatter() As IValidationErrorFormatter
    Set AdornerLabelFormatter = This.LabelFormatter
End Property

Friend Property Set AdornerLabelFormatter(ByVal RHS As IValidationErrorFormatter)
    Set This.LabelFormatter = RHS
End Property

Friend Property Get AdornerLabelPosition() As IDynamicPosition
    Set AdornerLabelPosition = This.LabelPosition
End Property

Friend Property Set AdornerLabelPosition(ByVal RHS As IDynamicPosition)
    Set This.LabelPosition = RHS
End Property

Friend Property Get AdornerIcon() As MSForms.Image
    Set AdornerIcon = This.AdornerIcon
End Property

Friend Property Set AdornerIcon(ByVal RHS As MSForms.Image)
    Set This.AdornerIcon = RHS
    Set This.IconControl = RHS
End Property

Friend Property Get AdornerIconFormatter() As IValidationErrorFormatter
    Set AdornerIconFormatter = This.IconFormatter
End Property

Friend Property Set AdornerIconFormatter(ByVal RHS As IValidationErrorFormatter)
    Set This.IconFormatter = RHS
End Property

Friend Property Get AdornerIconPosition() As IDynamicPosition
    Set AdornerIconPosition = This.IconPosition
End Property

Friend Property Set AdornerIconPosition(ByVal RHS As IDynamicPosition)
    Set This.IconPosition = RHS
End Property

Private Sub IDisposable_Dispose()
    Set This.Target = Nothing
    Set This.AdornerIcon = Nothing
    Set This.AdornerLabel = Nothing
    Set This.IconControl = Nothing
    Set This.LabelControl = Nothing
End Sub

Private Sub IDynamicAdorner_Hide()
    If Not This.Visible Then Exit Sub
    GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    GuardClauses.GuardNullReference This.TargetFormatter, TypeName(Me), "TargetFormatter is not set for this instance."
    
    This.TargetFormatter.Restore This.Target
    If Not This.IconControl Is Nothing Then This.IconFormatter.Restore This.IconControl
    If Not This.LabelControl Is Nothing Then This.LabelFormatter.Restore This.LabelControl

    This.Visible = False
End Sub

Private Sub IDynamicAdorner_Show(ByVal Message As String)
    'If This.Visible Then Exit Sub '<~ what if we just wanted to update the message?
    GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    GuardClauses.GuardNullReference This.TargetFormatter, TypeName(Me), "TargetFormatter is not set for this instance."
    
    This.TargetFormatter.Apply This.Target, Message
    If Not This.LabelControl Is Nothing Then ShowAdornerLabel Message
    If Not This.IconControl Is Nothing Then ShowAdornerIcon Message
    
    This.Visible = True
End Sub

Private Sub ShowAdornerLabel(ByVal Message As String)
    This.LabelFormatter.Apply This.LabelControl, Message
    If Not This.LabelPositionSet Then
        PositionAdornerControlX This.Target, This.AdornerLabel, This.LabelPosition
        PositionAdornerControlY This.Target, This.AdornerLabel, This.LabelPosition
        This.LabelPositionSet = True
    End If
End Sub

Private Sub ShowAdornerIcon(ByVal Message As String)
    This.IconFormatter.Apply This.IconControl, Message
    If Not This.IconPositionSet Then
        PositionAdornerControlX This.Target, This.AdornerIcon, This.IconPosition
        PositionAdornerControlY This.Target, This.AdornerIcon, This.IconPosition
        This.IconPositionSet = True
    End If
End Sub


