VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
'@ModuleDescription "An object that describes how a control changes its appearance given a validation error."
'@Folder MVVM.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "An object that describes how a control changes its appearance given a validation error. Handles MSForms and Excel.Range targets." '...which means the class is doing way too many things for its own good.
'@PredeclaredId
'@Exposed
'@IgnoreModule ValueRequiredInspection false positive, stdole.OLE_COLOR
Option Explicit
Implements IValidationErrorFormatter

Private Const DefaultErrorBackColor As Long = &HC0C0FF
Private Const DefaultErrorBorderColor As Long = &HC0
Private Const DefaultErrorBorderWidth As Long = 2
Private Const DefaultErrorForeColor As Long = &HC0

Private Type TState
    Applied As Boolean
    
    FormatBackgroundColor As Boolean
    FormatBorderColor As Boolean
    FormatForeColor As Boolean
    FormatFontBold As Boolean
    
    InitialBackgroundColor As Long
    InitialBorderColor As Long
    InitialBorderWidth As Double
    InitialForeColor As Long
    InitialFontBold As Boolean
    
    ErrorBackgroundColor As Long
    ErrorBorderColor As Long
    ErrorBorderWidth As Double
    ErrorForeColor As Long
    ErrorFontBold As Boolean
    
    TargetOnlyVisibleOnError As Boolean
    PropertyName As String
    
End Type

Private This As TState

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ValidationErrorFormatter
End Property

'@Description "A builder method to create or configure a formatter that sets the target's background color when applied."
Public Function WithErrorBackgroundColor(Optional ByVal Color As Variant = DefaultErrorBackColor) As ValidationErrorFormatter
Attribute WithErrorBackgroundColor.VB_Description = "A builder method to create or configure a formatter that sets the target's background color when applied."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithErrorBackgroundColor(Color)
    Else
        Set Result = Me
        This.ErrorBackgroundColor = Color
        This.FormatBackgroundColor = True
    End If
    Set WithErrorBackgroundColor = Result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's border color when applied."
Public Function WithErrorBorderColor(Optional ByVal Color As Variant = DefaultErrorBorderColor) As ValidationErrorFormatter
Attribute WithErrorBorderColor.VB_Description = "A builder method to create or configure a formatter that sets the target's border color when applied."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithErrorBorderColor(Color)
    Else
        Set Result = Me
        This.ErrorBorderColor = Color
        This.FormatBorderColor = True
    End If
    Set WithErrorBorderColor = Result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's border weight when applied (Excel.Range targets only)."
Public Function WithErrorBorderWidth(Optional ByVal Width As Double = DefaultErrorBorderWidth) As ValidationErrorFormatter
Attribute WithErrorBorderWidth.VB_Description = "A builder method to create or configure a formatter that sets the target's border weight when applied (Excel.Range targets only)."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithErrorBorderWidth(Width)
    Else
        Set Result = Me
        This.ErrorBorderWidth = Width
        This.FormatBorderColor = True
    End If
    Set WithErrorBorderWidth = Result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied."
Public Function WithErrorForeColor(Optional ByVal Color As Variant = DefaultErrorForeColor) As ValidationErrorFormatter
Attribute WithErrorForeColor.VB_Description = "A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithErrorForeColor(Color)
    Else
        Set Result = Me
        This.ErrorForeColor = Color
        This.FormatForeColor = True
    End If
    Set WithErrorForeColor = Result
End Function

'@Description "A builder method to create or configure a formatter that makes the target's font bold when applied."
Public Function WithErrorBoldFont() As ValidationErrorFormatter
Attribute WithErrorBoldFont.VB_Description = "A builder method to create or configure a formatter that makes the target's font bold when applied."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithErrorBoldFont
    Else
        Set Result = Me
        This.FormatFontBold = True
        This.ErrorFontBold = True
    End If
    Set WithErrorBoldFont = Result
End Function

'@Description "A builder method to create or configure a formatter that makes the target only visible when formatter is applied."
Public Function WithTargetOnlyVisibleOnError() As ValidationErrorFormatter
Attribute WithTargetOnlyVisibleOnError.VB_Description = "A builder method to create or configure a formatter that makes the target only visible when formatter is applied."
    Dim Result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set Result = New ValidationErrorFormatter
        Set Result = Result.WithTargetOnlyVisibleOnError
    Else
        Set Result = Me
        This.TargetOnlyVisibleOnError = True
    End If
    Set WithTargetOnlyVisibleOnError = Result
End Function

Private Sub StoreInitialFormat(ByVal RHS As Object)
    GuardClauses.GuardDefaultInstance Me, ValidationErrorFormatter, TypeName(Me)
    
    Select Case True
    
        Case TypeOf RHS Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = RHS
            This.InitialBackgroundColor = TextBoxTarget.BackColor
            This.InitialBorderColor = TextBoxTarget.BorderColor
            This.InitialFontBold = TextBoxTarget.Font.Bold
            This.InitialForeColor = TextBoxTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = RHS
            This.InitialFontBold = CheckBoxTarget.Font.Bold
            This.InitialForeColor = CheckBoxTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = RHS
            This.InitialBackgroundColor = LabelTarget.BackColor
            This.InitialBorderColor = LabelTarget.BorderColor
            This.InitialFontBold = LabelTarget.Font.Bold
            This.InitialForeColor = LabelTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = RHS
            This.InitialBackgroundColor = FrameTarget.BackColor
            This.InitialBorderColor = FrameTarget.BorderColor
            This.InitialFontBold = FrameTarget.Font.Bold
            This.InitialForeColor = FrameTarget.ForeColor
        
        Case TypeOf RHS Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = RHS
            This.InitialBackgroundColor = ImageTarget.BackColor
            This.InitialBorderColor = ImageTarget.BorderColor
            
        Case TypeOf RHS Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = RHS
            This.InitialBackgroundColor = RangeTarget.Interior.Color
            This.InitialBorderColor = RangeTarget.Borders.Color
            This.InitialBorderWidth = RangeTarget.Borders.Weight
            This.InitialFontBold = RangeTarget.Font.Bold
            This.InitialForeColor = RangeTarget.Font.Color
            
        Case Else
            Debug.Print TypeName(Me) & ": Target of type " & TypeName(RHS) & " is not currently supported."
            
    End Select
    
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal Target As Object, ByVal Message As String)

    If This.Applied Then Exit Sub
    StoreInitialFormat Target
    
    If TypeOf Target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = Target
        If Not This.TargetOnlyVisibleOnError Then
            Control.ControlTipText = Message
        End If
    End If
    
    Select Case True
    
        Case TypeOf Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = Target
            If This.FormatBackgroundColor Then TextBoxTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then TextBoxTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then TextBoxTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then TextBoxTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = Target
            If This.FormatFontBold Then CheckBoxTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then CheckBoxTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = Target
            If This.FormatBackgroundColor Then LabelTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then LabelTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then LabelTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then LabelTarget.ForeColor = This.ErrorForeColor
            If This.TargetOnlyVisibleOnError Then LabelTarget.Caption = Message
            This.Applied = True
            
        Case TypeOf Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = Target
            If This.FormatBackgroundColor Then FrameTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then FrameTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then FrameTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then FrameTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = Target
            If This.FormatBackgroundColor Then ImageTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then ImageTarget.BorderColor = This.ErrorBorderColor
            This.Applied = True
            
        Case TypeOf Target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = Target
            If This.FormatBackgroundColor Then RangeTarget.Interior.Color = This.ErrorBackgroundColor
            If This.FormatBorderColor Then
                RangeTarget.Borders.Color = This.ErrorBorderColor
                RangeTarget.Borders.Weight = This.ErrorBorderWidth
            End If
            If This.FormatFontBold Then RangeTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then RangeTarget.Font.Color = This.ErrorForeColor
    
    End Select
    
    If This.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = This.Applied
    End If

End Sub


Private Sub IValidationErrorFormatter_Restore(ByVal Target As Object)
    If Not This.Applied Then Exit Sub
    
    If TypeOf Target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = Target
        Control.ControlTipText = vbNullString
    End If
    
    Select Case True
    
        Case TypeOf Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = Target
            If This.FormatBackgroundColor Then TextBoxTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then TextBoxTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then TextBoxTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then TextBoxTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = Target
            If This.FormatFontBold Then CheckBoxTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then CheckBoxTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = Target
            If This.FormatBackgroundColor Then LabelTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then LabelTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then LabelTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then LabelTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = Target
            If This.FormatBackgroundColor Then FrameTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then FrameTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then FrameTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then FrameTarget.ForeColor = This.InitialForeColor
        
            This.Applied = False
            
        Case TypeOf Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = Target
            If This.FormatBackgroundColor Then ImageTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then ImageTarget.BorderColor = This.InitialBorderColor
            
            This.Applied = False
            
        Case TypeOf Target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = Target
            If This.FormatBackgroundColor Then RangeTarget.Interior.Color = This.InitialBackgroundColor
            If This.FormatBorderColor Then
                RangeTarget.Borders.Color = This.InitialBorderColor
                RangeTarget.Borders.Weight = This.InitialBorderWidth
            End If
            If This.FormatFontBold Then RangeTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then RangeTarget.Font.Color = This.InitialForeColor
            
    End Select

    If This.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = This.Applied
    End If

End Sub
