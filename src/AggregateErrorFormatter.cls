VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AggregateErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A composite error formatter implementation."
'@Folder MVVM.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "A composite error formatter implementation."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationErrorFormatter

Private Const AutoPropertyName As String = "AUTO"

Private Type TState
    Context As Object
    PropertyName As String
    Formatters As Collection
End Type

Private This As TState

'@Description "Creates a new composite formatter that aggregates/combines all the specified formatters."
Public Function Create(ParamArray Formatters() As Variant) As IValidationErrorFormatter
Attribute Create.VB_Description = "Creates a new composite formatter that aggregates/combines all the specified formatters."
    GuardClauses.GuardNonDefaultInstance Me, AggregateErrorFormatter, TypeName(Me)
    GuardClauses.GuardExpression UBound(Formatters) < 0, TypeName(Me), "At least one formatter is required."
    
    Dim Result As AggregateErrorFormatter
    Set Result = New AggregateErrorFormatter
    
    'handle ParamArray proxy args:
    Dim LocalFormatters() As Variant
    If IsArray(Formatters(0)) Then
        LocalFormatters = Formatters(0)
    Else
        ReDim LocalFormatters(UBound(Formatters))
        Dim j As Long
        For j = 0 To UBound(LocalFormatters)
            LocalFormatters(j) = Formatters(j)
        Next
    End If
    
    Dim i As Long
    For i = LBound(LocalFormatters) To UBound(LocalFormatters)
        
        If Not IsEmpty(LocalFormatters(i)) And Not LocalFormatters(i) Is Nothing Then
            Dim Formatter As IValidationErrorFormatter
            On Error Resume Next
            Set Formatter = LocalFormatters(i)
            On Error GoTo 0
            If Not Formatter Is Nothing Then Result.AddFormatter Formatter
        End If
        
    Next
    
    Set Create = Result
    
End Function

'@Description "Adds the specified formatter to this composite formatter."
Public Sub AddFormatter(ByVal Formatter As IValidationErrorFormatter)
Attribute AddFormatter.VB_Description = "Adds the specified formatter to this composite formatter."
    GuardClauses.GuardDefaultInstance Me, AggregateErrorFormatter, TypeName(Me)
    GuardClauses.GuardNullReference Formatter
    
    This.Formatters.Add Formatter

End Sub

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is AggregateErrorFormatter
End Property

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then Set This.Formatters = New Collection
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal Target As Object, ByVal Message As String)
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In This.Formatters
        Formatter.Apply Target, Message
    Next
End Sub

Private Sub IValidationErrorFormatter_Restore(ByVal Target As Object)
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In This.Formatters
        Formatter.Restore Target
    Next
End Sub
