VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A collection of IValidationError objects for a given binding context (ViewModel)."
'@Folder MVVM.Infrastructure.Validation
'@ModuleDescription "A collection of IValidationError objects for a given binding context (ViewModel)."
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TState
    BindingContext As Object
    ValidationErrors As Collection
End Type
Private This As TState

'@Description "Creates a new collection of IValidationError objects for the specified binding context (ViewModel)."
Public Function Create(ByVal BindingContext As Object)
Attribute Create.VB_Description = "Creates a new collection of IValidationError objects for the specified binding context (ViewModel)."
    Dim Result As ValidationErrors
    Set Result = New ValidationErrors
    Set Result.BindingContext = BindingContext
    Set Create = Result
End Function

'@Description "Gets/sets the binding context (ViewModel) for this instance."
Public Property Get BindingContext() As Object
Attribute BindingContext.VB_Description = "Gets/sets the binding context (ViewModel) for this instance."
    Set BindingContext = This.BindingContext
End Property

Friend Property Set BindingContext(ByVal RHS As Object)
    Set This.BindingContext = RHS
End Property

'@DefaultMember
'@Description "Gets the IValidationError for the specified property; returns Nothing if there is no validation error."
Public Property Get Item(ByVal PropertyName As String) As IValidationError
Attribute Item.VB_Description = "Gets the IValidationError for the specified property; returns Nothing if there is no validation error."
Attribute Item.VB_UserMemId = 0
'QUESTION: 'VF: Can this be set via Ruberduck?   Attribute Item.VB_UserMemId = 0
'I thought this were a trick I would be telling you as I think you did not mention it anywhere on your blog,  well you know them all ;-)
'actually, this trick needs some explanations for other to understand how come this works
    Dim Result As IValidationError
    On Error Resume Next
    Set Result = This.ValidationErrors.Item(PropertyName)
    On Error GoTo 0 'key not found: returns Nothing
    Set Item = Result
End Property

'@Description "Gets the number of properties with a validation error in this binding context."
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of properties with a validation error in this binding context."
    Count = This.ValidationErrors.Count
End Property

'@Enumerator
'@Description "Provides custom collection support for For Each enumeration."
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Provides custom collection support for For Each enumeration."
Attribute NewEnum.VB_UserMemId = -4
'QUESTION: 'VF: Can this be set via Ruberduck?
'Attribute NewEnum.VB_Description = "Provides custom collection support for For Each enumeration."
'Attribute NewEnum.VB_UserMemId = -4
'I thought this were a trick I would be telling you as I think you did not mention it anywhere on your blog,  well you know them all ;-)
'actually, this trick needs some explanations for other to understand how come this works
'regarding collections
'http://dailydoseofexcel.com/archives/2010/07/04/custom-collection-class/
'Attribute NewEnum.VB_MemberFlags = "40" 'for each/next in collection
    Set NewEnum = This.ValidationErrors.[_NewEnum]
End Property

'@Description "Adds (or replaces) a validation error."
Public Sub Add(ByVal ValidationError As IValidationError)
Attribute Add.VB_Description = "Adds (or replaces) a validation error."
    
    Dim PropertyName As String
    PropertyName = ValidationError.Binding.Source.PropertyName
    
    Remove PropertyName
    This.ValidationErrors.Add ValidationError, PropertyName
    
End Sub

'@Description "Removes validation error for the specified property if it exists. True if successful, False if nothing was done."
Public Function Remove(ByVal PropertyName As String) As Boolean
Attribute Remove.VB_Description = "Removes validation error for the specified property if it exists. True if successful, False if nothing was done."
    Dim Existing As IValidationError
    Set Existing = Item(PropertyName)
    If Not Existing Is Nothing Then
        This.ValidationErrors.Remove PropertyName
        Remove = True
    End If
End Function

Private Sub Class_Initialize()
    Set This.ValidationErrors = New Collection
End Sub

Private Sub Class_Terminate()
    Set This.ValidationErrors = Nothing
End Sub
