VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Holds the validation errors state for a ViewModel."
'@Folder MVVM.Infrastructure.Validation
'@ModuleDescription "Holds the validation errors state for a ViewModel."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationManager

Private Type TState
    ErrorAdornerFactory As IDynamicAdornerFactory
    ValidationErrors As Dictionary
    IsValid As Boolean
End Type

Private This As TState

Public Function Create(Optional ByVal ErrorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    Dim Result As ValidationManager
    Set Result = New ValidationManager
    InitAdornerFactory Result, ErrorAdornerFactory
    Set Create = Result
End Function

Private Sub InitAdornerFactory(ByVal Manager As MVVM.ValidationManager, ByVal AdornerFactory As MVVM.IDynamicAdornerFactory)
    Dim Factory As MVVM.IDynamicAdornerFactory
    
    If AdornerFactory Is Nothing Then
        Dim DefaultFactory As MVVM.IDynamicAdornerFactory
        Set DefaultFactory = New DefaultErrorAdornerFactory
        Set Factory = DefaultFactory
    Else
        'use injected instance (could be a test stub):
        Set Factory = AdornerFactory
        
    End If
    
    Set Manager.ErrorAdornerFactory = Factory
End Sub

Public Property Get ErrorAdornerFactory() As IDynamicAdornerFactory
    Set ErrorAdornerFactory = This.ErrorAdornerFactory
End Property

Friend Property Set ErrorAdornerFactory(ByVal RHS As IDynamicAdornerFactory)
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    GuardClauses.GuardNullReference RHS, TypeName(ValidationManager)
    GuardClauses.GuardDoubleInitialization This.ErrorAdornerFactory, TypeName(Me)
    Set This.ErrorAdornerFactory = RHS
End Property

Public Property Get IsValid(ByVal BindingContext As Object, Optional ByVal PropertyName As String) As Boolean
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    
    Dim Result As Boolean
    Result = True
    
    Dim Key As Variant
    For Each Key In This.ValidationErrors.Keys
        
        Dim Errors As ValidationErrors
        Set Errors = This.ValidationErrors.Item(Key)
        
        If BindingContext Is Errors.BindingContext Or Errors.BindingContext Is Nothing Then
            
            If PropertyName = vbNullString Then
                Result = Errors.Count = 0
            Else
                Dim Error As IValidationError
                For Each Error In Errors
                    Result = (Error.Binding.Source.PropertyName <> PropertyName)
                    If Not Result Then Exit For
                Next
            End If
            
        End If
        
        If Not Result Then Exit For
    Next
    
    IsValid = Result
    
End Property

Public Property Get ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Property
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    
    'ValidationErrors.Item getter returns Nothing if there's no error for the given property:
    Set ValidationError = Errors.Item(BindingPath.PropertyName)
End Property

Private Sub ClearValidationErrors(ByVal BindingPath As IBindingPath)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Sub
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    Errors.Remove BindingPath.PropertyName
End Sub

Private Sub OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(ValidationError.Binding.Source.Context))
    
    Dim Errors As ValidationErrors
    If This.ValidationErrors.Exists(ContextKey) Then
        Set Errors = This.ValidationErrors.Item(ContextKey)
    Else
        Set Errors = ValidationErrors.Create(Context)
    End If
    
    Errors.Add ValidationError
    Set This.ValidationErrors.Item(ContextKey) = Errors
End Sub

Private Sub Class_Initialize()
    Set This.ValidationErrors = New Dictionary
    This.IsValid = True
End Sub

Private Sub Class_Terminate()
    Set This.ValidationErrors = Nothing
End Sub

Private Property Get IValidationManager_AdornerFactory() As IDynamicAdornerFactory
    Set IValidationManager_AdornerFactory = This.ErrorAdornerFactory
End Property

Private Sub IValidationManager_ClearValidationError(ByVal BindingPath As IBindingPath)
    ClearValidationErrors BindingPath
End Sub

Private Property Get IValidationManager_IsValid(ByVal Context As Object, Optional ByVal PropertyName As String) As Boolean
    IValidationManager_IsValid = IsValid(Context, PropertyName)
End Property

Private Sub IValidationManager_OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    OnValidationError Context, ValidationError
End Sub

Private Property Get IValidationManager_ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Set IValidationManager_ValidationError = ValidationError(BindingPath)
End Property
