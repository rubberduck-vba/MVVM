VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing the validation errors in a ViewModel."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Represents an object responsible for managing the validation errors in a ViewModel."
'@Interface
'@Exposed
Option Explicit

'@Description "True if there are no validation errors, False otherwise."
Public Property Get IsValid(ByVal Context As Object, Optional ByVal PropertyName As String) As Boolean
Attribute IsValid.VB_Description = "True if there are no validation errors, False otherwise."
End Property

'@Description "Gets the IValidationError for the specified binding path, if it exists."
Public Property Get ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
Attribute ValidationError.VB_Description = "Gets the IValidationError for the specified binding path, if it exists."
End Property

'@Description "Gets the factory used for creating the default validation error adorners."
Public Property Get AdornerFactory() As IDynamicAdornerFactory
Attribute AdornerFactory.VB_Description = "Gets the factory used for creating the default validation error adorners."
End Property

'@Description "Adds the specified validation error to the validation context."
Public Sub OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
Attribute OnValidationError.VB_Description = "Adds the specified validation error to the validation context."
End Sub

'@Description "Removes the validation error for the specified binding path, if it exists."
Public Sub ClearValidationError(ByVal BindingPath As IBindingPath)
Attribute ClearValidationError.VB_Description = "Removes the validation error for the specified binding path, if it exists."
End Sub
