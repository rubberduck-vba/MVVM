VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can resolve a string property path to a value."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object that can resolve a string property path to a value."
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the context for resolving the Path."
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets the context for resolving the Path."
End Property

'@Description "Gets the property path of the binding."
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets the property path of the binding."
End Property

'@Description "Gets a reference to the binding source object resolved from the Path."
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets a reference to the binding source object resolved from the Path."
End Property

'@Description "Gets the name of the bound property resolved from the Path."
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property resolved from the Path."
End Property

'@Description "Resolves the binding source."
Public Sub Resolve()
Attribute Resolve.VB_Description = "Resolves the binding source."
End Sub

'@Description "Reads and returns the property value; False on error, True if successful."
Public Function TryReadPropertyValue(ByRef outValue As Variant) As Boolean
Attribute TryReadPropertyValue.VB_Description = "Reads and returns the property value; False on error, True if successful."
End Function

'@Description "Writes the value to the bound property; False on error, True if successful."
Public Function TryWritePropertyValue(ByVal Value As Variant) As Boolean
Attribute TryWritePropertyValue.VB_Description = "Writes the value to the bound property; False on error, True if successful."
End Function

'@Description "Returns a string representing this binding path for debugging purposes."
Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string representing this binding path for debugging purposes."
End Function
