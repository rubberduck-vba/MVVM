VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can resolve a string property path to a value."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "An object that can resolve a string property path to a value."
'@PredeclaredId
Option Explicit
Implements IBindingPath

Private Type TState
    Context As Object
    Path As String
    
    Object As Object
    PropertyName As String
End Type

Private This As TState

'@Description "Creates a new binding path from the specified property path string and binding context."
Public Function Create(ByVal Context As Object, ByVal Path As String) As IBindingPath
Attribute Create.VB_Description = "Creates a new binding path from the specified property path string and binding context."
    GuardClauses.GuardNonDefaultInstance Me, BindingPath, TypeName(Me)
    GuardClauses.GuardNullReference Context, TypeName(Me)
    GuardClauses.GuardEmptyString Path, TypeName(Me)
    
    Dim Result As BindingPath
    Set Result = New BindingPath
    Set Result.Context = Context
    Result.Path = Path
    
    Result.Resolve
    Set Create = Result
End Function

'@Description "Gets/Sets the binding context."
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets/Sets the binding context."
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As Object)
    GuardClauses.GuardDefaultInstance Me, BindingPath, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Context, TypeName(Me)
    Set This.Context = RHS
End Property

'@Description "Gets/Sets a string representing a property path against the binding context."
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets/Sets a string representing a property path against the binding context."
    Path = This.Path
End Property

Public Property Let Path(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, BindingPath, TypeName(Me)
    GuardClauses.GuardEmptyString RHS, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Path, TypeName(Me)
    This.Path = RHS
End Property

'@Description "Gets the bound object reference."
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets the bound object reference."
    Set Object = This.Object
End Property

'@Description "Gets the name of the bound property."
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property."
    PropertyName = This.PropertyName
End Property

'@Description "Resolves the Path to a bound object and property."
Public Sub Resolve()
Attribute Resolve.VB_Description = "Resolves the Path to a bound object and property."
    This.PropertyName = ResolvePropertyName(This.Path)
    Set This.Object = ResolvePropertyPath(This.Context, This.Path)
End Sub

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    ResolvePropertyName = Parts(UBound(Parts))
End Function

Private Function ResolvePropertyPath(ByVal Context As Object, ByVal PropertyPath As String) As Object
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    
    If UBound(Parts) = LBound(Parts) Then
        Set ResolvePropertyPath = Context
    Else
        Dim RecursiveProperty As Object
        Set RecursiveProperty = CallByName(Context, Parts(0), VbGet)
        If RecursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(RecursiveProperty, Right$(PropertyPath, Len(PropertyPath) - Len(Parts(0)) - 1))
    End If
    
End Function

Private Property Get IBindingPath_Context() As Object
    Set IBindingPath_Context = This.Context
End Property

Private Property Get IBindingPath_Path() As String
    IBindingPath_Path = This.Path
End Property

Private Property Get IBindingPath_Object() As Object
    Set IBindingPath_Object = This.Object
End Property

Private Property Get IBindingPath_PropertyName() As String
    IBindingPath_PropertyName = This.PropertyName
End Property

Private Sub IBindingPath_Resolve()
    Set This.Object = ResolvePropertyPath(This.Context, This.Path)
End Sub

Private Function IBindingPath_ToString() As String
    IBindingPath_ToString = StringBuilderNet _
        .AppendFormat("Context: {0}; Path: {1}", TypeName(This.Context), This.Path) _
        .ToString
End Function

Private Function IBindingPath_TryReadPropertyValue(ByRef outValue As Variant) As Boolean
    If This.Object Is Nothing Then Resolve
    On Error Resume Next
    outValue = VBA.Interaction.CallByName(This.Object, This.PropertyName, VbGet)
    IBindingPath_TryReadPropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function IBindingPath_TryWritePropertyValue(ByVal Value As Variant) As Boolean
    If This.Object Is Nothing Then Resolve
    On Error Resume Next
    VBA.Interaction.CallByName This.Object, This.PropertyName, VbLet, Value
    IBindingPath_TryWritePropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function
