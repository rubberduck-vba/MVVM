VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultErrorAdornerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder MVVM.Infrastructure.View.Dynamic
Option Explicit
Implements IDynamicAdornerFactory

Private Function IDynamicAdornerFactory_Create(ByVal Target As Object) As IDynamicAdorner
    Dim Result As IDynamicAdorner
    Select Case True
        
        Case TypeOf Target Is MSForms.TextBox
            Set Result = ValidationErrorAdorner.Create(Target, LabelPosition:=AdornerPosition.Create(AboveLeft, 0))
        
        Case TypeOf Target Is MSForms.ComboBox
            Set Result = ValidationErrorAdorner.Create(Target, IconPosition:=AdornerPosition.Create(InsideRight, 0, TopMargin:=3, RightMargin:=12))
        
        Case Else
            Debug.Print "A dynamic adorner was requested, but the target type isn't supported."
        
    End Select
    Set IDynamicAdornerFactory_Create = Result
End Function
