VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping command and property bindings in scope."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "An object keeping command and property bindings in scope."
'@Interface
'@Exposed
Option Explicit

'@Description "Gets an object responsible for managing command bindings."
Public Property Get Commands() As ICommandManager
Attribute Commands.VB_Description = "Gets an object responsible for managing command bindings."
End Property

'@Description "Gets an object responsible for managing property bindings."
Public Property Get Bindings() As MVVM.IBindingManager
Attribute Bindings.VB_Description = "Gets an object responsible for managing property bindings."
End Property

'@Description "Gets an object responsible for managing binding validation errors."
Public Property Get Validation() As MVVM.IValidationManager
Attribute Validation.VB_Description = "Gets an object responsible for managing binding validation errors."
End Property

'@Description "Gets an abstract factory that creates objects that can format strings."
Public Property Get StringFormatterFactory() As MVVM.IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets an abstract factory that creates objects that can format strings."
End Property
