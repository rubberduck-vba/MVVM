# MVVM in VBA!

With [Rubberduck](https://github.com/rubberduck-vba/Rubberduck), object-oriented programming in VBA is easier than ever: large projects with many small and specialized class modules can be neatly organized in a custom folder hierarchy, for one.

This project demonstrates that not only OOP but also *Model-View-ViewModel* can be leveraged in VBA, mainly for educational and inspirational purposes. 

## Features

The 100+ modules solve many problems related to building and programming user interfaces in VBA, and provide an object model that gives an application a solid, decoupled backbone structure.

### Object Model

The `IAppContext` interface, and its `AppContext` implementation, are at the top of the MVVM object model. This *context* object exposes `IBindingManager`, `ICommandManager`, and `IValidationManager` objects (among others), each holding their own piece of the application's state (property bindings, command bindings, and binding validation errors, respectively).

### Property Bindings

The `INotifyPropertyChanged` interface allows property bindings to work both from the source (ViewModel) to the target (UI controls), and from the target to the source. Hence, by implementing this interface on ViewModel classes, UI code can bind a ViewModel property to a `MSForms.TextBox` control (or anything), via the `IBindingManager.BindPropertyPath` method - by letting the manager infer most of everything...

```vba
With Context.Bindings 'where Context is an IAppContext object reference
    ' use IBindingManager.BindPropertyPath to bind a ViewModel property to a property of a MSForms control target.
    .BindPropertyPath ViewModel, "Instructions", Me.InstructionsLabel
End With
```

...or by configuring every aspect of the binding explicitly.

### Validation

Application code may implement the `IValueValidator` interface to supply a property binding with a `Validator` argument. Bindings that fail validation use the default *dynamic error adorner* (that was configured when the top-level `AppContext` is created) to display configurable visual indicators (border, background, font colors, but also dynamic tooltips, icons, and labels); when the binding is valid again, the visual cues are hidden and the `IValidationManager` holds no more `IValidationError` objects in its `ValidationErrors` collection for the ViewModel's binding context (each ViewModel gets its own "validation scope").

By default, an invalid field visually looks like this:

![an invalid string property binding with the default dynamic adorner shown](https://user-images.githubusercontent.com/5751684/97099459-ac19ac80-165f-11eb-9430-7fda96dc4d8b.png)


### Command Bindings

The `ICommand` interface can be implemented for anything that needs to happen in response to the user clicking a button: in MVVM you don't handle `Click` events anymore, instead you *bind* an implementation of the `ICommand` interface to a `MSForms.CommandButton` control: the MVVM infrastructure code automatically takes care to enable or disable that control (you provide the `ICommand.CanExecute` Boolean logic, MVVM automatically invokes it).

```vba
With Context.Commands 'where Context is an IAppContext object reference
    ' use ICommandManager.BindCommand to bind a MSForms.CommandButton to any ICommand object.
    .BindCommand ViewModel, Me.CommandButton1, ViewModel.SomeCommand
End With
```

### Dynamic UI

This part of the API is still very much subject to breaking changes since it's very much alpha-stage, but the idea is to provide an API to make it easy to programmatically *generate* a user interface from VBA code, and automatically create the associated property and command bindings.

Whether your UI is dynamic or made at design-time, the recommendation would be to create the bindings in a dedicated `InitializeView` procedure in the form's code-behind.

This example snippet is from the `ExampleDynamicView` module - remember to invoke `IBindingManager.Apply` to bring it all to life:

```vba
Private Sub InitializeView()
    
    Dim Layout As IContainerLayout
    Set Layout = ContainerLayout.Create(Me.Controls, TopToBottom)
    
    With DynamicControls.Create(This.Context, Layout)
        
        With .LabelFor("All controls on this form are created at run-time.")
            .Font.Bold = True
        End With
        
        .LabelFor BindingPath.Create(This.ViewModel, "Instructions")
        
        .TextBoxFor BindingPath.Create(This.ViewModel, "StringProperty"), _
                    Validator:=New RequiredStringValidator, _
                    TitleSource:="Some String:"
                    
        .TextBoxFor BindingPath.Create(This.ViewModel, "CurrencyProperty"), _
                    FormatString:="{0:C2}", _
                    Validator:=New DecimalKeyValidator, _
                    TitleSource:="Some Amount:"
        
        .CommandButtonFor AcceptCommand.Create(Me, This.Context.Validation), This.ViewModel, "Close"
        
    End With
    
    This.Context.Bindings.Apply This.ViewModel
End Sub
```
