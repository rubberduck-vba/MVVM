# MVVM in VBA!

Thanks to [Rubberduck](https://github.com/rubberduck-vba/Rubberduck), object-oriented programming in VBA is easier than ever; large projects with dozens of class modules can now be neatly organized in a custom folder hierarchy, for example.

This project demonstrates that not only OOP but also *Model-View-ViewModel* can be leveraged in VBA, mainly for educational and inspirational purposes.

# MVVM Infrastructure Overview

**Model-View-ViewModel** is an object-oriented UI design pattern that makes it easier to write decoupled, modular applications a user needs to interact with. The main components are the *model*, the *view*, and the *view model*.

Creating an MVVM application in VBA begins with an `AppContext` object instance, which provides the MVVM infrastructure API and gives you `BindingManager`, `CommandManager`, and `ValidationManager` objects. See the [API] worksheet for documentation and example usage.

---

## Model

The **Model** is your application's data and the objects responsible for retrieving it.
It might consist of multiple classes, including services abstracting database operations, for example.
It's also "data transfer objects" (DTO) that carry data from a source: classes with nothing but read/write properties (or public fields).

This project provides no out-of-the-box infrastructure code for the model component, because the model is inherently application-specific.

## View

The **View** is how your application communicates with its user, and how the user communicates with the application.
The role of the view is to present the **ViewModel** to the user, being strictly concerned with presentation.

The **View** is responsible for everything directly related to the user interface (UI), and should implement the **IView** interface.

![IView interface](https://user-images.githubusercontent.com/5751684/97098041-f8102580-164e-11eb-884b-85d6348b6cac.png)

### Dependencies
In order to function properly with MVVM, the **View** should hold an instance-level reference to the `IAppContext` interface.
Additionally, in order to configure the property bindings the **View** will require a reference to a specific **ViewModel** object/type.

### Remarks
Modal dialogs that can be cancelled should implement the `ICancellable` interface.
The code-behind module contains member calls against the `IAppContext` object model, to configure property and command bindings for the **View**.

## ViewModel

![INotifyPropertyChanged interface](https://user-images.githubusercontent.com/5751684/97098071-458c9280-164f-11eb-98f6-2483a25f1e0a.png)

The **ViewModel** is how the magic happens. View models should implement the `INotifyPropertyChanged` interface.
Moreover, `Property Let` procedures should only conditionally assign the backing instance state, and invoke the `OnPropertyChanged` method:

```vb
Public Property Let StringProperty(ByVal RHS As String)
    If This.StringProperty <> RHS Then
        This.StringProperty = RHS
        OnPropertyChanged "StringProperty"
    End If
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Me, PropertyName
End Sub
```

### Dependencies

In order to facilitate and standardize implementing `INotifyPropertyChanged`, the MVVM infrastructure provides the `PropertyChangeNotifierBase` class. The object reference for this dependency can be assigned in the `Class_Initialize` handler procedure:

```vb
Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifierBase
End Sub
```

A **ViewModel** may also expose multiple `ICommand` properties that the **View** can use to configure command bindings.
The MVVM infrastructure API provides `AcceptCommand` and `CancelCommand` standard commands; application-specific commands should be implemented as needed.
