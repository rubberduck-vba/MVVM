VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotifierBaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A factory that creates PropertyChangeNotifierBase instances."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "A factory that creates PropertyChangeNotifierBase instances."
Option Explicit
Implements INotifierFactory

Private Function INotifierFactory_Create() As INotifyPropertyChanged
    Set INotifierFactory_Create = New PropertyChangeNotifierBase
End Function
