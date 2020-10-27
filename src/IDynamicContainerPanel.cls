VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicContainerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes a panel layout component."
'@Exposed
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Describes a panel layout component."
'@Interface
Option Explicit

Public Enum LayoutDirection
    TopToBottom
    LeftToRight
    RightToLeft
End Enum

'@Description "Gets a value that controls how child controls are sized."
Public Property Get Direction() As LayoutDirection
Attribute Direction.VB_Description = "Gets a value that controls how child controls are sized."
End Property

'@Description "Gets the width of the panel."
Public Property Get Width() As Single
Attribute Width.VB_Description = "Gets the width of the panel."
End Property

'@Description "Gets the height of the panel."
Public Property Get Height() As Single
Attribute Height.VB_Description = "Gets the height of the panel."
End Property
