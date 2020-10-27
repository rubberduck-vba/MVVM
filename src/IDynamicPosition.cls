VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents the relative position and margin of a dynamic UI component."
'@Folder MVVM.Infrastructure.Abstract
'@ModuleDescription "Represents the relative position and margin of a dynamic UI component."
'@Interface
'@Exposed
Option Explicit

Public Enum RelativePosition
    Default = 0
    AboveLeft
    AboveRight
    BelowLeft
    BelowRight
    InsideLeft
    InsideRight
End Enum

Public Enum MarginSide
    LeftSide = 1
    TopSide = 2
    RightSide = 4
    BottomSide = 8
End Enum

'@Description "Gets the position of the adorner relative to its target."
Public Property Get Position() As RelativePosition
Attribute Position.VB_Description = "Gets the position of the adorner relative to its target."
End Property

'@Description "Gets the margin value for the specified side."
Public Property Get Margin(ByVal Side As MarginSide) As Double
Attribute Margin.VB_Description = "Gets the margin value for the specified side."
End Property
