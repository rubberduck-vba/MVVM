VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContainerLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
'@Folder MVVM.Infrastructure.View.Dynamic
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IContainerLayout

Private Type TState
    Parent As MSForms.Controls
    Direction As LayoutDirection
    NextControlTop As Long
    NextControlLeft As Long
    Margin As Long
End Type

Private Const DefaultMargin As Long = 3

Private This As TState

Public Function Create(ByVal Parent As MSForms.Controls, Optional ByVal Direction As LayoutDirection, Optional ByVal Top As Long = 0, Optional ByVal Left As Long = 0, Optional ByVal Margin As Long = DefaultMargin) As IContainerLayout
    Dim Result As ContainerLayout
    Set Result = New ContainerLayout
    Set Result.Parent = Parent
    Result.NextControlTop = Top + Margin
    Result.NextControlLeft = Left + Margin
    Result.Direction = Direction
    Result.Margin = Margin
    Set Create = Result
End Function

Public Property Get Parent() As MSForms.Controls
    Set Parent = This.Parent
End Property

Public Property Set Parent(ByVal RHS As MSForms.Controls)
    Set This.Parent = RHS
End Property

Public Property Get NextControlTop() As Long
    NextControlTop = This.NextControlTop
End Property

Public Property Let NextControlTop(ByVal RHS As Long)
    This.NextControlTop = RHS
End Property

Public Property Get NextControlLeft() As Long
    NextControlLeft = This.NextControlLeft
End Property

Public Property Let NextControlLeft(ByVal RHS As Long)
    This.NextControlLeft = RHS
End Property

Public Property Get Direction() As LayoutDirection
    Direction = This.Direction
End Property

Public Property Let Direction(ByVal RHS As LayoutDirection)
    This.Direction = RHS
End Property

Public Property Get Margin() As Long
    Margin = This.Margin
End Property

Public Property Let Margin(ByVal RHS As Long)
    This.Margin = RHS
End Property

Private Sub OffsetControlSize(ByVal Control As MSForms.Control, ByVal Margin As Long)
    Select Case This.Direction
    
        Case LayoutDirection.TopToBottom
            This.NextControlTop = Control.Top + Control.Height + Margin
        
        Case LayoutDirection.LeftToRight
            This.NextControlLeft = Control.Left + Control.Width + Margin
            
        Case Else
            Debug.Assert False ' not implemented...
            
    End Select
End Sub

Private Function IContainerLayout_Add(ByVal ProgID As String, Optional ByVal Height As Variant, Optional ByVal Width As Variant, Optional ByVal ApplyMargin As Boolean = True) As Object
    Dim Result As Control
    Set Result = This.Parent.Add(ProgID)
    
    Result.Top = This.NextControlTop
    Result.Left = This.NextControlLeft
    
    If Not IsMissing(Height) Then
        Result.Height = Height
    End If
    
    If Not IsMissing(Width) Then
        Result.Width = Width
    End If
    
    OffsetControlSize Result, IIf(ApplyMargin, This.Margin, 0)
    Set IContainerLayout_Add = Result
    
End Function

Private Sub IContainerLayout_SkipPosition(Optional ByVal Margin As Long)
    Select Case This.Direction
        
        Case LayoutDirection.TopToBottom
            This.NextControlTop = This.NextControlTop + IIf(Margin > 0, Margin, This.Margin * 2)
            
        Case LayoutDirection.LeftToRight
            This.NextControlLeft = This.NextControlLeft + IIf(Margin > 0, Margin, This.Margin * 2)
            
    End Select
End Sub
