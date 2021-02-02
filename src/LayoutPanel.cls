VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LayoutPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder MVVM.Infrastructure.View.Dynamic
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TState
    Frame As MSForms.Frame
    Direction As LayoutDirection
    NextControlTop As Long
    NextControlLeft As Long
    Margin As Long
End Type

Private This As TState

Public Function Create(ByVal ParentForm As MSForms.UserForm, _
Optional ByVal Parent As MSForms.Controls, _
Optional ByVal Height As Single, _
Optional ByVal Width As Single) As LayoutPanel
    
    Dim Result As LayoutPanel
    Set Result = New LayoutPanel
    
    If Parent Is Nothing Then Set Parent = ParentForm.Controls
    
    Dim FrameControl As MSForms.Control
    Set FrameControl = Parent.Add(FormsProgID.FrameProgId, Visible:=False)
    If Height <> 0 Then
        FrameControl.Height = Height
    Else
        FrameControl.Height = ParentForm.Height '<~ RHS member call is late-bound
    End If
    If Width <> 0 Then
        FrameControl.Width = Width
    Else
        FrameControl.Width = ParentForm.Width '<~ RHS member call is late-bound
    End If
    
    Set Result.Frame = FrameControl
    Result.Frame.BorderStyle = fmBorderStyleNone
    Result.Frame.Caption = vbNullString
    Result.Frame.BackColor = ParentForm.BackColor
    Result.Frame.ScrollBars = fmScrollBarsBoth
    Result.Frame.KeepScrollBarsVisible = fmScrollBarsNone
    
    Set Create = Result
End Function

Public Property Get Frame() As MSForms.Frame
    Set Frame = This.Frame
End Property

Friend Property Set Frame(ByVal RHS As MSForms.Frame)
    Set This.Frame = RHS
End Property
