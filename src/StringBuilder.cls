VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "StringBuilder implementation modified from https://codereview.stackexchange.com/q/196076"
'@Folder MVVM.Infrastructure.StringFormatting
'@ModuleDescription "StringBuilder implementation modified from https://codereview.stackexchange.com/q/196076"
'>> modified identifier names to better fit this project
'>> modified to prevent a stateful default instance
'@IgnoreModule ParameterCanBeByVal: string parameters are passed ByRef to avoid copying a string of unknown length, for performance reasons.
'@PredeclaredId
'@Exposed
Option Explicit
'Based off vba_stringbuilder:Class - http://blog.livedoor.jp/midorityo/archives/50744656.html
Private Type TState
    AppendsBufferSize As Long
    Builder As Object
    Index As Long
End Type
Private AppendsBuffer() As String
Private This As TState

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Const DEFAULT_APPENDSBUFFER_SIZE As Long = 10000
        Set This.Builder = CreateObject("System.Text.StringBuilder")
        SetAppendsBufferSize DEFAULT_APPENDSBUFFER_SIZE
        This.Index = -1
    End If
End Sub

Public Property Get Capacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Capacity = This.Builder.Capacity
End Property

Private Property Let Capacity(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    This.Builder.Capacity = RHS
End Property

Private Property Get Length() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Length = This.Builder.Length
End Property

Public Property Let Length(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    This.Builder.Length = RHS
End Property

Public Function MaxCapacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    MaxCapacity = This.Builder.MaxCapacity
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is StringBuilder
End Property

Public Function Append(ByRef Text As String, Optional RepeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Append = Result.Append(Text, RepeatCount)
        Exit Function
    End If
    
    This.Index = This.Index + 1
    AppendsBuffer(This.Index) = Text
    If RepeatCount > 0 Then Result.Append Text, (RepeatCount - 1)
    If This.Index = This.AppendsBufferSize Then Compact
    Set Append = Me
End Function

Public Function AppendFormat(ByVal Format As Variant, ParamArray Args() As Variant) As StringBuilder
    Dim LocalArgs() As Variant
    If IsArray(Args(LBound(Args))) Then
        LocalArgs = Args(LBound(Args))
    Else
        ReDim LocalArgs(LBound(Args) To UBound(Args))
        Dim Index As Long
        For Index = LBound(Args) To UBound(Args)
            LocalArgs(Index) = Args(Index)
        Next
    End If
    
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set AppendFormat = Result.AppendFormat(Format, LocalArgs)
        Exit Function
    End If
    
    Compact

    If Format = vbNullString Then
        Set AppendFormat = Me.Append(VBA.Strings.Join(LocalArgs))
        Exit Function
    End If

    This.Builder.AppendFormat_4 Format, LocalArgs
    Set AppendFormat = Me
End Function

Public Function InsertFormat(ByVal Format As Variant, StartIndex As Long, RepeatCount As Long, ParamArray Args() As Variant) As StringBuilder
    Dim LocalArgs() As Variant
    If IsArray(Args(LBound(Args))) Then
        LocalArgs = Args(LBound(Args))
    Else
        ReDim LocalArgs(LBound(Args) To UBound(Args))
        Dim Index As Long
        For Index = LBound(Args) To UBound(Args)
            LocalArgs(Index) = Args(Index)
        Next
    End If
    
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set InsertFormat = Result.InsertFormat(Format, StartIndex, RepeatCount, LocalArgs)
        Exit Function
    End If
    
    If Format = vbNullString Then
        Set InsertFormat = Me.Insert(StartIndex, VBA.Strings.Join(LocalArgs))
        Exit Function
    End If
    
    With New StringBuilder
        .AppendFormat Format, LocalArgs
        Me.Insert StartIndex, .ToString
    End With
    Set InsertFormat = Me
End Function

Public Function AppendLine(Optional ByRef Text As String, Optional ByVal RepeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set AppendLine = Result.AppendLine(Text, RepeatCount)
        Exit Function
    End If
    
    This.Index = This.Index + 1
    AppendsBuffer(This.Index) = Text & vbNewLine
    If RepeatCount > 0 Then Me.AppendLine Text, (RepeatCount - 1)
    If This.Index = This.AppendsBufferSize Then Compact
    Set AppendLine = Me
End Function

Public Sub Clear()
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Me.Length = 0
End Sub

Private Sub Compact()
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    If This.Index > -1 Then
        This.Builder.Append_3 Join(AppendsBuffer, vbNullString)
        This.Index = -1
    End If
    ReDim AppendsBuffer(This.AppendsBufferSize)
End Sub

Public Function EnsureCapacity(ByVal Length As Long) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set EnsureCapacity = Result.EnsureCapacity(Length)
        Exit Function
    End If
    This.Builder.EnsureCapacity Length
    Set EnsureCapacity = Me
End Function

'@Ignore ParameterCanBeByVal: Text is passed ByRef to avoid copying a string of unknown length, for performance reasons.
Public Function Insert(ByVal StartIndex As Long, ByRef Text As String, Optional ByVal RepeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Insert = Result.Insert(StartIndex, Text, RepeatCount)
        Exit Function
    End If
    Compact
    This.Builder.Insert_2 StartIndex, Text
    If RepeatCount > 0 Then Me.Insert StartIndex, Text, (RepeatCount - 1)
    Set Insert = Me
End Function

Public Function Remove(ByVal StartIndex As Long, ByVal Length As Long) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Remove = Result.Remove(StartIndex, Length)
        Exit Function
    End If
    Compact
    This.Builder.Remove StartIndex, Length
    Set Remove = Me
End Function

Public Function Replace(ByRef OldValue As String, ByRef NewValue As String, Optional ByVal StartIndex As Long = -1, Optional ByVal RepeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Replace = Result.Replace(OldValue, NewValue, StartIndex, RepeatCount)
        Exit Function
    End If
    
    Compact

    If StartIndex > -1 And RepeatCount > 0 Then
        This.Builder.Replace_2 OldValue, NewValue, StartIndex, RepeatCount
    Else
        This.Builder.Replace OldValue, NewValue
    End If
    Set Replace = Me
End Function

Private Sub SetAppendsBufferSize(ByVal Length As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    This.AppendsBufferSize = Length
    Compact
End Sub

Public Function ToString(Optional ByVal StartIndex As Long = -1, Optional ByVal Length As Long = 0) As String
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Compact

    If StartIndex > -1 And Length > 0 Then
        ToString = This.Builder.ToString_2(StartIndex, Length)
    Else
        ToString = This.Builder.ToString
    End If
End Function
