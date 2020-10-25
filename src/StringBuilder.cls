VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "StringBuilder implementation from https://codereview.stackexchange.com/q/196076"
'@Folder Framework.Strings
'@ModuleDescription "StringBuilder implementation from https://codereview.stackexchange.com/q/196076"
'@PredeclaredId
'@Exposed
Option Explicit
'Based off vba_stringbuilder:Class - http://blog.livedoor.jp/midorityo/archives/50744656.html
Private Type Members
    AppendsBufferSize As Long
    Builder As Object
    Index As Long
End Type
Private AppendsBuffer() As String
Private m As Members

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Const DEFAULT_APPENDSBUFFER_SIZE As Long = 10000
        Set m.Builder = CreateObject("System.Text.StringBuilder")
        setAppendsBufferSize DEFAULT_APPENDSBUFFER_SIZE
        m.Index = -1
    End If
End Sub

Public Property Get Capacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Capacity = m.Builder.Capacity
End Property

Public Property Let Capacity(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    m.Builder.Capacity = RHS
End Property

Public Property Get Length() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Length = m.Builder.Length
End Property

Public Property Let Length(ByVal RHS As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    m.Builder.Length = RHS
End Property

Public Function MaxCapacity() As Long
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    MaxCapacity = m.Builder.MaxCapacity
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is StringBuilder
End Property

Public Function Append(ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Append = Result.Append(Text, repeatCount)
        Exit Function
    End If
    
    m.Index = m.Index + 1
    AppendsBuffer(m.Index) = Text
    If repeatCount > 0 Then Result.Append Text, (repeatCount - 1)
    If m.Index = m.AppendsBufferSize Then Compact
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
    
    Me.Compact

    If Format = vbNullString Then
        Set AppendFormat = Me.Append(VBA.Strings.Join(LocalArgs))
        Exit Function
    End If

    m.Builder.AppendFormat_4 Format, LocalArgs
    Set AppendFormat = Me
End Function

Public Function InsertFormat(ByVal Format As Variant, startIndex As Long, repeatCount As Long, ParamArray Args() As Variant) As StringBuilder
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
        Set InsertFormat = Result.InsertFormat(Format, startIndex, repeatCount, LocalArgs)
        Exit Function
    End If
    
    If Format = vbNullString Then
        Set InsertFormat = Me.Insert(startIndex, VBA.Strings.Join(LocalArgs))
        Exit Function
    End If
    
    With New StringBuilder
        .AppendFormat Format, LocalArgs
        Me.Insert startIndex, .ToString
    End With
    Set InsertFormat = Me
End Function

Public Function AppendLine(Optional ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set AppendLine = Result.AppendLine(Text, repeatCount)
        Exit Function
    End If
    
    m.Index = m.Index + 1
    AppendsBuffer(m.Index) = Text & vbNewLine
    If repeatCount > 0 Then Me.AppendLine Text, (repeatCount - 1)
    If m.Index = m.AppendsBufferSize Then Compact
    Set AppendLine = Me
End Function

Public Sub Clear()
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Me.Length = 0
End Sub

Public Sub Compact()
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    If m.Index > -1 Then
        m.Builder.Append_3 Join(AppendsBuffer, vbNullString)
        m.Index = -1
    End If
    ReDim AppendsBuffer(m.AppendsBufferSize)
End Sub

Public Function EnsureCapacity(ByVal Length As Long) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set EnsureCapacity = Result.EnsureCapacity(Length)
        Exit Function
    End If
    m.Builder.EnsureCapacity Length
    Set EnsureCapacity = Me
End Function

Public Function Insert(startIndex As Long, Text As String, Optional repeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Insert = Result.Insert(startIndex, Text, repeatCount)
        Exit Function
    End If
    Me.Compact
    m.Builder.Insert_2 startIndex, Text
    If repeatCount > 0 Then Me.Insert startIndex, Text, (repeatCount - 1)
    Set Insert = Me
End Function

Public Function Remove(ByVal startIndex As Long, ByVal Length As Long) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Remove = Result.Remove(startIndex, Length)
        Exit Function
    End If
    Me.Compact
    m.Builder.Remove startIndex, Length
    Set Remove = Me
End Function

Public Function Replace(oldValue As String, newValue As String, Optional startIndex As Long = -1, Optional repeatCount As Long = 0) As StringBuilder
    If IsDefaultInstance Then
        Dim Result As StringBuilder
        Set Result = New StringBuilder
        Set Replace = Result.Replace(oldValue, newValue, startIndex, repeatCount)
        Exit Function
    End If
    
    Me.Compact

    If startIndex > -1 And repeatCount > 0 Then
        m.Builder.Replace_2 oldValue, newValue, startIndex, repeatCount
    Else
        m.Builder.Replace oldValue, newValue
    End If
    Set Replace = Me
End Function

Public Sub setAppendsBufferSize(ByVal Length As Long)
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    m.AppendsBufferSize = Length
    Me.Compact
End Sub

Public Function ToString(Optional ByVal startIndex As Long = -1, Optional ByVal Length As Long = 0) As String
    GuardClauses.GuardDefaultInstance Me, StringBuilder
    Me.Compact

    If startIndex > -1 And Length > 0 Then
        ToString = m.Builder.ToString_2(startIndex, Length)
    Else
        ToString = m.Builder.ToString
    End If
End Function
