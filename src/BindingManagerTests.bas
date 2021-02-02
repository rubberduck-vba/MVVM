Attribute VB_Name = "BindingManagerTests"
'@Folder Tests
'@TestModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
Private Assert As Object
#Else
Private Assert As Rubberduck.AssertClass
#End If

Private Type TState
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As BindingManager
    AbstractSUT As IBindingManager
    HandlePropertyChangedSUT As IHandlePropertyChanged
    
    CommandManager As ICommandManager
    CommandManagerStub As ITestStub
    
    BindingSource As TestBindingObject
    BindingTarget As TestBindingObject
    
    SourceProperty As String
    TargetProperty As String
    
    SourcePropertyPath As String
    TargetPropertyPath As String
    
    Command As TestCommand
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
#If LateBind Then
    'requires HKCU registration of the Rubberduck COM library.
    Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
#Else
    'requires project reference to the Rubberduck COM library.
    Set Assert = New Rubberduck.PermissiveAssertClass
#End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Test.CommandManager = New TestCommandManager
    Set Test.CommandManagerStub = Test.CommandManager
    Set Test.ConcreteSUT = BindingManager.Create(Test.CommandManager, New StringFormatterNetFactory)
    Set Test.AbstractSUT = Test.ConcreteSUT
    Set Test.HandlePropertyChangedSUT = Test.ConcreteSUT
    Set Test.BindingSource = New TestBindingObject
    Set Test.BindingTarget = New TestBindingObject
    Set Test.Command = New TestCommand
    Test.SourcePropertyPath = "TestStringProperty"
    Test.TargetPropertyPath = "TestStringProperty"
    Test.SourceProperty = "TestStringProperty"
    Test.TargetProperty = "TestStringProperty"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
    Set Test.AbstractSUT = Nothing
    Set Test.HandlePropertyChangedSUT = Nothing
    Set Test.BindingSource = Nothing
    Set Test.BindingTarget = Nothing
    Set Test.Command = Nothing
    Test.SourcePropertyPath = vbNullString
    Test.TargetPropertyPath = vbNullString
    Test.ExpectedErrNumber = 0
    Test.ExpectedErrorCaught = False
    Test.ExpectedErrSource = vbNullString
End Sub

Private Sub ExpectError()
    Dim Message As String
    If Err.Number = Test.ExpectedErrNumber Then
        If (Test.ExpectedErrSource = vbNullString) Or (Err.Source = Test.ExpectedErrSource) Then
            Test.ExpectedErrorCaught = True
        Else
            Message = "An error was raised, but not from the expected source. " & _
                      "Expected: '" & TypeName(Test.ConcreteSUT) & "'; Actual: '" & Err.Source & "'."
        End If
    ElseIf Err.Number <> 0 Then
        Message = "An error was raised, but not with the expected number. Expected: '" & Test.ExpectedErrNumber & "'; Actual: '" & Err.Number & "'."
    Else
        Message = "No error was raised."
    End If
    
    If Not Test.ExpectedErrorCaught Then Assert.Fail Message
End Sub

'@TestMethod("GuardClauses")
Private Sub Create_GuardsNonDefaultInstance()
    Test.ExpectedErrNumber = GuardClauseErrors.InvalidFromNonDefaultInstance
    With New BindingManager
        On Error Resume Next
            '@Ignore FunctionReturnValueDiscarded, FunctionReturnValueNotUsed
            .Create Test.CommandManager, New StringFormatterNetFactory
            ExpectError
        On Error GoTo 0
    End With
End Sub

Private Function DefaultPropertyPathBindingFor(ByVal ProgID As String, ByRef outTarget As Object) As IPropertyBinding
    Set outTarget = CreateObject(ProgID)
    Set DefaultPropertyPathBindingFor = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, Test.SourcePropertyPath, outTarget, Test.TargetPropertyPath)
End Function

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_CheckBoxTargetCreatesCheckBoxPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.CheckBoxProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is CheckBoxPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_CheckBoxTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.CheckBoxProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ComboBoxTargetCreatesComboBoxPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ComboBoxProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is ComboBoxPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ComboBoxTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ComboBoxProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ListBoxTargetCreatesListBoxPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ListBoxProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is ListBoxPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ListBoxTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ListBoxProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_MultiPageTargetCreatesMultiPagePropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.MultiPageProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is MultiPagePropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_MultiPageTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.MultiPageProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName, "Actual: " & Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_OptionButtonTargetCreatesOptionButtonPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.OptionButtonProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is OptionButtonPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_OptionButtonTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.OptionButtonProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ScrollBarTargetCreatesScrollBarPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ScrollBarProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is ScrollBarPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_ScrollBarTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.ScrollBarProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_SpinButtonTargetCreatesSpinButtonPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.SpinButtonProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is SpinButtonPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_SpinButtonTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.SpinButtonProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_TabStripTargetCreatesTabStripPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.TabStripProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is TabStripPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_TabStripTargetBindsValueByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.TabStripProgId, outTarget:=Target)
    Assert.AreEqual "Value", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_TextBoxTargetCreatesTextBoxPropertyBinding()
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.TextBoxProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is TextBoxPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_TextBoxTargetBindsTextPropertyByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.TextBoxProgId, outTarget:=Target)
    Assert.AreEqual "Text", Result.Target.PropertyName
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_FrameTargetCreatesOneWayBindingWithNonDefaultTarget()
    Test.TargetPropertyPath = "Font.Bold"
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.FrameProgId, outTarget:=Target)
    Assert.AreEqual TypeName(OneWayPropertyBinding), TypeName(Result), "Actual: " & TypeName(Result)
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_LabelTargetCreatesOneWayBindingWithNonDefaultTarget()
    Test.TargetPropertyPath = "Font.Bold"
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.LabelProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is OneWayPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_FrameTargetBindsCaptionPropertyByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.FrameProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is CaptionPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_LabelTargetBindsCaptionPropertyByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim Result As IPropertyBinding
    Set Result = DefaultPropertyPathBindingFor(FormsProgID.LabelProgId, outTarget:=Target)
    Assert.IsTrue TypeOf Result Is CaptionPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_NonControlTargetCreatesOneWayBinding()
    Dim Result As IPropertyBinding
    Set Result = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, Test.SourcePropertyPath, Test.BindingTarget, Test.TargetPropertyPath)
    Assert.IsTrue TypeOf Result Is OneWayPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_NonControlTargetRequiresTargetPropertyPath()
    Test.ExpectedErrNumber = GuardClauseErrors.StringCannotBeEmpty
    On Error Resume Next
        Test.AbstractSUT.BindPropertyPath _
            Test.BindingSource, _
            Test.SourcePropertyPath, _
            Test.BindingTarget, _
            TargetProperty:=vbNullString
        ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("CallbackPropagation")
Private Sub BindPropertyPath_AddsToPropertyBindingsCollection()
    Dim Result As IPropertyBinding
    Set Result = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, Test.SourcePropertyPath, Test.BindingTarget, Test.TargetPropertyPath)
    Assert.AreEqual 1, Test.ConcreteSUT.PropertyBindings.Count
    Assert.AreSame Result, Test.ConcreteSUT.PropertyBindings.Item(1)
End Sub

'@TestMethod("CallbackPropagation")
Private Sub HandlePropertyChanged_EvaluatesCommandCanExecute()
    Test.HandlePropertyChangedSUT.HandlePropertyChanged Test.BindingSource, Test.SourceProperty
    Test.CommandManagerStub.Verify Assert, "EvaluateCanExecute", 1
End Sub

'@TestMethod("CallbackPropagation")
Private Sub HandlePropertyChanged_EvaluatesCommandCanExecuteForAnyPropertyChange()
    Test.HandlePropertyChangedSUT.HandlePropertyChanged Test.BindingSource, Test.SourceProperty
    Test.HandlePropertyChangedSUT.HandlePropertyChanged Test.BindingSource, "Not" & Test.SourceProperty
    Test.CommandManagerStub.Verify Assert, "EvaluateCanExecute", 2
End Sub


