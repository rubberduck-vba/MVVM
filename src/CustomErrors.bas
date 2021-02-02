Attribute VB_Name = "CustomErrors"
Attribute VB_Description = "Global, general-purpose procedures involving run-time errors."
'@Folder MVVM.CustomErrors
'@ModuleDescription("Global, general-purpose procedures involving run-time errors.")
Option Explicit
Option Private Module

Public Const CustomError As Long = vbObjectError Or 32 'QUESTION: VF: why this value?

'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Information.Err
        If .Number <> 0 Then
            Debug.Print "Error " & .Number, .Description
            .Raise .Number
        End If
    End With
End Sub
