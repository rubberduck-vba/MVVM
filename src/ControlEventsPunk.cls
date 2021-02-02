VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEventsPunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides an event sink to relay MSForms.Control events."
'@IgnoreModule IntegerDataType
'@Folder MVVM.Infrastructure.Win32
'@ModuleDescription "Provides an event sink to relay MSForms.Control events."
'based on https://stackoverflow.com/a/51936950
'based on https://stackoverflow.com/a/61893857
Option Explicit
Implements IControlEvents

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'[This function is available through Windows XP and Windows Server 2003. It might be altered or unavailable in subsequent versions of Windows.]
'https://docs.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-connecttoconnectionpoint
#If VBA7 Then
Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal Punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal PunkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
#Else
Private Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal Punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal PunkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
#End If

Private Type TState
    RefIID As GUID 'The IID of the interface on the connection point container whose connection point object is being requested.
    Connected As Boolean
    PunkTarget As Object
    Cookie As Long
    
    Handlers As Collection
End Type

Private This As TState

'@Description "Gets/sets the target MSForms.Control reference."
Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets/sets the target MSForms.Control reference."
    Set Target = This.PunkTarget
End Property

Public Property Set Target(ByVal RHS As Object)
    Set This.PunkTarget = RHS
End Property

'@Description "Registers the listener."
Public Function Connect() As Boolean
Attribute Connect.VB_Description = "Registers the listener."
    GuardClauses.GuardNullReference This.PunkTarget, TypeName(Me), "Target is not set."
    ConnectToConnectionPoint Me, This.RefIID, True, This.PunkTarget, This.Cookie, 0&
    This.Connected = This.Cookie <> 0
    Connect = This.Connected
End Function

'@Description "De-registers the listener."
Public Function Disconnect() As Boolean
Attribute Disconnect.VB_Description = "De-registers the listener."
    If Not This.Connected Then Exit Function
    ConnectToConnectionPoint Me, This.RefIID, False, This.PunkTarget, This.Cookie, 0&
    This.Connected = False
    Disconnect = True
End Function

'@Description "A callback that handles MSForms.Control.AfterUpdate events for the registered target control."
Public Sub OnAfterUpdate()
Attribute OnAfterUpdate.VB_Description = "A callback that handles MSForms.Control.AfterUpdate events for the registered target control."
Attribute OnAfterUpdate.VB_UserMemId = -2147384832
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleAfterUpdate
    Next
End Sub

'@Description "A callback that handles MSForms.Control.BeforeUpdate events for the registered target control."
Public Sub OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
Attribute OnBeforeUpdate.VB_Description = "A callback that handles MSForms.Control.BeforeUpdate events for the registered target control."
Attribute OnBeforeUpdate.VB_UserMemId = -2147384831
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        
        Handler.HandleBeforeUpdate LocalCancel
        Cancel.Value = LocalCancel
        If LocalCancel Then Exit For
    Next
End Sub

'@Description "A callback that handles MSForms.Control.Exit events for the registered target control."
Public Sub OnExit(ByVal Cancel As MSForms.IReturnBoolean)
Attribute OnExit.VB_Description = "A callback that handles MSForms.Control.Exit events for the registered target control."
Attribute OnExit.VB_UserMemId = -2147384829
    Dim LocalCancel As Boolean
    LocalCancel = Cancel.Value
    
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleExit LocalCancel
        Cancel.Value = LocalCancel
        If LocalCancel Then Exit For
    Next
End Sub

'@Description "A callback that handles MSForms.Control.Enter events for the registered target control."
Public Sub OnEnter()
Attribute OnEnter.VB_Description = "A callback that handles MSForms.Control.Enter events for the registered target control."
Attribute OnEnter.VB_UserMemId = -2147384830
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleEnter
    Next
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox, replaced OnKeyPress in TextBox
'@Description "A callback that handles MSForms.Control.Change events for the registered target control."
Public Sub OnChange()
    Dim Handler As IHandleControlEvents
    For Each Handler In This.Handlers
        Handler.HandleChange
    Next
End Sub

'@Description "Registers the specified object to handle the relayed control events."
Public Sub RegisterHandler(ByVal Handler As IHandleControlEvents)
Attribute RegisterHandler.VB_Description = "Registers the specified object to handle the relayed control events."
    This.Handlers.Add Handler
End Sub

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
    This.RefIID.Data1 = &H20400
    This.RefIID.Data4(0) = &HC0
    This.RefIID.Data4(7) = &H46
End Sub

Private Sub Class_Terminate()
    If This.Cookie <> 0 Then Disconnect
End Sub

Private Sub IControlEvents_OnAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IControlEvents_OnBeforeUpdate(ByVal Cancel As MSForms.IReturnBoolean)
    OnBeforeUpdate Cancel
End Sub
'VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox control, basically replacing OnKeyPress in TextBox
Private Sub IControlEvents_OnChange()
    OnChange
End Sub

Private Sub IControlEvents_OnEnter()
    OnEnter
End Sub

Private Sub IControlEvents_OnExit(ByVal Cancel As MSForms.IReturnBoolean)
    OnExit Cancel
End Sub

Private Sub IControlEvents_RegisterHandler(ByVal Handler As IHandleControlEvents)
    RegisterHandler Handler
End Sub
