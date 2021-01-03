VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExploreTextboxEvents 
   Caption         =   "UserForm1"
   ClientHeight    =   1404
   ClientLeft      =   -48
   ClientTop       =   -192
   ClientWidth     =   744
   OleObjectBlob   =   "ExploreTextboxEvents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExploreTextboxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#VF_CanBeDeleted
'@Description: "'VF: userform used for exploring textbox events to capture cut and paste to identify suitable validation trigger; turned out to be _Change"
Option Explicit

Private Sub TextBox1_AfterUpdate()
Label1.Caption = Label1.Caption & vbLf & "TextBox1_AfterUpdate"
End Sub

Private Sub TextBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_BeforeDropOrPaste"
End Sub

Private Sub TextBox1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_BeforeUpdate"
End Sub

'VF: implement OnChange for Textbox   paste/delete/cut easiest captured by _change, alternatively fiddle with KeyCodes
Private Sub TextBox1_Change()
    Label1.Caption = Label1.Caption & vbLf & "TextBox1_Change"
End Sub

Private Sub TextBox1_Enter()
Label1.Caption = Label1.Caption & vbLf & "TextBox1_Enter"
End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_Exit"
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_KeyDown"
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_KeyPress"
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Label1.Caption = Label1.Caption & vbLf & "TextBox1_KeyUp"
End Sub

