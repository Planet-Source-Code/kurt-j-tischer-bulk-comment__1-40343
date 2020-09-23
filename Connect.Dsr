VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8625
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11565
   _ExtentX        =   20399
   _ExtentY        =   15214
   _Version        =   393216
   Description     =   "Adds comment to code window"
   DisplayName     =   "Bulk Cmoment"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed As Boolean
Public VBInstance As VBIDE.VBE
Dim mcbMenuCommandBar1 As Office.CommandBarControl
Dim mcbMenuCommandBar2 As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1


'command bar event handlers
Public WithEvents AddCommentHandler As CommandBarEvents
Attribute AddCommentHandler.VB_VarHelpID = -1
Public WithEvents RemoveCommentHandler As CommandBarEvents
Attribute RemoveCommentHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application

    'create additional links here
    Set mcbMenuCommandBar1 = AddToAnotherCommandBar("Add Comments", "Code Window")
    Set Me.AddCommentHandler = Application.Events.CommandBarEvents(mcbMenuCommandBar1)

    Set mcbMenuCommandBar2 = AddToAnotherCommandBar("Remove Comments", "Code Window")
    Set Me.RemoveCommentHandler = Application.Events.CommandBarEvents(mcbMenuCommandBar2)
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    Resume Next
    
End Sub

Function AddToAnotherCommandBar(sCaption As String, Optional sBarName As String = "Add-Ins") As Office.CommandBarControl

    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars(sBarName)
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAnotherCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
    MsgBox Err.Description

End Function



'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar1.Delete
    mcbMenuCommandBar2.Delete
    
End Sub

Private Sub AddCommentHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    BulkComment False

End Sub

Public Sub BulkComment(Optional bRemove As Boolean = False)
'============================================================
'Routine:   BulkComment
'Purpose:   Adds/removes comments to/from selected lines in code pane
'Author:    Kurt J. Tischer
'Created:   5/12/1998  Copyright © 1998-2002 3rd Ear Productions
'
'Notes:
'
'
On Error Resume Next
'============================================================

Dim sl As Long
Dim el As Long
Dim sc As Long
Dim ec As Long
Dim i As Long
Dim sLine As String

'get the current selection
VBInstance.ActiveCodePane.GetSelection sl, sc, el, ec

If Not bRemove Then
    'add comment
    If sl > 0 Then ' make sure we have at least one line selected
        For i = sl To el - 1
            sLine = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
            sLine = "'" & sLine
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, sLine
        Next i
    End If
Else
    For i = sl To el - 1 ' make sure we have at least one line selected
        sLine = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
        If Left(sLine, 1) = "'" Then
            sLine = Right$(sLine, Len(sLine) - 1)
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, sLine
        End If
    Next i
End If

End Sub
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function






Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

End Sub

Private Sub RemoveCommentHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    BulkComment True

End Sub


