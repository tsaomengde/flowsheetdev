VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Launcher 
   Caption         =   "Workflow Launcher"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UF_Launcher.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UF_Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.caption = "Workflow Launcher"
    Me.cmdAddPhone.caption = "Add Phone Number"
    Me.lblTitle.caption = "Launcher"
    Me.lblStatus.caption = "Loading..."
    On Error GoTo 0
End Sub

' Optional: ensure the button is enabled only when ready (Workbook will sync it)
Private Sub UserForm_Activate()
    On Error Resume Next
    ' Defer to SyncLauncherToSheet to set caption/status/enabled state
    If TypeOf Application.ActiveSheet Is Worksheet Then
        modWorkflowTables.SyncLauncherToSheet Application.ActiveSheet
    End If
End Sub

Private Sub cmdAddPhone_Click()
    ' Directly start the workflow for the active sheet
    modWorkflowTables.StartAddPhone_FromLauncher
End Sub
