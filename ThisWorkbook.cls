VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Always show the persistent modeless launcher and sync it to the current sheet
    On Error Resume Next
    modWorkflowTables.ShowLauncher
    modWorkflowTables.SyncLauncherToSheet ActiveSheet
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' If the launcher was closed (X) or hidden, re-show it and sync to the active sheet
    On Error Resume Next
    modWorkflowTables.ShowLauncher
    modWorkflowTables.SyncLauncherToSheet Sh
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Tidy up the modeless form on close
    On Error Resume Next
    modWorkflowTables.HideLauncher True   ' True = unload the form
    On Error GoTo 0
End Sub
