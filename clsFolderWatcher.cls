VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents items As Outlook.items
Attribute items.VB_VarHelpID = -1
Public FolderStoreID As String
Public FolderEntryID As String
Public FolderPath As String
Public isUnified As Boolean

Private Sub Items_ItemAdd(ByVal Item As Object)
    On Error Resume Next

    If isUnified Then Exit Sub
    If gIsDeleteSyncing Then Exit Sub

    If TypeOf Item Is Outlook.MailItem Then
        Debug.Print "ItemAdd fired in source inbox: " & FolderPath
        HandleSourceItemAdd Me, Item
    End If
End Sub

Private Sub Items_ItemRemove()
    On Error Resume Next

    Debug.Print "ItemRemove fired in: " & FolderPath
    HandleFolderRemove Me
End Sub
