Attribute VB_Name = "modUnified"
Option Explicit

' =========================
' GLOBALS
' =========================

Public FolderWatchers As Collection
Public FolderSnapshots As Object ' Dictionary(folderKey -> Dictionary(itemEntryID -> packed links))
Public gIsDeleteSyncing As Boolean

Private Const APP_NAME As String = "UnifiedInbox"
Private Const APP_SECTION As String = "Config"

Private Const PROP_SOURCE_ENTRYID As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_SourceEntryID"
Private Const PROP_SOURCE_STOREID As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_SourceStoreID"
Private Const PROP_SOURCE_FOLDER As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_SourceFolderPath"

Private Const PROP_UNIFIED_ENTRYID As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_CopyEntryID"
Private Const PROP_UNIFIED_STOREID As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_CopyStoreID"
Private Const PROP_UNIFIED_FOLDER As String = "http://schemas.microsoft.com/mapi/string/{5C2D5A1F-9F4B-4C4B-9C45-8E1D3D8A1001}/Unified_CopyFolderPath"

' =========================
' STARTUP
' =========================

Public Sub ManualStartup()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.store
    Dim inbox As Outlook.folder
    Dim unified As Outlook.folder
    Dim watcher As clsFolderWatcher

    Set ns = Application.Session
    Set FolderWatchers = New Collection
    Set FolderSnapshots = CreateObject("Scripting.Dictionary")

    ' Watch every real inbox
    For Each store In ns.Stores
        Set inbox = Nothing

        On Error Resume Next
        Set inbox = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0

        If Not inbox Is Nothing Then
            Set watcher = New clsFolderWatcher
            Set watcher.items = inbox.items
            watcher.FolderStoreID = inbox.storeID
            watcher.FolderEntryID = inbox.entryID
            watcher.FolderPath = inbox.FolderPath
            watcher.isUnified = False
            FolderWatchers.Add watcher

            Debug.Print "Watching source inbox: " & inbox.FolderPath
        End If
    Next store

    ' Watch unified folder
    Set unified = GetUnifiedFolder()
    If Not unified Is Nothing Then
        Set watcher = New clsFolderWatcher
        Set watcher.items = unified.items
        watcher.FolderStoreID = unified.storeID
        watcher.FolderEntryID = unified.entryID
        watcher.FolderPath = unified.FolderPath
        watcher.isUnified = True
        FolderWatchers.Add watcher

        Debug.Print "Watching unified folder: " & unified.FolderPath
    Else
        Debug.Print "No unified folder selected."
    End If

    RefreshAllSnapshots
    Debug.Print "Startup complete. Watchers: " & FolderWatchers.Count
End Sub

' =========================
' MAIN EVENT HANDLERS
' =========================

Public Sub HandleSourceItemAdd(ByVal watcher As clsFolderWatcher, ByVal Item As Outlook.MailItem)
    Dim unified As Outlook.folder
    Dim sourceFolder As Outlook.folder
    Dim tmpCopy As Outlook.MailItem
    Dim unifiedCopy As Outlook.MailItem
    Dim movedObj As Object
    Dim existingUnified As Object
    Dim existingUnifiedID As String
    Dim existingUnifiedStoreID As String
    Dim srcLive As Object

    On Error GoTo CleanFail

    Set unified = GetUnifiedFolder()
    If unified Is Nothing Then Exit Sub

    Set sourceFolder = Item.Parent

    ' Skip if this somehow is already in unified
    If sourceFolder.entryID = unified.entryID Then Exit Sub

    ' If source already linked to a live unified copy, skip
    existingUnifiedID = GetUserPropValue(Item, PROP_UNIFIED_ENTRYID)
    existingUnifiedStoreID = GetUserPropValue(Item, PROP_UNIFIED_STOREID)

    If Len(existingUnifiedID) > 0 And Len(existingUnifiedStoreID) > 0 Then
        On Error Resume Next
        Set existingUnified = Application.Session.GetItemFromID(existingUnifiedID, existingUnifiedStoreID)
        On Error GoTo CleanFail

        If Not existingUnified Is Nothing Then
            Debug.Print "Source already linked, skipping copy: " & Item.Subject
            Exit Sub
        End If
    End If

    Debug.Print "Copying new source item to unified: " & Item.Subject

    Set tmpCopy = Item.Copy
    If tmpCopy Is Nothing Then
        Debug.Print "Copy failed."
        Exit Sub
    End If

    ' Stamp the COPY with source references BEFORE moving it
    AddOrSetUserProp tmpCopy, PROP_SOURCE_ENTRYID, Item.entryID
    AddOrSetUserProp tmpCopy, PROP_SOURCE_STOREID, sourceFolder.storeID
    AddOrSetUserProp tmpCopy, PROP_SOURCE_FOLDER, sourceFolder.FolderPath

    tmpCopy.Save

    Set movedObj = tmpCopy.Move(unified)
    If movedObj Is Nothing Then
        Debug.Print "Move to unified failed."
        Exit Sub
    End If

    If Not TypeOf movedObj Is Outlook.MailItem Then
        Debug.Print "Moved object is not a MailItem."
        Exit Sub
    End If

    Set unifiedCopy = movedObj

    Debug.Print "Unified copy moved successfully."
    Debug.Print "  Unified EntryID: " & unifiedCopy.entryID
    Debug.Print "  Unified StoreID: " & unified.storeID

    ' Reacquire the live source item, then stamp it with unified references
    Set srcLive = Nothing
    On Error Resume Next
    Set srcLive = Application.Session.GetItemFromID(Item.entryID, sourceFolder.storeID)
    On Error GoTo CleanFail

    If Not srcLive Is Nothing Then
        AddOrSetUserProp srcLive, PROP_UNIFIED_ENTRYID, unifiedCopy.entryID
        AddOrSetUserProp srcLive, PROP_UNIFIED_STOREID, unified.storeID
        AddOrSetUserProp srcLive, PROP_UNIFIED_FOLDER, unified.FolderPath
        Debug.Print "Stamped source item with unified IDs."
    Else
        Debug.Print "Could not reacquire live source item for unified backlink."
    End If

    Debug.Print "Linked source <-> unified"
    Debug.Print "  Source:  " & Item.Subject
    Debug.Print "  Unified: " & unifiedCopy.Subject

    RefreshSnapshotForFolderKey MakeFolderKey(sourceFolder.storeID, sourceFolder.entryID)
    RefreshSnapshotForFolderKey MakeFolderKey(unified.storeID, unified.entryID)

CleanExit:
    Set srcLive = Nothing
    Set existingUnified = Nothing
    Set movedObj = Nothing
    Set unifiedCopy = Nothing
    Set tmpCopy = Nothing
    Set sourceFolder = Nothing
    Set unified = Nothing
    Exit Sub

CleanFail:
    Debug.Print "HandleSourceItemAdd ERROR: " & Err.Description
    Resume CleanExit
End Sub

Public Sub HandleFolderRemove(ByVal watcher As clsFolderWatcher)
    Dim folderKey As String
    Dim oldSnap As Object
    Dim newSnap As Object
    Dim missingKeys As Collection
    Dim k As Variant
    Dim packed As String
    Dim link As Object

    On Error GoTo CleanFail

    folderKey = MakeFolderKey(watcher.FolderStoreID, watcher.FolderEntryID)

    If gIsDeleteSyncing Then
        Debug.Print "Delete sync already in progress. Rebuilding snapshots only."
        RefreshAllSnapshots
        Exit Sub
    End If

    If FolderSnapshots Is Nothing Then
        Set FolderSnapshots = CreateObject("Scripting.Dictionary")
    End If

    If Not FolderSnapshots.Exists(folderKey) Then
        Debug.Print "No old snapshot found for folder, rebuilding."
        RefreshSnapshotForFolderKey folderKey
        Exit Sub
    End If

    Set oldSnap = FolderSnapshots(folderKey)
    Set newSnap = BuildSnapshotForFolder(GetFolderByIDs(watcher.FolderStoreID, watcher.FolderEntryID), watcher.isUnified)
    Set missingKeys = FindMissingKeys(oldSnap, newSnap)

    If missingKeys.Count = 0 Then
        Set FolderSnapshots(folderKey) = newSnap
        Debug.Print "No missing items detected after ItemRemove."
        Exit Sub
    End If

    gIsDeleteSyncing = True

    For Each k In missingKeys
        packed = CStr(oldSnap(CStr(k)))
        Set link = UnpackLink(packed)

        Debug.Print "Missing item detected from: " & watcher.FolderPath
        Debug.Print "  Removed item key: " & CStr(k)

        If watcher.isUnified Then
            ' Deleted from unified -> delete source original
            DeleteCounterpart link("sourceEntryID"), link("sourceStoreID"), "source"
        Else
            ' Deleted from source inbox -> delete unified copy
            DeleteCounterpart link("unifiedEntryID"), link("unifiedStoreID"), "unified"
        End If
    Next k

    RefreshAllSnapshots

CleanExit:
    gIsDeleteSyncing = False
    Exit Sub

CleanFail:
    Debug.Print "HandleFolderRemove ERROR: " & Err.Description
    Resume CleanExit
End Sub

Private Sub DeleteCounterpart(ByVal entryID As String, ByVal storeID As String, ByVal sideName As String)
    Dim obj As Object

    If Len(entryID) = 0 Or Len(storeID) = 0 Then
        Debug.Print "DeleteCounterpart skipped: missing " & sideName & " IDs."
        Exit Sub
    End If

    On Error Resume Next
    Set obj = Application.Session.GetItemFromID(entryID, storeID)
    On Error GoTo 0

    If obj Is Nothing Then
        Debug.Print "DeleteCounterpart failed: could not resolve " & sideName & " item."
        Exit Sub
    End If

    Debug.Print "Deleting " & sideName & " counterpart: " & obj.Subject
    obj.Delete
End Sub

' =========================
' SNAPSHOTS
' =========================

Public Sub RefreshAllSnapshots()
    Dim i As Long
    Dim watcher As clsFolderWatcher
    Dim folderKey As String

    If FolderWatchers Is Nothing Then
        Debug.Print "RefreshAllSnapshots aborted: FolderWatchers is Nothing."
        Exit Sub
    End If

    If FolderSnapshots Is Nothing Then
        Set FolderSnapshots = CreateObject("Scripting.Dictionary")
    End If

    For i = 1 To FolderWatchers.Count
        Set watcher = FolderWatchers(i)

        If watcher Is Nothing Then
            Debug.Print "Watcher #" & i & " is Nothing. Skipping."
        Else
            folderKey = MakeFolderKey(watcher.FolderStoreID, watcher.FolderEntryID)

            If FolderSnapshots.Exists(folderKey) Then
                Set FolderSnapshots(folderKey) = BuildSnapshotForFolder( _
                    GetFolderByIDs(watcher.FolderStoreID, watcher.FolderEntryID), _
                    watcher.isUnified)
            Else
                FolderSnapshots.Add folderKey, BuildSnapshotForFolder( _
                    GetFolderByIDs(watcher.FolderStoreID, watcher.FolderEntryID), _
                    watcher.isUnified)
            End If
        End If
    Next i

    Debug.Print "All snapshots refreshed."
End Sub

Public Sub RefreshSnapshotForFolderKey(ByVal folderKey As String)
    Dim i As Long
    Dim watcher As clsFolderWatcher

    If FolderWatchers Is Nothing Then
        Debug.Print "RefreshSnapshotForFolderKey aborted: FolderWatchers is Nothing."
        Exit Sub
    End If

    If FolderSnapshots Is Nothing Then
        Set FolderSnapshots = CreateObject("Scripting.Dictionary")
    End If

    For i = 1 To FolderWatchers.Count
        Set watcher = FolderWatchers(i)

        If Not watcher Is Nothing Then
            If MakeFolderKey(watcher.FolderStoreID, watcher.FolderEntryID) = folderKey Then
                If FolderSnapshots.Exists(folderKey) Then
                    Set FolderSnapshots(folderKey) = BuildSnapshotForFolder( _
                        GetFolderByIDs(watcher.FolderStoreID, watcher.FolderEntryID), _
                        watcher.isUnified)
                Else
                    FolderSnapshots.Add folderKey, BuildSnapshotForFolder( _
                        GetFolderByIDs(watcher.FolderStoreID, watcher.FolderEntryID), _
                        watcher.isUnified)
                End If

                Debug.Print "Snapshot refreshed: " & watcher.FolderPath
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Function BuildSnapshotForFolder(ByVal folder As Outlook.folder, ByVal isUnified As Boolean) As Object
    Dim dict As Object
    Dim items As Outlook.items
    Dim i As Long
    Dim itm As Object
    Dim mail As Outlook.MailItem
    Dim packed As String

    Set dict = CreateObject("Scripting.Dictionary")

    If folder Is Nothing Then
        Set BuildSnapshotForFolder = dict
        Exit Function
    End If

    Set items = folder.items

    For i = 1 To items.Count
        Set itm = items(i)

        If TypeOf itm Is Outlook.MailItem Then
            Set mail = itm

            If isUnified Then
                packed = PackLink( _
                    GetUserPropValue(mail, PROP_SOURCE_ENTRYID), _
                    GetUserPropValue(mail, PROP_SOURCE_STOREID), _
                    mail.entryID, _
                    folder.storeID)
            Else
                packed = PackLink( _
                    mail.entryID, _
                    folder.storeID, _
                    GetUserPropValue(mail, PROP_UNIFIED_ENTRYID), _
                    GetUserPropValue(mail, PROP_UNIFIED_STOREID))
            End If

            dict(mail.entryID) = packed
        End If
    Next i

    Set BuildSnapshotForFolder = dict
End Function

Private Function FindMissingKeys(ByVal oldSnap As Object, ByVal newSnap As Object) As Collection
    Dim c As New Collection
    Dim k As Variant

    For Each k In oldSnap.Keys
        If Not newSnap.Exists(CStr(k)) Then
            c.Add CStr(k)
        End If
    Next k

    Set FindMissingKeys = c
End Function

' =========================
' HELPERS
' =========================

Private Function MakeFolderKey(ByVal storeID As String, ByVal entryID As String) As String
    MakeFolderKey = storeID & "|" & entryID
End Function

Private Function PackLink(ByVal sourceEntryID As String, ByVal sourceStoreID As String, _
                          ByVal unifiedEntryID As String, ByVal unifiedStoreID As String) As String
    PackLink = sourceEntryID & "||" & sourceStoreID & "||" & unifiedEntryID & "||" & unifiedStoreID
End Function

Private Function UnpackLink(ByVal packed As String) As Object
    Dim d As Object
    Dim parts() As String

    Set d = CreateObject("Scripting.Dictionary")
    parts = Split(packed, "||")

    d("sourceEntryID") = ""
    d("sourceStoreID") = ""
    d("unifiedEntryID") = ""
    d("unifiedStoreID") = ""

    If UBound(parts) >= 0 Then d("sourceEntryID") = parts(0)
    If UBound(parts) >= 1 Then d("sourceStoreID") = parts(1)
    If UBound(parts) >= 2 Then d("unifiedEntryID") = parts(2)
    If UBound(parts) >= 3 Then d("unifiedStoreID") = parts(3)

    Set UnpackLink = d
End Function

Private Function GetFolderByIDs(ByVal storeID As String, ByVal entryID As String) As Outlook.folder
    Dim f As Outlook.folder

    On Error Resume Next
    Set f = Application.Session.GetFolderFromID(entryID, storeID)
    On Error GoTo 0

    Set GetFolderByIDs = f
End Function

Private Sub AddOrSetUserProp(ByVal mail As Outlook.MailItem, ByVal propName As String, ByVal propValue As String)
    On Error GoTo EH

    mail.PropertyAccessor.SetProperty propName, propValue
    mail.Save

    Debug.Print "SetProperty OK: " & propName & " = " & propValue
    Exit Sub

EH:
    Debug.Print "SetProperty FAILED: " & propName & " | " & Err.Number & " | " & Err.Description
End Sub

Private Function GetUserPropValue(ByVal mail As Outlook.MailItem, ByVal propName As String) As String
    On Error GoTo EH

    GetUserPropValue = CStr(mail.PropertyAccessor.GetProperty(propName))
    Exit Function

EH:
    GetUserPropValue = ""
End Function

' =========================
' UNIFIED FOLDER PICKER
' =========================

Public Function GetUnifiedFolder() As Outlook.folder
    Static cachedFolder As Outlook.folder
    Dim ns As Outlook.NameSpace
    Dim folderID As String
    Dim storeID As String

    Set ns = Application.Session

    If Not cachedFolder Is Nothing Then
        Set GetUnifiedFolder = cachedFolder
        Exit Function
    End If

    folderID = GetSetting(APP_NAME, APP_SECTION, "FolderID", "")
    storeID = GetSetting(APP_NAME, APP_SECTION, "StoreID", "")

    If folderID <> "" And storeID <> "" Then
        On Error Resume Next
        Set cachedFolder = ns.GetFolderFromID(folderID, storeID)
        On Error GoTo 0
    End If

    If cachedFolder Is Nothing Then
        MsgBox "Choose the real folder you want to use as your Unified Inbox." & vbCrLf & vbCrLf & _
               "Do NOT choose Outlook's built-in All Inboxes view.", _
               vbInformation, "Choose Unified Inbox Folder"

        Set cachedFolder = ns.PickFolder

        If cachedFolder Is Nothing Then
            MsgBox "No folder selected. The unified inbox sync will not run.", vbExclamation
            Exit Function
        End If

        SaveSetting APP_NAME, APP_SECTION, "FolderID", cachedFolder.entryID
        SaveSetting APP_NAME, APP_SECTION, "StoreID", cachedFolder.storeID
    End If

    Set GetUnifiedFolder = cachedFolder
End Function

Public Sub ResetUnifiedFolder()
    On Error Resume Next
    DeleteSetting APP_NAME, APP_SECTION
    On Error GoTo 0
    MsgBox "Unified folder setting reset. Restart Outlook to choose it again.", vbInformation
End Sub

' =========================
' TESTING / DEBUG
' =========================

Public Sub TestUnifiedFolder()
    Dim f As Outlook.folder

    Set f = GetUnifiedFolder()

    If f Is Nothing Then
        Debug.Print "No unified folder selected."
    Else
        Debug.Print "Unified folder: " & f.FolderPath
    End If
End Sub

Public Sub RebuildAllSnapshots()
    If FolderWatchers Is Nothing Then
        Debug.Print "FolderWatchers was not initialized. Running ManualStartup first."
        ManualStartup
    End If

    If FolderWatchers Is Nothing Then
        Debug.Print "FolderWatchers is still Nothing after ManualStartup."
        Exit Sub
    End If

    RefreshAllSnapshots
End Sub

Public Sub ShowSelectedItemLinks()
    Dim mail As Outlook.MailItem

    If Application.ActiveExplorer.Selection.Count = 0 Then
        Debug.Print "No item selected."
        Exit Sub
    End If

    If Not TypeOf Application.ActiveExplorer.Selection.Item(1) Is Outlook.MailItem Then
        Debug.Print "Selected item is not a mail item."
        Exit Sub
    End If

    Set mail = Application.ActiveExplorer.Selection.Item(1)

    Debug.Print "Subject: " & mail.Subject
    Debug.Print "SOURCE ENTRY:  " & GetUserPropValue(mail, PROP_SOURCE_ENTRYID)
    Debug.Print "SOURCE STORE:  " & GetUserPropValue(mail, PROP_SOURCE_STOREID)
    Debug.Print "UNIFIED ENTRY: " & GetUserPropValue(mail, PROP_UNIFIED_ENTRYID)
    Debug.Print "UNIFIED STORE: " & GetUserPropValue(mail, PROP_UNIFIED_STOREID)
End Sub

