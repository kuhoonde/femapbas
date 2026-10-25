' FEMAP API / WinWrap Basic - LIST ONLY.
' Same SHARED_PERIPHERAL criterion as the CBUSH conversion preflight:
' a peripheral node belongs to spiders of two or more DISTINCT selected CBUSHs.
' Standard centered RBE2/RBE3 spiders only. No model entities are changed.
' RBE3 peripheral nodes may currently be independent; this reports sharing
' which would duplicate dependent membership after separate RBE2 conversions.
Option Explicit

' Empty uses the Windows TEMP folder. A custom folder must already exist.
Const REPORT_FOLDER As String = ""
Const SHOW_SHARED_NODES_IN_MESSAGES As Boolean = False

Sub Main
    Dim App As femap.model
    Dim pick As femap.Set, duplicates As femap.Set
    Dim bush As femap.Elem
    Dim owners As Object, links As Object, cache As Object, notes As Object
    Dim d As Object, key As Variant
    Dim eid As Long, ga As Long, gb As Long, rc As Long
    Dim checked As Long, incomplete As Long, beforeCount As Long
    Dim sideA As Boolean, sideB As Boolean, outputBase As String
    Dim failure As String
    On Error GoTo Failed

    Set App = feFemap()
    Set pick = App.feSet
    Set duplicates = App.feSet
    Set bush = App.feElem
    Set owners = NewDictionary()
    Set links = NewDictionary()
    Set cache = NewDictionary()
    Set notes = NewDictionary()

    rc = pick.Select(FT_ELEM, True, "Select ALL CBUSHs to check for shared peripheral nodes - LIST ONLY")
    If rc <> FE_OK Then Exit Sub
    If pick.Count = 0 Then Exit Sub
    rc = App.feAppMessage(FCM_NORMAL, "SHARED PERIPHERAL CHECK | LIST ONLY | selected=" & CStr(pick.Count))

    eid = pick.First
    Do While eid > 0
        If bush.Get(eid) <> FE_OK Then
            AddNote notes, "CBUSH " & CStr(eid) & ": selected element cannot be read."
            incomplete = incomplete + 1
        ElseIf bush.type <> FET_L_SPRING Or bush.topology <> 0 Then
            AddNote notes, "Element " & CStr(eid) & ": not a two-node Spring/Damper; select CBUSH elements only."
            incomplete = incomplete + 1
        Else
            ga = bush.node(0)
            gb = bush.node(1)
            If ga <= 0 Or gb <= 0 Or ga = gb Then
                AddNote notes, "CBUSH " & CStr(eid) & ": two distinct endpoint node IDs are required."
                incomplete = incomplete + 1
            Else
                sideA = ReadSide(App, eid, ga, gb, "A", owners, links, cache, notes)
                sideB = ReadSide(App, eid, gb, ga, "B", owners, links, cache, notes)
                If Not sideA Or Not sideB Then incomplete = incomplete + 1
            End If
        End If
        checked = checked + 1
        If checked Mod 100 = 0 Or checked = pick.Count Then
            rc = App.feAppMessage(FCM_NORMAL, "Checked " & CStr(checked) & "/" & CStr(pick.Count) & " selected elements")
        End If
        eid = pick.Next
    Loop

    ' A repeated entry or several spiders of ONE CBUSH is not cross-CBUSH sharing.
    For Each key In owners.Keys
        Set d = owners.Item(CStr(key))
        If d.Count > 1 Then
            beforeCount = duplicates.Count
            rc = duplicates.Add(CLng(key))
            If duplicates.Count <> beforeCount + 1 Then Err.Raise 513, "Shared node list", "Cannot collect result node " & CStr(key)
        End If
    Next key

    outputBase = MakeOutputBase()
    ExportLists App, duplicates, owners, links, notes, outputBase
    rc = App.feAppMessage(FCM_NORMAL, "DONE | selected=" & CStr(pick.Count) & _
        " | shared peripheral nodes=" & CStr(duplicates.Count) & " | incompletely scanned selections=" & CStr(incomplete))
    rc = App.feAppMessage(FCM_NORMAL, "NODE IDs: " & outputBase & "_nodes.txt")
    rc = App.feAppMessage(FCM_NORMAL, "CBUSH / RBE links: " & outputBase & "_links.csv")
    If incomplete > 0 Or notes.Count > 0 Then
        rc = App.feAppMessage(FCM_ERROR, "PARTIAL SCOPE: check " & outputBase & "_scan_notes.txt. Additional shared nodes may exist in unread/unsupported connections.")
    End If
    Exit Sub

Failed:
    failure = CStr(Err.Number) & " - " & Err.Description
    On Error Resume Next
    rc = App.feAppMessage(FCM_ERROR, "Shared-node scan stopped: " & failure)
    MsgBox "Shared-node scan stopped: " & failure
End Sub

Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Sub AddNote(ByVal notes As Object, ByVal text As String)
    If Not notes.Exists(text) Then notes.Add text, True
End Sub

Function ReadSide(ByVal App As femap.model, ByVal bushID As Long, ByVal center As Long, _
    ByVal otherCenter As Long, ByVal side As String, ByVal owners As Object, _
    ByVal links As Object, ByVal cache As Object, ByVal notes As Object) As Boolean
    Dim adjacent As femap.Set, r As femap.Elem
    Dim rc As Long, id As Long, count As Long, i As Long, nid As Long, found As Long
    Dim ids As Variant, faces As Variant, weights As Variant, dofs As Variant
    Dim cached As Variant, kind As String, ok As Boolean, label As String
    On Error GoTo ReadFailed
    ok = True
    label = "CBUSH " & CStr(bushID) & ", endpoint " & CStr(center)
    Set adjacent = App.feSet
    Set r = App.feElem
    rc = adjacent.AddRule(center, FGD_ELEM_BYNODE)
    If rc <> FE_OK Then Err.Raise 513, "Shared node list", "Cannot read endpoint adjacency"
    id = adjacent.First
    Do While id > 0
        If cache.Exists(CStr(id)) Then
            cached = cache.Item(CStr(id))
            If CLng(cached(0)) = center Then
                found = found + 1
                ids = cached(3)
                count = CLng(cached(2))
                kind = CStr(cached(1))
                RecordNodes bushID, id, kind, center, otherCenter, side, count, ids, owners, links, notes, ok
            Else
                AddNote notes, label & ": RBE " & CStr(id) & " touches the endpoint away from its reference node."
                ok = False
            End If
        ElseIf r.Get(id) <> FE_OK Then
            AddNote notes, label & ": cannot read adjacent element " & CStr(id)
            ok = False
        ElseIf r.type = FET_L_RIGID Then
            If r.topology <> 13 Or r.node(0) <> center Or r.formulation(0) <> 0 Or r.formulation(1) <> 0 Then
                AddNote notes, label & ": rigid " & CStr(id) & _
                    " is outside ordinary centered RBE2/RBE3 scope; topology=" & CStr(r.topology) & _
                    "; reference=" & CStr(r.node(0)) & "; formulation=" & CStr(r.formulation(0)) & "," & CStr(r.formulation(1))
                ok = False
            Else
                rc = r.GetNodeList(0, count, ids, faces, weights, dofs)
                If rc <> FE_OK Or count <= 0 Then
                    AddNote notes, label & ": RBE " & CStr(id) & " has an unreadable/empty peripheral list."
                    ok = False
                Else
                    found = found + 1
                    kind = "RBE2"
                    If r.RigidInterpolate Then kind = "RBE3"
                    cache.Add CStr(id), Array(center, kind, count, ids)
                    RecordNodes bushID, id, kind, center, otherCenter, side, count, ids, owners, links, notes, ok
                End If
            End If
        End If
        id = adjacent.Next
    Loop
    If found = 0 Then
        AddNote notes, label & ": no readable centered RBE2/RBE3 spider was found."
        ok = False
    End If
    ReadSide = ok
    Exit Function
ReadFailed:
    AddNote notes, label & ": read error " & CStr(Err.Number) & " - " & Err.Description
    ReadSide = False
End Function

Sub RecordNodes(ByVal bushID As Long, ByVal rigidID As Long, ByVal kind As String, _
    ByVal center As Long, ByVal otherCenter As Long, ByVal side As String, ByVal count As Long, _
    ByVal ids As Variant, ByVal owners As Object, ByVal links As Object, ByVal notes As Object, ByRef ok As Boolean)
    Dim i As Long, nid As Long, relationKey As String
    Dim d As Object, records As Object
    For i = 0 To count - 1
        nid = CLng(ids(LBound(ids) + i))
        If nid <= 0 Or nid = center Or nid = otherCenter Then
            AddNote notes, "CBUSH " & CStr(bushID) & ", RBE " & CStr(rigidID) & _
                ": invalid/self-referencing peripheral node " & CStr(nid) & " was omitted."
            ok = False
        Else
            If Not owners.Exists(CStr(nid)) Then
                Set d = NewDictionary()
                Set records = NewDictionary()
                owners.Add CStr(nid), d
                links.Add CStr(nid), records
            End If
            Set d = owners.Item(CStr(nid))
            If Not d.Exists(CStr(bushID)) Then d.Add CStr(bushID), True
            Set records = links.Item(CStr(nid))
            relationKey = CStr(bushID) & ":" & CStr(rigidID)
            If Not records.Exists(relationKey) Then
                records.Add relationKey, Array(bushID, rigidID, kind, center, side)
            End If
        End If
    Next i
End Sub

Function MakeOutputBase() As String
    Dim folder As String
    folder = REPORT_FOLDER
    If folder = "" Then folder = Environ$("TEMP")
    If folder = "" Then folder = CurDir$()
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    MakeOutputBase = folder & "CBUSH_shared_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000#))
End Function

Sub ExportLists(ByVal App As femap.model, ByVal duplicates As femap.Set, ByVal owners As Object, _
    ByVal links As Object, ByVal notes As Object, ByVal outputBase As String)
    Dim th As Integer, ch As Integer, nh As Integer
    Dim textOpen As Boolean, csvOpen As Boolean, notesOpen As Boolean
    Dim nid As Long, rc As Long, nodeKey As String, key As Variant, row As Variant
    Dim d As Object, records As Object, failure As String
    On Error GoTo WriteFailed
    th = FreeFile()
    Open outputBase & "_nodes.txt" For Output As #th
    textOpen = True
    ch = FreeFile()
    Open outputBase & "_links.csv" For Output As #ch
    csvOpen = True
    Print #ch, "node_id,distinct_cbush_count,cbush_id,rbe_id,rbe_type,rbe_reference_node,cbush_end"
    nid = duplicates.First
    Do While nid > 0
        nodeKey = CStr(nid)
        Print #th, nodeKey
        Set d = owners.Item(nodeKey)
        Set records = links.Item(nodeKey)
        For Each key In records.Keys
            row = records.Item(key)
            Print #ch, nodeKey & "," & CStr(d.Count) & "," & CStr(row(0)) & "," & CStr(row(1)) & _
                "," & CStr(row(2)) & "," & CStr(row(3)) & "," & CStr(row(4))
        Next key
        If SHOW_SHARED_NODES_IN_MESSAGES Then
            rc = App.feAppMessage(FCM_NORMAL, "SHARED node=" & nodeKey & " CBUSHs=" & JoinKeys(d))
        End If
        nid = duplicates.Next
    Loop
    Close #th
    textOpen = False
    Close #ch
    csvOpen = False
    If notes.Count > 0 Then
        nh = FreeFile()
        Open outputBase & "_scan_notes.txt" For Output As #nh
        notesOpen = True
        For Each key In notes.Keys
            Print #nh, CStr(key)
        Next key
        Close #nh
        notesOpen = False
    End If
    Exit Sub
WriteFailed:
    failure = Err.Description
    On Error Resume Next
    If textOpen Then Close #th
    If csvOpen Then Close #ch
    If notesOpen Then Close #nh
    On Error GoTo 0
    Err.Raise 513, "Shared node list", "List export failed; files may be incomplete: " & failure
End Sub

Function JoinKeys(ByVal d As Object) As String
    Dim key As Variant, text As String
    For Each key In d.Keys
        If Len(text) > 0 Then text = text & ","
        text = text & CStr(key)
    Next key
    JoinKeys = text
End Function
