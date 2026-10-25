' FEMAP API Programming / WinWrap Basic
' Keep the shortest selected two-node RBE2 for each DEPENDENT node.
' Scope: select only the RBE2 links you want to thin, e.g. a new Closest Link batch.
' Non-rigids, multiple-node spiders and weighted interpolation lists are skipped.
' Dependent groups with different constrained DOFs are left unchanged.
' Existing nodes, kept element IDs, CTEs, constraints and materials are not edited.
' Exact equal lengths: keep the element with the lower ID.
' Creates KEEP, DROP and (when needed) SKIPPED groups before any deletion.
' No = preview only (default). Yes = delete the DROP elements only.
' Save a model copy before choosing Yes. This thins the physical connection;
' it is not a remedy or guarantee for a zero-stress check.
' The COM integration must be checked in the installed FEMAP version.
' mesh-connect-closet link 후 중복된 rbe2를 제거하고 가장 짧은 rbe2만 남김. 현재버전은 종속노드가 모두 같고 독립노드가 다른 rbe2를 삭제할때 사용.

Option Explicit

Sub Main
    On Error GoTo Failed

    Const ELEM_ENTITY As Long = 8
    Const API_OK As Long = -1
    Const RIGID_TYPE As Long = 29
    Const RIGID_TOPOLOGY As Long = 13

    Dim app As Object, chosen As Object, elem As Object
    Dim nodeA As Object, nodeB As Object
    Dim keepSet As Object, dropSet As Object, skipSet As Object
    Set app = feFemap()
    Set chosen = app.feSet
    Set elem = app.feElem
    Set nodeA = app.feNode
    Set nodeB = app.feNode
    Set keepSet = app.feSet
    Set dropSet = app.feSet
    Set skipSet = app.feSet

    Dim rc As Long, capacity As Long
    rc = chosen.Select(ELEM_ENTITY, True, _
        "Select the two-node RBE2 links to process (not the whole model)")
    If rc <> API_OK Then Exit Sub
    capacity = chosen.Count
    If capacity < 1 Then Exit Sub

    Dim itemID() As Long, itemGroup() As Long
    Dim dependentID() As Long, winner() As Long, groupDOF() As Long
    Dim bestD2() As Double, blocked() As Boolean
    ReDim itemID(0 To capacity - 1)
    ReDim itemGroup(0 To capacity - 1)
    ReDim dependentID(0 To capacity - 1)
    ReDim winner(0 To capacity - 1)
    ReDim groupDOF(0 To capacity - 1)
    ReDim bestD2(0 To capacity - 1)
    ReDim blocked(0 To capacity - 1)

    Dim eid As Long, gn As Long, gm As Long, nList As Long
    Dim nodes As Variant, faces As Variant, weights As Variant, dofs As Variant
    Dim nItems As Long, nGroups As Long, g As Long, j As Long
    Dim mask As Long, d As Long, d2 As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim eligible As Boolean

    chosen.Reset
    eid = chosen.Next
    Do While eid > 0
        eligible = False
        rc = elem.Get(eid)
        If rc = API_OK Then
            If elem.Type = RIGID_TYPE And elem.topology = RIGID_TOPOLOGY Then
                gn = elem.Node(0)
                nList = 0
                nodes = Empty
                faces = Empty
                weights = Empty
                dofs = Empty
                rc = elem.GetNodeList(0, nList, nodes, faces, weights, dofs)
                If rc = API_OK And nList = 1 And gn > 0 Then
                    If IsArray(nodes) Then
                        If Not HasWeights(weights) Then
                            gm = CLng(nodes(LBound(nodes)))
                            If gm > 0 And gm <> gn Then
                                rc = nodeA.Get(gn)
                                If rc = API_OK Then
                                    rc = nodeB.Get(gm)
                                    If rc = API_OK Then eligible = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If eligible Then
            ' FEMAP stores nodal location coordinates in global rectangular form.
            dx = nodeA.x - nodeB.x
            dy = nodeA.y - nodeB.y
            dz = nodeA.z - nodeB.z
            d2 = dx * dx + dy * dy + dz * dz

            mask = 0
            For d = 0 To 5
                If elem.Release(0, d) <> 0 Then mask = mask + CLng(2 ^ d)
            Next d

            g = -1
            For j = 0 To nGroups - 1
                If dependentID(j) = gm Then
                    g = j
                    Exit For
                End If
            Next j

            If g = -1 Then
                g = nGroups
                dependentID(g) = gm
                groupDOF(g) = mask
                winner(g) = eid
                bestD2(g) = d2
                blocked(g) = False
                nGroups = nGroups + 1
            Else
                If groupDOF(g) <> mask Then blocked(g) = True
                If d2 < bestD2(g) Or _
                    (d2 = bestD2(g) And eid < winner(g)) Then
                    winner(g) = eid
                    bestD2(g) = d2
                End If
            End If
            itemID(nItems) = eid
            itemGroup(nItems) = g
            nItems = nItems + 1
        Else
            skipSet.Add eid
        End If
        eid = chosen.Next
    Loop

    If nItems = 0 Then
        MsgBox "No eligible two-node RBE2 links were found." & vbCrLf & _
            "Select the RBE2 links created by Closest Link.", 48, "Keep shortest RBE2"
        Exit Sub
    End If

    Dim i As Long, protectedGroups As Long
    For g = 0 To nGroups - 1
        If blocked(g) Then protectedGroups = protectedGroups + 1
    Next g
    For i = 0 To nItems - 1
        g = itemGroup(i)
        If blocked(g) Then
            keepSet.Add itemID(i)
            skipSet.Add itemID(i)
        ElseIf itemID(i) = winner(g) Then
            keepSet.Add itemID(i)
        Else
            dropSet.Add itemID(i)
        End If
    Next i

    Dim keepGroup As Long, dropGroup As Long, skipGroup As Long
    keepGroup = NewElementGroup(app, keepSet, "RBE2_DEP_KEEP_SHORTEST")
    If dropSet.Count > 0 Then
        dropGroup = NewElementGroup(app, dropSet, "RBE2_DEP_DROP_LONGER")
    End If
    If skipSet.Count > 0 Then
        skipGroup = NewElementGroup(app, skipSet, "RBE2_SKIPPED_REVIEW")
    End If

    Dim prompt As String, choice As Long
    prompt = "Grouping by DEPENDENT node (not independent node)." & vbCrLf & _
        "Selected elements: " & CStr(capacity) & vbCrLf & _
        "Dependent node groups: " & CStr(nGroups) & vbCrLf & _
        "Keep: " & CStr(keepSet.Count) & vbCrLf & _
        "Delete candidates: " & CStr(dropSet.Count) & vbCrLf & _
        "Skipped / protected: " & CStr(skipSet.Count) & vbCrLf & _
        "Mixed-DOF dependent nodes protected: " & CStr(protectedGroups) & vbCrLf & vbCrLf & _
        "KEEP group ID: " & CStr(keepGroup) & vbCrLf & _
        "DROP group ID: " & CStr(dropGroup)
    If skipGroup > 0 Then prompt = prompt & vbCrLf & "SKIPPED group ID: " & CStr(skipGroup)

    If dropSet.Count = 0 Then
        MsgBox prompt & vbCrLf & vbCrLf & "Nothing was deleted.", 64, "Keep shortest RBE2"
        Exit Sub
    End If

    prompt = prompt & vbCrLf & vbCrLf & _
        "Yes = DELETE the candidate RBE2 elements now." & vbCrLf & _
        "No = PREVIEW ONLY; inspect the groups first." & vbCrLf & _
        "Cancel = leave the elements unchanged." & vbCrLf & vbCrLf & _
        "Deletion reduces the interface connection density." & vbCrLf & _
        "Use a saved model copy before choosing Yes."
    choice = MsgBox(prompt, 3 + 256 + 32, "Keep shortest RBE2 - default is preview")
    If choice <> 6 Then Exit Sub

    ' Only element entities are deleted; no node deletion or property cleanup.
    rc = app.feDelete(ELEM_ENTITY, dropSet.ID)

    Dim requested As Long, remaining As Long
    For i = 0 To nItems - 1
        g = itemGroup(i)
        If Not blocked(g) Then
            If itemID(i) <> winner(g) Then
                requested = requested + 1
                If elem.Get(itemID(i)) = API_OK Then remaining = remaining + 1
            End If
        End If
    Next i
    app.feViewRegenerate 0
    MsgBox "Deleted: " & CStr(requested - remaining) & vbCrLf & _
        "Candidate elements still present: " & CStr(remaining) & vbCrLf & _
        "Kept elements and their CTE values were not edited.", _
        64, "Keep shortest RBE2"
    Exit Sub

Failed:
    MsgBox "Stopped: " & Err.Description & vbCrLf & _
        "The script does not roll back an operation already completed." & vbCrLf & _
        "Inspect the model and the preview groups before continuing.", _
        16, "Keep shortest RBE2"
End Sub

Private Function HasWeights(ByVal values As Variant) As Boolean
    Dim k As Long
    HasWeights = False
    If IsArray(values) Then
        For k = LBound(values) To UBound(values)
            If CDbl(values(k)) <> 0# Then
                HasWeights = True
                Exit Function
            End If
        Next k
    End If
End Function

Private Function NewElementGroup(ByVal app As Object, ByVal elementSet As Object, _
    ByVal groupName As String) As Long
    Dim groupObj As Object, gid As Long, rc As Long
    Set groupObj = app.feGroup
    gid = groupObj.NextEmptyID
    groupObj.title = groupName & "_" & CStr(gid)
    groupObj.SetAdd 8, elementSet.ID
    rc = groupObj.Put(gid)
    If rc <> -1 Then
        Err.Raise 513, "NewElementGroup", "Unable to create the preview group."
    End If
    NewElementGroup = gid
End Function

