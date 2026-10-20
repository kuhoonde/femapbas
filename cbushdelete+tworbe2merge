' FEMAP API / WinWrap Basic
' Selected CBUSH + endpoint-centered RBE2/RBE3 spiders -> one NEW RBE2.
' One separate midpoint node and one RBE2 for EACH selected CBUSH.
' Revision 3: use ordinary RBE2 common DOFs; add staged progress messages.
'
' RUN ON A SAVED WORKING COPY. This intentionally changes joint stiffness.
' RBE3 weights/interpolation and CBUSH stiffness/releases are NOT retained.
' Peripheral mesh nodes are never moved or merged. No proximity matching.
' Old endpoint nodes are deleted only after checking their remaining references.
' Does not save the model or modify analysis-set options (keep RIGID=LAGRAN).
'
' 1. Leave PREVIEW_ONLY=True, run Main (F5), and select actual CBUSH elements.
' 2. Review ALL messages. A single blocking issue prevents the entire batch.
' 3. Change PREVIEW_ONLY=False and run again with the same selection to apply.
'
' OUTPUT_DEPENDENT_DOFS="123456": explicitly connect translations/rotations
' at ALL new dependent nodes. This is a target setting, NOT preservation of
' RBE3 peripheral kinematics. Set "123" for translations only, or "" to
' require a COMMON inherited peripheral DOF selection instead.
'   RBE2: read its common DOFs from release(0,...).
'   RBE3: read the DOFs of its surrounding nodes from GetNodeList(0,...).
' In inherited mode only, differing source DOFs block conversion.
' The SOURCE messages report original types/DOFs; RBE3 outer nodes become dependent.
' CTE is copied from the original RBE elements and must agree within each pair.
'
' Supports ordinary single-reference-node RBE2/RBE3 spiders (topology 13).
' Rejects RBE1/multilist, RSPLINE, extra node lists/non-Boolean DOF flags, nested/shared
' endpoint chains, parallel springs, cross-pair peripheral sharing, external
' rigid elements touching the new dependent nodes, and conflicting nodal BCs.
' All endpoint-adjacent elements must be this CBUSH or centered RBE spiders.
' Several ordinary spiders centered on the SAME endpoint are allowed.
'
' Nodal/elemental loads on records to be deleted are NOT transferred. They
' cause a blocking issue. This includes explicit temperatures on old centers.
' For the current uniform 0->100 check, the new center uses the load set's
' default nodal temperatures (TEMPD). Supply temperatures separately when
' later using nonuniform nodal temperatures without suitable defaults.
' Geometry-based BCs are NOT expanded; recheck normal solver export afterward.
' Does not update groups, output requests, or external ID-based references.
'
' Preflight is read-only. Apply verifies new connectivity and deleted records.
' On a write failure, attempts to restore original node/element records from
' separate in-memory FEMAP objects. This is NOT a full database transaction;
' use the saved copy if recovery fails or ancillary references need recovery.
' FEMAP/WinWrap integration must be checked in the user's installed FEMAP.
'
' API source: Siemens FEMAP API Reference, Common Entity / Node / Elem /
' BCNode / BCEqn / LoadMesh / LoadNTemp / LoadETemp objects.
' https://www.scribd.com/document/150750537/Femap-API-Reference

Option Explicit

Const PREVIEW_ONLY As Boolean = False
Const OUTPUT_DEPENDENT_DOFS As String = "123456"
Const PROGRESS_INTERVAL_SECONDS As Double = 5#
Const PROGRESS_EVERY_ITEMS As Long = 100

Sub Main
    Dim App As femap.model
    Dim pick As femap.Set
    Dim el As femap.Elem, savedEl As femap.Elem
    Dim nd As femap.Node, savedNode As femap.Node
    Dim oldElements As Object, oldNodes As Object, plans As Object
    Dim issues As Object, centerOwners As Object, outerOwners As Object
    Dim outerMasks As Object, createdElements As Object, createdNodes As Object
    Dim legs As Object, spiders As Object
    Dim row As Variant, key As Variant, nk As Variant, mid As Variant
    Dim legIDs As Variant, spiderIDs As Variant
    Dim eid As Long, ga As Long, gb As Long, rc As Long, i As Long
    Dim mask As Long, overrideMask As Long, baseMask As Long
    Dim nextNode As Long, nextElement As Long, sequence As Long
    Dim alpha As Double, haveAlpha As Boolean, writing As Boolean
    Dim failure As String, recoveryOK As Boolean
    Dim processed As Long, progressTick As Double, mode As String

    On Error GoTo Failed
    Set App = feFemap()
    Set pick = App.feSet
    Set el = App.feElem
    Set nd = App.feNode
    Set oldElements = NewDictionary()
    Set oldNodes = NewDictionary()
    Set plans = NewDictionary()
    Set issues = NewDictionary()
    Set centerOwners = NewDictionary()
    Set outerOwners = NewDictionary()
    Set outerMasks = NewDictionary()
    Set createdElements = NewDictionary()
    Set createdNodes = NewDictionary()
    overrideMask = ParseDOFs(OUTPUT_DEPENDENT_DOFS)

    rc = pick.Select(FT_ELEM, True, "Select CBUSH elements - replace both RBE spiders with ONE RBE2")
    If rc <> FE_OK Then Exit Sub
    If pick.Count = 0 Then Exit Sub
    mode = "APPLY"
    If PREVIEW_ONLY Then mode = "PREVIEW"
    rc = App.feAppMessage(FCM_NORMAL, "REVISION 3 | " & mode & " | " & CStr(pick.Count) & " selected CBUSH elements")
    If overrideMask <> 0 Then
        rc = App.feAppMessage(FCM_NORMAL, "TARGET: new RBE2 dependent DOF=" & _
            DOFText(overrideMask) & " for every peripheral node (explicit override; source DOFs may differ).")
    End If

    ' Reserve fresh IDs above the ORIGINAL maxima, never IDs of deleted records.
    If nd.Last <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "No nodes in the model."
    If el.Last <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "No elements in the model."
    If CDbl(nd.ID) + CDbl(pick.Count) > 2147483647# Or _
       CDbl(el.ID) + CDbl(pick.Count) > 2147483647# Then
        Err.Raise 513, "CBUSH to RBE2", "Not enough fresh IDs for this batch."
    End If
    nextNode = nd.ID
    nextElement = el.ID

    ReportProgress App, "1/7 Read CBUSH chains", 0, pick.Count, progressTick
    eid = pick.First
    Do While eid > 0
        If el.Get(eid) <> FE_OK Then
            AddIssue issues, "Cannot read selected element " & CStr(eid)
        ElseIf el.type <> FET_L_SPRING Or el.topology <> 0 Then
            AddIssue issues, "Element " & CStr(eid) & " is not a two-node Spring/Damper (select CBUSH only)."
        Else
            ga = el.node(0)
            gb = el.node(1)
            If ga <= 0 Or gb <= 0 Or ga = gb Then
                AddIssue issues, "CBUSH " & CStr(eid) & " needs two DISTINCT endpoint node IDs."
            Else
                SaveElement App, eid, oldElements
                RegisterOwner centerOwners, CStr(ga), CStr(eid), "Endpoint", issues
                RegisterOwner centerOwners, CStr(gb), CStr(eid), "Endpoint", issues
                SaveCenter App, ga, oldNodes, issues
                SaveCenter App, gb, oldNodes, issues
                Set legs = NewDictionary()
                Set spiders = NewDictionary()
                baseMask = 0
                alpha = 0#
                haveAlpha = False
                CollectSide App, eid, ga, gb, legs, spiders, oldElements, _
                    baseMask, alpha, haveAlpha, overrideMask, issues
                CollectSide App, eid, gb, ga, legs, spiders, oldElements, _
                    baseMask, alpha, haveAlpha, overrideMask, issues
                mask = baseMask
                If overrideMask <> 0 Then mask = overrideMask
                If mask = 0 Then AddIssue issues, "CBUSH " & CStr(eid) & ": no usable peripheral DOFs."
                If legs.Count = 0 Then AddIssue issues, "CBUSH " & CStr(eid) & ": no peripheral nodes."
                For Each nk In legs.Keys
                    RegisterOwner outerOwners, CStr(nk), CStr(eid), "Peripheral node", issues
                    If Not outerMasks.Exists(CStr(nk)) Then outerMasks.Add CStr(nk), mask
                Next nk
                If oldNodes.Exists(CStr(ga)) And oldNodes.Exists(CStr(gb)) Then
                    Set savedNode = oldNodes.Item(CStr(ga))
                    mid = Array(savedNode.x, savedNode.y, savedNode.z)
                    Set savedNode = oldNodes.Item(CStr(gb))
                    mid(0) = mid(0) + (savedNode.x - mid(0)) / 2#
                    mid(1) = mid(1) + (savedNode.y - mid(1)) / 2#
                    mid(2) = mid(2) + (savedNode.z - mid(2)) / 2#
                    sequence = sequence + 1
                    legIDs = legs.Keys
                    spiderIDs = spiders.Keys
                    ' row: old centers, XYZ, outer IDs, old RBE IDs, DOFs, CTE, fresh IDs.
                    plans.Add CStr(eid), Array(ga, gb, mid, legIDs, spiderIDs, mask, _
                        alpha, nextNode + sequence, nextElement + sequence)
                End If
            End If
        End If
        processed = processed + 1
        ReportProgress App, "1/7 Read CBUSH chains", processed, pick.Count, progressTick
        eid = pick.Next
    Loop

    For Each nk In outerOwners.Keys
        If centerOwners.Exists(CStr(nk)) Then
            AddIssue issues, "Node " & CStr(nk) & " is both an endpoint and a peripheral node in this batch."
        End If
    Next nk
    CheckOuterNodes App, outerOwners, outerMasks, oldElements, issues
    CheckConstraints App, oldNodes, outerMasks, issues
    CheckLoads App, oldNodes, oldElements, outerMasks, issues
    rc = App.feAppMessage(FCM_NORMAL, "PREFLIGHT COMPLETE: " & CStr(plans.Count) & _
        " planned conversions; " & CStr(issues.Count) & " blocking issue(s).")

    For Each key In plans.Keys
        row = plans.Item(key)
        mid = row(2)
        legIDs = row(3)
        spiderIDs = row(4)
        rc = App.feAppMessage(FCM_NORMAL, "CBUSH " & CStr(key) & ": remove RBE [" & _
            JoinIDs(spiderIDs) & "]; centers " & CStr(row(0)) & "," & CStr(row(1)) & _
            " -> NEW node " & CStr(row(7)) & " at (" & CStr(mid(0)) & ", " & _
            CStr(mid(1)) & ", " & CStr(mid(2)) & "); NEW RBE2 " & CStr(row(8)) & _
            "; " & CStr(IDCount(legIDs)) & " peripheral nodes" & _
            "; DOF=" & DOFText(CLng(row(5))) & "; CTE=" & CStr(row(6)))
        If CDbl(row(6)) = 0# Then
            rc = App.feAppMessage(FCM_NORMAL, "NOTE: CBUSH " & CStr(key) & _
                " inherits ZERO rigid CTE. Verify this is intended for a thermal check.")
        End If
    Next key
    If issues.Count > 0 Then
        For Each key In issues.Keys
            rc = App.feAppMessage(FCM_ERROR, "BLOCKED: " & CStr(key))
        Next key
        rc = App.feAppMessage(FCM_ERROR, CStr(issues.Count) & _
            " blocking issue(s). No node/element writes or deletes were started.")
        Exit Sub
    End If
    If plans.Count <> pick.Count Then Err.Raise 513, "CBUSH to RBE2", "Incomplete plan; no changes."
    If PREVIEW_ONLY Then
        rc = App.feAppMessage(FCM_NORMAL, "PREVIEW ONLY: " & CStr(plans.Count) & _
            " CBUSH -> " & CStr(plans.Count) & " separate new RBE2s. No model changes. " & _
            "Check DOF/CTE/topology, then set PREVIEW_ONLY=False to apply on a saved copy.")
        Exit Sub
    End If

    processed = 0
    ReportProgress App, "5/7 Check element deletability", 0, oldElements.Count, progressTick
    For Each key In oldElements.Keys
        If el.Deletable(CLng(key)) <> FE_OK Then
            Err.Raise 513, "CBUSH to RBE2", "Element " & CStr(key) & " has a remaining model reference and cannot be deleted."
        End If
        processed = processed + 1
        ReportProgress App, "5/7 Check element deletability", processed, oldElements.Count, progressTick
    Next key
    ' First mutation. No node coordinate or connectivity was written above.
    writing = True
    processed = 0
    ReportProgress App, "5/7 Delete old elements", 0, oldElements.Count, progressTick
    For Each key In oldElements.Keys
        If el.Delete(CLng(key)) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot delete element " & CStr(key)
        processed = processed + 1
        ReportProgress App, "5/7 Delete old elements", processed, oldElements.Count, progressTick
    Next key
    ' Load/BC preflight plus FEMAP's reference checks protect old endpoint data.
    processed = 0
    ReportProgress App, "5/7 Check old center references", 0, oldNodes.Count, progressTick
    For Each key In oldNodes.Keys
        If nd.Deletable(CLng(key)) <> FE_OK Then
            Err.Raise 513, "CBUSH to RBE2", "Old center " & CStr(key) & _
                " still has a model reference; no automatic load/constraint transfer is allowed."
        End If
        processed = processed + 1
        ReportProgress App, "5/7 Check old center references", processed, oldNodes.Count, progressTick
    Next key
    processed = 0
    ReportProgress App, "5/7 Delete old centers", 0, oldNodes.Count, progressTick
    For Each key In oldNodes.Keys
        If nd.Delete(CLng(key)) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot delete old center " & CStr(key)
        processed = processed + 1
        ReportProgress App, "5/7 Delete old centers", processed, oldNodes.Count, progressTick
    Next key

    processed = 0
    ReportProgress App, "6/7 Create midpoint nodes and RBE2s", 0, plans.Count, progressTick
    For Each key In plans.Keys
        row = plans.Item(key)
        If nd.Exist(CLng(row(7))) = FE_OK Or el.Exist(CLng(row(8))) = FE_OK Then
            Err.Raise 513, "CBUSH to RBE2", "Fresh ID collision; conversion stopped."
        End If
        Set savedNode = oldNodes.Item(CStr(row(0)))
        mid = row(2)
        Set nd = App.feNode
        nd.type = 0
        nd.defCSys = 0
        nd.outCSys = 0
        nd.superelementID = 0
        nd.layer = savedNode.layer
        nd.color = savedNode.color
        nd.x = CDbl(mid(0))
        nd.y = CDbl(mid(1))
        nd.z = CDbl(mid(2))
        For i = 0 To 5
            nd.permBC(i) = False
        Next i
        createdNodes.Add CStr(row(7)), True
        If nd.Put(CLng(row(7))) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot create new center " & CStr(row(7))
        Set savedEl = oldElements.Item(CStr(key))
        Set el = App.feElem
        PrepareRBE2 el, CLng(row(7)), row(3), CLng(row(5)), CDbl(row(6))
        el.layer = savedEl.layer
        el.color = savedEl.color
        createdElements.Add CStr(row(8)), True
        If el.Put(CLng(row(8))) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot create new RBE2 " & CStr(row(8))
        processed = processed + 1
        ReportProgress App, "6/7 Create midpoint nodes and RBE2s", processed, plans.Count, progressTick
    Next key
    processed = 0
    ReportProgress App, "7/7 Verify new connectivity", 0, plans.Count, progressTick
    For Each key In plans.Keys
        row = plans.Item(key)
        VerifyNew App, row
        processed = processed + 1
        ReportProgress App, "7/7 Verify new connectivity", processed, plans.Count, progressTick
    Next key
    For Each key In oldElements.Keys
        If el.Exist(CLng(key)) = FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Old element remains: " & CStr(key)
    Next key
    For Each key In oldNodes.Keys
        If nd.Exist(CLng(key)) = FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Old center remains: " & CStr(key)
    Next key
    writing = False
    rc = App.feViewRegenerate(0)
    rc = App.feAppMessage(FCM_NORMAL, "DONE: " & CStr(plans.Count) & " CBUSH chains replaced by " & _
        CStr(plans.Count) & " new RBE2s and midpoint nodes. Peripheral node IDs/coordinates unchanged. " & _
        "Model not saved. Check exported RBE2 ALPHA and keep RIGID=LAGRAN for the thermal check.")
    Exit Sub

Failed:
    failure = Err.Description
    On Error Resume Next
    If writing Then
        rc = App.feAppMessage(FCM_ERROR, "RECOVERY: restoring original node/element records after a failed write.")
        recoveryOK = RestoreMesh(App, oldNodes, oldElements, createdNodes, createdElements)
        If recoveryOK Then
            failure = failure & " Original node/element records restored. Ancillary references are not a full database rollback."
        Else
            failure = failure & " RESTORE INCOMPLETE: reload your saved working copy; do not save this state."
        End If
    Else
        failure = failure & " No node/element writes or deletes were started."
    End If
    rc = App.feAppMessage(FCM_ERROR, failure)
End Sub

Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Sub AddIssue(ByVal issues As Object, ByVal text As String)
    If Not issues.Exists(text) Then issues.Add text, True
End Sub

Sub RegisterOwner(ByVal owners As Object, ByVal key As String, ByVal owner As String, _
                  ByVal label As String, ByVal issues As Object)
    If owners.Exists(key) Then
        If CStr(owners.Item(key)) <> owner Then
            AddIssue issues, label & " " & key & " is shared by CBUSH " & _
                CStr(owners.Item(key)) & " and " & owner & "; no automatic chain flattening."
        End If
    Else
        owners.Add key, owner
    End If
End Sub

Sub SaveElement(ByVal App As femap.model, ByVal eid As Long, ByVal originals As Object)
    Dim snapshot As femap.Elem
    If originals.Exists(CStr(eid)) Then Exit Sub
    Set snapshot = App.feElem
    If snapshot.Get(eid) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot snapshot element " & CStr(eid)
    originals.Add CStr(eid), snapshot
End Sub

Sub SaveCenter(ByVal App As femap.model, ByVal nid As Long, ByVal originals As Object, ByVal issues As Object)
    Dim snapshot As femap.Node
    Dim k As Long
    If originals.Exists(CStr(nid)) Then Exit Sub
    Set snapshot = App.feNode
    If snapshot.Get(nid) <> FE_OK Then
        AddIssue issues, "Cannot read endpoint " & CStr(nid)
        Exit Sub
    End If
    If snapshot.type <> 0 Or snapshot.superelementID <> 0 Then
        AddIssue issues, "Endpoint " & CStr(nid) & " is not an ordinary residual GRID."
    End If
    For k = 0 To 5
        If snapshot.permBC(k) Then AddIssue issues, "Endpoint " & CStr(nid) & " has a permanent constraint."
    Next k
    originals.Add CStr(nid), snapshot
End Sub

Sub CollectSide(ByVal App As femap.model, ByVal bushID As Long, ByVal center As Long, _
    ByVal otherCenter As Long, ByVal legs As Object, ByVal spiders As Object, _
    ByVal originals As Object, ByRef baseMask As Long, ByRef alpha As Double, _
    ByRef haveAlpha As Boolean, ByVal overrideMask As Long, ByVal issues As Object)
    Dim adjacent As femap.Set
    Dim r As femap.Elem
    Dim id As Long, count As Long, n As Long, j As Long, k As Long, found As Long
    Dim ids As Variant, faces As Variant, weights As Variant, dofs As Variant
    Dim mask As Long, entryMask As Long, flag As Long
    Dim rc As Long, label As String, kind As String
    Dim sourceMasks As Object
    Set adjacent = App.feSet
    Set r = App.feElem
    rc = adjacent.AddRule(center, FGD_ELEM_BYNODE)
    id = adjacent.First
    Do While id > 0
        If id <> bushID Then
            label = "CBUSH " & CStr(bushID) & ", endpoint " & CStr(center) & ", element " & CStr(id)
            If r.Get(id) <> FE_OK Then
                AddIssue issues, label & ": cannot read adjacent element."
            ElseIf r.type <> FET_L_RIGID Then
                AddIssue issues, label & ": unsupported adjacency (FEMAP type " & CStr(r.type) & ")."
            ElseIf r.topology <> 13 Or r.node(0) <> center Then
                AddIssue issues, label & ": endpoint is not the reference/center of an ordinary RBE2/RBE3."
            ElseIf r.formulation(0) <> 0 Or r.formulation(1) <> 0 Then
                AddIssue issues, label & ": nondefault rigid formulation is not supported."
            Else
                found = found + 1
                For k = 1 To 19
                    If r.node(k) <> 0 Then AddIssue issues, label & ": extra reference nodes are not supported."
                Next k
                ' release(1,*) is not a documented stand-alone UM indicator.
                ' Do not classify ordinary RBE3s as UM based on those raw flags.
                CheckExtraRigidLists r, label, issues
                If r.GetNodeList(0, count, ids, faces, weights, dofs) <> FE_OK Then
                    AddIssue issues, label & ": cannot read peripheral node list."
                ElseIf count <= 0 Then
                    AddIssue issues, label & ": empty peripheral node list."
                Else
                    If Not spiders.Exists(CStr(id)) Then spiders.Add CStr(id), True
                    SaveElement App, id, originals
                    Set sourceMasks = NewDictionary()
                    If Not haveAlpha Then
                        alpha = r.RigidThermalExpansion
                        haveAlpha = True
                    ElseIf Abs(alpha - r.RigidThermalExpansion) > _
                        0.000000000001 * (1# + Abs(alpha) + Abs(r.RigidThermalExpansion)) Then
                        AddIssue issues, label & ": RBE thermal expansion coefficients differ."
                    End If
                    mask = 0
                    If Not r.RigidInterpolate Then
                        For k = 0 To 5
                            If r.release(0, k) Then mask = mask Or DOFBit(k)
                        Next k
                        If mask = 0 And overrideMask = 0 Then AddIssue issues, _
                            label & ": RBE2 has no common DOFs to inherit; explicitly set OUTPUT_DEPENDENT_DOFS."
                    End If
                    For n = 0 To count - 1
                        j = CLng(ids(LBound(ids) + n))
                        If r.RigidInterpolate Then
                            ' Only interpolation elements use per-node DOF flags.
                            ' Ordinary RBE2 uses its COMMON release(0,...) DOFs.
                            entryMask = 0
                            For k = 0 To 5
                                flag = CLng(dofs(LBound(dofs) + n * 6 + k))
                                If flag <> 0 And flag <> 1 And flag <> -1 Then
                                    AddIssue issues, label & ": non-Boolean RBE3 peripheral DOF flags are unsupported."
                                End If
                                If flag <> 0 Then entryMask = entryMask Or DOFBit(k)
                            Next k
                            If entryMask = 0 Then AddIssue issues, label & ": RBE3 peripheral node has no DOFs."
                            mask = entryMask
                        End If
                        If Not sourceMasks.Exists(DOFText(mask)) Then sourceMasks.Add DOFText(mask), True
                        If overrideMask = 0 Then
                            If baseMask = 0 Then baseMask = mask
                            If mask <> baseMask Then
                                AddIssue issues, label & ": peripheral DOFs " & DOFText(mask) & _
                                    " differ from " & DOFText(baseMask) & "; explicitly set OUTPUT_DEPENDENT_DOFS if intended."
                            End If
                        End If
                        If j <= 0 Or j = center Or j = otherCenter Then
                            AddIssue issues, label & ": invalid/self-referencing peripheral node " & CStr(j)
                        ElseIf Not legs.Exists(CStr(j)) Then
                            legs.Add CStr(j), True
                        End If
                    Next n
                    kind = "RBE2"
                    If r.RigidInterpolate Then kind = "RBE3"
                    rc = App.feAppMessage(FCM_NORMAL, "SOURCE " & kind & " " & CStr(id) & _
                        ": center=" & CStr(center) & "; " & CStr(count) & " peripheral entries" & _
                        "; peripheral DOFs=[" & JoinIDs(sourceMasks.Keys) & "]" & _
                        "; CTE=" & CStr(r.RigidThermalExpansion) & _
                        "; raw release0=" & DOFText(RawReleaseMask(r, 0)) & _
                        "; raw release1=" & DOFText(RawReleaseMask(r, 1)))
                End If
            End If
        End If
        id = adjacent.Next
    Loop
    If found = 0 Then AddIssue issues, "CBUSH " & CStr(bushID) & ": endpoint " & _
        CStr(center) & " has no supported centered RBE2/RBE3."
End Sub

Sub CheckExtraRigidLists(ByVal r As femap.Elem, ByVal label As String, ByVal issues As Object)
    Dim listIndex As Long, count As Long, rc As Long
    Dim ids As Variant, faces As Variant, weights As Variant, flags As Variant
    ' Ordinary spiders use list 0. Invalid/unavailable lists return a non-FE_OK
    ' code; an actual populated additional list requires separate treatment.
    For listIndex = 1 To 3
        count = 0
        rc = r.GetNodeList(listIndex, count, ids, faces, weights, flags)
        If rc = FE_OK And count > 0 Then AddIssue issues, label & _
            ": additional rigid node list " & CStr(listIndex) & " is populated; this form is unsupported."
    Next listIndex
End Sub

Function RawReleaseMask(ByVal r As femap.Elem, ByVal side As Long) As Long
    Dim k As Long, mask As Long
    For k = 0 To 5
        If r.release(side, k) Then mask = mask Or DOFBit(k)
    Next k
    RawReleaseMask = mask
End Function

Sub CheckOuterNodes(ByVal App As femap.model, ByVal owners As Object, ByVal masks As Object, _
                    ByVal oldElements As Object, ByVal issues As Object)
    Dim adjacent As femap.Set
    Dim nd As femap.Node, el As femap.Elem
    Dim key As Variant, id As Long, rc As Long, k As Long
    Dim processed As Long, progressTick As Double
    Set adjacent = App.feSet
    Set nd = App.feNode
    Set el = App.feElem
    ReportProgress App, "2/7 Check peripheral connectivity", 0, owners.Count, progressTick
    For Each key In owners.Keys
        If nd.Get(CLng(key)) <> FE_OK Then
            AddIssue issues, "Cannot read peripheral node " & CStr(key)
        Else
            If nd.type <> 0 Or nd.superelementID <> 0 Then AddIssue issues, _
                "Peripheral node " & CStr(key) & " is not an ordinary residual GRID."
            For k = 0 To 5
                If nd.permBC(k) And ((CLng(masks.Item(key)) And DOFBit(k)) <> 0) Then
                    AddIssue issues, "Peripheral node " & CStr(key) & " has a conflicting permanent constraint."
                End If
            Next k
            rc = adjacent.Clear
            rc = adjacent.AddRule(CLng(key), FGD_ELEM_BYNODE)
            id = adjacent.First
            Do While id > 0
                If Not oldElements.Exists(CStr(id)) Then
                    If el.Get(id) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot read adjacent element " & CStr(id)
                    If el.type = FET_L_RIGID Then AddIssue issues, "Peripheral node " & CStr(key) & _
                        " also touches external rigid element " & CStr(id) & "; nested/external RBE requires manual review."
                End If
                id = adjacent.Next
            Loop
        End If
        processed = processed + 1
        ReportProgress App, "2/7 Check peripheral connectivity", processed, owners.Count, progressTick
    Next key
End Sub

Sub CheckConstraints(ByVal App As femap.model, ByVal centers As Object, ByVal masks As Object, ByVal issues As Object)
    Dim bs As femap.BCSet, bc As femap.BCNode, eq As femap.BCEqn
    Dim key As Variant, rc As Long, k As Long, id As Long, mask As Long
    Dim relevant As Object
    Dim processed As Long, scanned As Long, progressTick As Double, stage As String
    Set relevant = NewDictionary()
    For Each key In centers.Keys
        relevant.Add CStr(key), 63
    Next key
    For Each key In masks.Keys
        If Not relevant.Exists(CStr(key)) Then relevant.Add CStr(key), CLng(masks.Item(key))
    Next key
    Set bs = App.feBCSet
    Set bc = App.feBCNode
    Set eq = App.feBCEqn
    rc = App.feAppMessage(FCM_NORMAL, "STAGE 3/7: Check constraints and constraint equations")
    rc = bs.First
    Do While rc = FE_OK
        processed = 0
        stage = "3/7 Nodal constraints, set " & CStr(bs.ID)
        ReportProgress App, stage, 0, relevant.Count, progressTick
        bc.SetID = bs.ID
        eq.SetID = bs.ID
        For Each key In relevant.Keys
            If bc.Get(CLng(key)) = FE_OK Then
                For k = 0 To 5
                    If bc.dof(k) And ((CLng(relevant.Item(key)) And DOFBit(k)) <> 0) Then
                        AddIssue issues, "Constraint set " & CStr(bs.ID) & ", node " & CStr(key) & _
                            ": constraint conflicts with deletion/new dependent DOFs."
                    End If
                Next k
            End If
            processed = processed + 1
            ReportProgress App, stage, processed, relevant.Count, progressTick
        Next key
        scanned = 0
        stage = "3/7 Constraint equations, set " & CStr(bs.ID)
        ReportProgress App, stage, 0, 0, progressTick
        If eq.First = FE_OK Then
            Do
                For k = 0 To eq.count - 1
                    id = eq.node(k)
                    If relevant.Exists(CStr(id)) Then
                        mask = CLng(relevant.Item(CStr(id)))
                        If eq.dof(k) >= 1 And eq.dof(k) <= 6 Then
                            If (mask And DOFBit(eq.dof(k) - 1)) <> 0 Then AddIssue issues, _
                                "Constraint equation " & CStr(eq.ID) & " in set " & CStr(bs.ID) & _
                                " references node " & CStr(id) & "; review manually."
                        End If
                    End If
                Next k
                scanned = scanned + 1
                ReportProgress App, stage, scanned, 0, progressTick
            Loop While eq.Next = FE_OK
        End If
        ReportProgress App, stage, scanned, scanned, progressTick
        rc = bs.Next
    Loop
End Sub

Sub CheckLoads(ByVal App As femap.model, ByVal centers As Object, ByVal elements As Object, _
               ByVal masks As Object, ByVal issues As Object)
    Dim ls As femap.LoadSet, nt As femap.LoadNTemp, et As femap.LoadETemp, lm As femap.LoadMesh
    Dim key As Variant, rc As Long, id As Long, k As Long
    Dim processed As Long, scanned As Long, progressTick As Double, stage As String
    Set ls = App.feLoadSet
    Set nt = App.feLoadNTemp
    Set et = App.feLoadETemp
    Set lm = App.feLoadMesh
    rc = App.feAppMessage(FCM_NORMAL, "STAGE 4/7: Check loads and explicit temperatures")
    rc = ls.First
    Do While rc = FE_OK
        processed = 0
        stage = "4/7 Temperature references, load set " & CStr(ls.ID)
        ReportProgress App, stage, 0, centers.Count + elements.Count, progressTick
        nt.SetID = ls.ID
        et.SetID = ls.ID
        lm.SetID = ls.ID
        For Each key In centers.Keys
            If nt.Get(CLng(key)) = FE_OK Then AddIssue issues, "Load set " & CStr(ls.ID) & _
                ": old center " & CStr(key) & " has an explicit nodal temperature; no automatic transfer."
            processed = processed + 1
            ReportProgress App, stage, processed, centers.Count + elements.Count, progressTick
        Next key
        For Each key In elements.Keys
            If et.Get(CLng(key)) = FE_OK Then AddIssue issues, "Load set " & CStr(ls.ID) & _
                ": deleted element " & CStr(key) & " has an explicit elemental temperature."
            processed = processed + 1
            ReportProgress App, stage, processed, centers.Count + elements.Count, progressTick
        Next key
        ' Scan non-temperature loads once per set, not once per selected node.
        scanned = 0
        stage = "4/7 Mesh loads scanned, load set " & CStr(ls.ID)
        ReportProgress App, stage, 0, 0, progressTick
        If lm.First = FE_OK Then
            Do
                id = lm.meshID
                If lm.type < 40 Then
                    If centers.Exists(CStr(id)) Then AddIssue issues, "Load set " & CStr(ls.ID) & _
                        ": old center " & CStr(id) & " has a nodal load/displacement (type " & CStr(lm.type) & ")."
                    For k = 0 To 1
                        If lm.addlID(k) > 0 Then
                            If centers.Exists(CStr(lm.addlID(k))) Then AddIssue issues, "Load set " & _
                                CStr(ls.ID) & ": load references old center " & CStr(lm.addlID(k)) & "."
                        End If
                    Next k
                    If masks.Exists(CStr(id)) And (lm.type = 3 Or lm.type = 4) Then
                        AddIssue issues, "Load set " & CStr(ls.ID) & ": prescribed displacement/rotation at new dependent node " & CStr(id) & "."
                    End If
                ElseIf elements.Exists(CStr(id)) Then
                    AddIssue issues, "Load set " & CStr(ls.ID) & ": elemental load on deleted element " & CStr(id) & "."
                End If
                scanned = scanned + 1
                ReportProgress App, stage, scanned, 0, progressTick
            Loop While lm.Next = FE_OK
        End If
        ReportProgress App, stage, scanned, scanned, progressTick
        rc = ls.Next
    Loop
End Sub

Sub ReportProgress(ByVal App As femap.model, ByVal stage As String, ByVal completed As Long, _
                   ByVal total As Long, ByRef lastTick As Double)
    Dim nowTick As Double, elapsed As Double, due As Boolean, rc As Long
    Dim countText As String
    nowTick = Timer
    elapsed = nowTick - lastTick
    If elapsed < 0# Then elapsed = elapsed + 86400#
    due = (completed = 0) Or (total > 0 And completed = total)
    If PROGRESS_EVERY_ITEMS > 0 Then
        If completed Mod PROGRESS_EVERY_ITEMS = 0 Then due = True
    End If
    If elapsed >= PROGRESS_INTERVAL_SECONDS Then due = True
    If due Then
        countText = CStr(completed) & " checked"
        If total > 0 Then countText = CStr(completed) & "/" & CStr(total)
        rc = App.feAppMessage(FCM_NORMAL, "PROGRESS " & stage & ": " & countText)
        lastTick = nowTick
    End If
End Sub

Sub PrepareRBE2(ByVal el As femap.Elem, ByVal center As Long, ByVal ids As Variant, _
                ByVal mask As Long, ByVal alpha As Double)
    Dim count As Long, k As Long, n As Long, rc As Long
    Dim nodeIDs() As Long, faces() As Long, flags() As Long, weights() As Double
    Dim vn As Variant, vf As Variant, vd As Variant, vw As Variant
    count = IDCount(ids)
    If count = 0 Then Err.Raise 513, "CBUSH to RBE2", "Cannot prepare an empty RBE2."
    ReDim nodeIDs(count - 1)
    ReDim faces(count - 1)
    ReDim weights(count - 1)
    ReDim flags(count * 6 - 1)
    For n = 0 To count - 1
        nodeIDs(n) = CLng(ids(LBound(ids) + n))
        weights(n) = 1#
        ' Per-node flags unused for ordinary RBE2; common DOFs are release(0,...).
    Next n
    el.type = FET_L_RIGID
    el.topology = 13
    el.propID = 0
    el.formulation(0) = 0
    el.formulation(1) = 0
    el.orientID = 0
    For k = 0 To 19
        el.node(k) = 0
    Next k
    el.node(0) = center
    For k = 0 To 5
        el.release(0, k) = ((mask And DOFBit(k)) <> 0)
        el.release(1, k) = False
    Next k
    el.RigidInterpolate = False
    el.RigidInterpolationFactor = 1#
    el.RigidThermalExpansion = alpha
    vn = nodeIDs
    vf = faces
    vw = weights
    vd = flags
    rc = el.PutNodeList(0, count, vn, vf, vw, vd)
    If rc <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot prepare the new RBE2 node list."
End Sub

Sub VerifyNew(ByVal App As femap.model, ByVal row As Variant)
    Dim nd As femap.Node, el As femap.Elem
    Dim ids As Variant, faces As Variant, weights As Variant, dofs As Variant
    Dim expected As Variant, mid As Variant, seen As Object
    Dim count As Long, n As Long, k As Long
    Set nd = App.feNode
    Set el = App.feElem
    If nd.Get(CLng(row(7))) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "New center missing."
    mid = row(2)
    If Not CloseValue(nd.x, CDbl(mid(0))) Or Not CloseValue(nd.y, CDbl(mid(1))) Or _
       Not CloseValue(nd.z, CDbl(mid(2))) Then Err.Raise 513, "CBUSH to RBE2", "Midpoint coordinate verification failed."
    If nd.defCSys <> 0 Or nd.outCSys <> 0 Or nd.type <> 0 Then Err.Raise 513, "CBUSH to RBE2", "New center metadata mismatch."
    If el.Get(CLng(row(8))) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "New RBE2 missing."
    If el.type <> FET_L_RIGID Or el.topology <> 13 Or el.RigidInterpolate Or _
       el.node(0) <> CLng(row(7)) Then Err.Raise 513, "CBUSH to RBE2", "New RBE2 type/reference verification failed."
    For k = 0 To 5
        If CBool(el.release(0, k)) <> ((CLng(row(5)) And DOFBit(k)) <> 0) Then
            Err.Raise 513, "CBUSH to RBE2", "New RBE2 DOF verification failed."
        End If
    Next k
    If Not CloseCTE(el.RigidThermalExpansion, CDbl(row(6))) Then Err.Raise 513, "CBUSH to RBE2", "New RBE2 CTE verification failed."
    If el.GetNodeList(0, count, ids, faces, weights, dofs) <> FE_OK Then Err.Raise 513, "CBUSH to RBE2", "Cannot verify new RBE2 node list."
    expected = row(3)
    If count <> IDCount(expected) Then Err.Raise 513, "CBUSH to RBE2", "New RBE2 peripheral count mismatch."
    Set seen = NewDictionary()
    For n = 0 To count - 1
        If seen.Exists(CStr(ids(LBound(ids) + n))) Then Err.Raise 513, "CBUSH to RBE2", "Duplicate new dependent node."
        seen.Add CStr(ids(LBound(ids) + n)), True
    Next n
    For n = LBound(expected) To UBound(expected)
        If Not seen.Exists(CStr(expected(n))) Then Err.Raise 513, "CBUSH to RBE2", "New RBE2 peripheral connectivity mismatch."
    Next n
End Sub

Function RestoreMesh(ByVal App As femap.model, ByVal oldNodes As Object, ByVal oldElements As Object, _
                     ByVal newNodes As Object, ByVal newElements As Object) As Boolean
    Dim nd As femap.Node, el As femap.Elem
    Dim originalNode As femap.Node, originalElement As femap.Elem
    Dim key As Variant, rc As Long, good As Boolean
    On Error Resume Next
    good = True
    Set nd = App.feNode
    Set el = App.feElem
    For Each key In newElements.Keys
        Err.Clear
        If el.Exist(CLng(key)) = FE_OK Then
            rc = 0
            rc = el.Delete(CLng(key))
            If rc <> FE_OK Or Err.Number <> 0 Then good = False
        End If
        If Err.Number <> 0 Then good = False
    Next key
    For Each key In oldNodes.Keys
        Err.Clear
        Set originalNode = oldNodes.Item(key)
        rc = 0
        rc = originalNode.Put(CLng(key))
        If rc <> FE_OK Or Err.Number <> 0 Then good = False
    Next key
    For Each key In oldElements.Keys
        Err.Clear
        Set originalElement = oldElements.Item(key)
        rc = 0
        rc = originalElement.Put(CLng(key))
        If rc <> FE_OK Or Err.Number <> 0 Then good = False
    Next key
    For Each key In newNodes.Keys
        Err.Clear
        If nd.Exist(CLng(key)) = FE_OK Then
            rc = 0
            rc = nd.Delete(CLng(key))
            If rc <> FE_OK Or Err.Number <> 0 Then good = False
        End If
        If Err.Number <> 0 Then good = False
    Next key
    For Each key In oldElements.Keys
        Err.Clear
        If el.Get(CLng(key)) <> FE_OK Then good = False
        Set originalElement = oldElements.Item(key)
        If Not SameElement(el, originalElement) Then good = False
        If Err.Number <> 0 Then good = False
    Next key
    For Each key In oldNodes.Keys
        Err.Clear
        If nd.Get(CLng(key)) <> FE_OK Then good = False
        Set originalNode = oldNodes.Item(key)
        If Not CloseValue(nd.x, originalNode.x) Or Not CloseValue(nd.y, originalNode.y) Or _
           Not CloseValue(nd.z, originalNode.z) Then good = False
        If nd.type <> originalNode.type Or nd.defCSys <> originalNode.defCSys Or _
           nd.outCSys <> originalNode.outCSys Or nd.superelementID <> originalNode.superelementID Then good = False
        If Err.Number <> 0 Then good = False
    Next key
    For Each key In newElements.Keys
        Err.Clear
        If el.Exist(CLng(key)) = FE_OK Then good = False
        If Err.Number <> 0 Then good = False
    Next key
    For Each key In newNodes.Keys
        Err.Clear
        If nd.Exist(CLng(key)) = FE_OK Then good = False
        If Err.Number <> 0 Then good = False
    Next key
    rc = App.feViewRegenerate(0)
    RestoreMesh = good
End Function

Function SameElement(ByVal a As femap.Elem, ByVal b As femap.Elem) As Boolean
    Dim n As Long, k As Long, na As Long, nb As Long
    Dim ia As Variant, fa As Variant, wa As Variant, da As Variant
    Dim ib As Variant, fb As Variant, wb As Variant, db As Variant
    SameElement = False
    If a.type <> b.type Or a.topology <> b.topology Or a.propID <> b.propID Or _
       a.orientID <> b.orientID Or a.layer <> b.layer Or a.color <> b.color Then Exit Function
    For n = 0 To 19
        If a.node(n) <> b.node(n) Then Exit Function
    Next n
    For n = 0 To 1
        If a.formulation(n) <> b.formulation(n) Then Exit Function
        For k = 0 To 5
            If a.release(n, k) <> b.release(n, k) Then Exit Function
        Next k
        For k = 0 To 2
            If Not CloseValue(a.offset(n, k), b.offset(n, k)) Then Exit Function
        Next k
    Next n
    For k = 0 To 2
        If Not CloseValue(a.orient(k), b.orient(k)) Then Exit Function
    Next k
    If a.type = FET_L_RIGID Then
        If a.RigidInterpolate <> b.RigidInterpolate Then Exit Function
        If Not CloseCTE(a.RigidThermalExpansion, b.RigidThermalExpansion) Then Exit Function
        If Not CloseValue(a.RigidInterpolationFactor, b.RigidInterpolationFactor) Then Exit Function
        If a.GetNodeList(0, na, ia, fa, wa, da) <> FE_OK Then Exit Function
        If b.GetNodeList(0, nb, ib, fb, wb, db) <> FE_OK Then Exit Function
        If na <> nb Then Exit Function
        For n = 0 To na - 1
            If CLng(ia(LBound(ia) + n)) <> CLng(ib(LBound(ib) + n)) Then Exit Function
            If Not CloseValue(CDbl(wa(LBound(wa) + n)), CDbl(wb(LBound(wb) + n))) Then Exit Function
            For k = 0 To 5
                If CLng(da(LBound(da) + n * 6 + k)) <> CLng(db(LBound(db) + n * 6 + k)) Then Exit Function
            Next k
        Next n
    End If
    SameElement = True
End Function

Function DOFBit(ByVal zeroBasedDOF As Long) As Long
    DOFBit = CLng(2 ^ zeroBasedDOF)
End Function

Function ParseDOFs(ByVal text As String) As Long
    Dim k As Long, d As Long, mask As Long, digit As String
    For k = 1 To Len(text)
        digit = Mid$(text, k, 1)
        If InStr("123456", digit) = 0 Then Err.Raise 513, "CBUSH to RBE2", "OUTPUT_DEPENDENT_DOFS must be empty or digits 1..6."
        d = CLng(digit) - 1
        If (mask And DOFBit(d)) <> 0 Then Err.Raise 513, "CBUSH to RBE2", "Duplicate DOF in OUTPUT_DEPENDENT_DOFS."
        mask = mask Or DOFBit(d)
    Next k
    ParseDOFs = mask
End Function

Function DOFText(ByVal mask As Long) As String
    Dim k As Long, text As String
    For k = 0 To 5
        If (mask And DOFBit(k)) <> 0 Then text = text & CStr(k + 1)
    Next k
    DOFText = text
End Function

Function JoinIDs(ByVal ids As Variant) As String
    Dim k As Long, text As String
    If IDCount(ids) = 0 Then
        JoinIDs = "none"
        Exit Function
    End If
    For k = LBound(ids) To UBound(ids)
        If Len(text) > 0 Then text = text & ","
        text = text & CStr(ids(k))
    Next k
    JoinIDs = text
End Function

Function IDCount(ByVal ids As Variant) As Long
    On Error GoTo EmptyIDs
    IDCount = UBound(ids) - LBound(ids) + 1
    Exit Function
EmptyIDs:
    IDCount = 0
End Function

Function CloseValue(ByVal a As Double, ByVal b As Double) As Boolean
    CloseValue = (Abs(a - b) <= 0.0000000001 * (1# + Abs(a) + Abs(b)))
End Function

Function CloseCTE(ByVal a As Double, ByVal b As Double) As Boolean
    CloseCTE = (Abs(a - b) <= 0.000000000001 * (1# + Abs(a) + Abs(b)))
End Function
