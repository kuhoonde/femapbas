Option Explicit

' Rebuild the spider node list of selected RBE2/RBE3 elements.
' Preferred: 2x2 patch of four adjacent quad solid faces near the center.
' Use every distinct node on the patch (usually 9 corners, plus midsides if present).
' Fallback: nearest single face if no valid, conflict-free 2x2 patch exists.
' Candidates = actual mesh face nodes on the selected NEW Solid CAD surfaces.
' FAST: read selected surface faces and index all rigid node ownership once.
' No repeated whole-model FGD_ELEM_BYNODE queries.
' ALL old spider ends are replaced; center and element ID stay unchanged.
' No radius or node-count input. Patch topology determines node count.
' Select only spiders whose entire attachment is to be rebuilt.
' Standard RBE3 only (no UM DOFs; uniform independent-node weights/DOFs).
' No node creation/deletion. Same element ID and reference node.
' 이미 나사가 체결되어있었지만 여러이유로 기존의 솔리드가 지우고 유사한형상의 솔리드가 들어올때. 중심노드를 기준으로 2x2페이스를 찾아 자동연결
Sub Main
    Dim App As femap.model
    Set App = GetObject(, "femap.model")
    Dim es As femap.Set, ss As femap.Set, ns As femap.Set
    Dim el As femap.Elem, chk As femap.Elem
    Dim sf As femap.Surface, nd As femap.Node
    Dim allRigids As femap.Set, rNodes As femap.Set, owner As Object
    Set es = App.feSet
    Set ss = App.feSet
    Set ns = App.feSet
    Set el = App.feElem
    Set chk = App.feElem
    Set sf = App.feSurface
    Set nd = App.feNode
    Set allRigids = App.feSet
    Set rNodes = App.feSet
    Set owner = CreateObject("Scripting.Dictionary")

    Dim rc As Long, eid As Long, nid As Long, ne As Long, nt As Long
    Dim i As Long, j As Long, k As Long, p As Long, m As Long, center As Long
    Dim oldN As Long, newN As Long, planCount As Long
    Dim skipped As Long, written As Long, planned As Long, state As Long
    Dim d2 As Double, dx As Double, dy As Double, dz As Double
    Dim cx As Double, cy As Double, cz As Double, maxD2 As Double
    Dim s As String, reason As String, elemLabel As String, coord As Variant
    Dim isRbe3 As Boolean, asked As Boolean, allowRbe3 As Boolean
    Dim vNodes As Variant, vFace As Variant, vWeight As Variant, vDOF As Variant
    Dim ids() As Long, px() As Double, py() As Double, pz() As Double
    Dim reserved() As Long, picked() As Long, newNodes() As Long
    Dim planEid() As Long, planCenter() As Long, planType() As Boolean
    Dim planNodes() As Variant, planOld() As Variant, planN() As Long
    Dim planOldN() As Long, planDof() As Variant, planWeight() As Double
    Dim newFace() As Long, newWeight() As Double, newDOF() As Long
    Dim nv As Variant, fv As Variant, wv As Variant, dv As Variant, expectedOld As Variant
    Dim oldDof As Variant, rbNodes As Variant, rbFace As Variant, rbWeight As Variant, rbDof As Variant
    Dim rbCount As Long, removedConnected As Long, warnCount As Long
    Dim uniformWeight As Double, sameNode As Boolean, checkIndex As Long
    Dim sid As Long, nf As Long, f As Long, q As Long, rid As Long
    Dim faceElems As Variant, faceIds As Variant, fn As Variant
    Dim coordIds As Variant, coordXYZ As Variant, coordExists As Variant
    Dim ownerId As Long, nRigid As Long, readRigids As Long, key As String
    Dim faceTotal As Long, faceBase As Long, fi As Long, bestFace As Long, pointCount As Long
    Dim faceNodes() As Long, faceCount() As Long, faceElement() As Long, faceNumber() As Long
    Dim faceSurface() As Long, firstInc() As Long, incFace() As Long, nextInc() As Long
    Dim incN As Long, incPtr As Long, patchN As Long, trialN As Long, v As Long, u As Long
    Dim patchFaces(0 To 3) As Long, trialFaces(0 To 3) As Long
    Dim patchMiddle As Long, trialMiddle As Long, patchScore As Double, trialScore As Double
    Dim trialOK As Boolean
    Dim bestFaceD2 As Double, tieD2 As Double, centroidD2 As Double, tx As Double, ty As Double, tz As Double
    Dim a As Long, b As Long, c As Long, d As Long, corners As Long, conflict As Boolean
    Dim coordIndex As Object, idx As Long
    Set coordIndex = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed

    If es.Select(FT_ELEM, True, "Select RBE2/RBE3 spiders to rebuild") <> FE_OK Then Exit Sub
    If es.Count = 0 Then Exit Sub
    If ss.Select(FT_SURFACE, True, "Select NEW solid bolt attachment surfaces ONLY") <> FE_OK Then Exit Sub
    If ss.Count = 0 Then Exit Sub

    rc = App.feAppMessage(FCM_NORMAL, "--- 2x2 QUAD PATCH: collecting selected solid faces ---")
    rc = ss.Reset
    sid = ss.Next
    Do While sid > 0
        rc = App.feAppMessage(FCM_NORMAL, "Reading Surface " & CStr(sid) & " solid element faces...")
        If sf.Get(sid) <> FE_OK Then
            MsgBox "Cannot read selected surface " & CStr(sid), vbExclamation
            Exit Sub
        End If
        ' Only actual solid element faces; exclude shell faces and orphan nodes.
        nf = 0
        rc = sf.ElementFaces(False, True, False, nf, faceElems, faceIds)
        If rc <> FE_OK Or nf = 0 Then
            MsgBox "No solid element faces on Surface " & CStr(sid) & ". Check mesh/geometry association.", vbExclamation
            Exit Sub
        End If
        rc = App.feAppMessage(FCM_NORMAL, "  " & CStr(nf) & " solid faces; extracting face nodes...")
        faceBase = faceTotal
        faceTotal = faceTotal + nf
        If faceBase = 0 Then
            ReDim faceNodes(0 To 7, 0 To faceTotal - 1)
            ReDim faceCount(0 To faceTotal - 1)
            ReDim faceElement(0 To faceTotal - 1)
            ReDim faceNumber(0 To faceTotal - 1)
            ReDim faceSurface(0 To faceTotal - 1)
        Else
            ReDim Preserve faceNodes(0 To 7, 0 To faceTotal - 1)
            ReDim Preserve faceCount(0 To faceTotal - 1)
            ReDim Preserve faceElement(0 To faceTotal - 1)
            ReDim Preserve faceNumber(0 To faceTotal - 1)
            ReDim Preserve faceSurface(0 To faceTotal - 1)
        End If
        For f = 0 To nf - 1
            fi = faceBase + f
            faceElement(fi) = CLng(faceElems(f))
            faceNumber(fi) = CLng(faceIds(f))
            faceSurface(fi) = sid
            pointCount = 0
            If chk.Get(CLng(faceElems(f))) <> FE_OK Then
                MsgBox "Cannot read solid element " & CStr(faceElems(f)), vbExclamation
                Exit Sub
            End If
            rc = chk.GetFaceNodes(CLng(faceIds(f)), fn)
            If rc <> FE_OK Then
                MsgBox "Cannot read solid face nodes for element " & CStr(faceElems(f)), vbExclamation
                Exit Sub
            End If
            For q = 0 To 7
                nid = CLng(fn(q))
                If nid > 0 Then
                    rc = ns.Add(nid)
                    faceNodes(pointCount, fi) = nid
                    pointCount = pointCount + 1
                End If
            Next q
            faceCount(fi) = pointCount
            If pointCount <> 3 And pointCount <> 4 And pointCount <> 6 And pointCount <> 8 Then
                MsgBox "Unsupported solid face node count: " & CStr(pointCount), vbExclamation
                Exit Sub
            End If
            If f > 0 And f Mod 2000 = 0 Then rc = App.feAppMessage(FCM_NORMAL, "  Faces read: " & CStr(f) & "/" & CStr(nf))
        Next f
        sid = ss.Next
    Loop
    nt = ns.Count
    If nt < 3 Then
        MsgBox "Fewer than 3 solid face nodes.", vbExclamation
        Exit Sub
    End If
    rc = App.feAppMessage(FCM_NORMAL, "Collecting XYZ for " & CStr(nt) & " surface nodes...")
    rc = nd.GetCoordArray2(ns.ID, nt, coordIds, coordExists, coordXYZ)
    If rc <> FE_OK Or nt < 3 Then
        MsgBox "Cannot read surface coordinates.", vbExclamation
        Exit Sub
    End If
    ReDim ids(0 To nt - 1)
    ReDim px(0 To nt - 1)
    ReDim py(0 To nt - 1)
    ReDim pz(0 To nt - 1)
    ReDim reserved(0 To nt - 1)
    ReDim picked(0 To nt - 1)
    For p = 0 To nt - 1
        If Not CBool(coordExists(p)) Then
            MsgBox "Surface node missing: " & CStr(coordIds(p)), vbExclamation
            Exit Sub
        End If
        ids(p) = CLng(coordIds(p))
        coordIndex.Add CStr(ids(p)), p
        px(p) = CDbl(coordXYZ(3 * p))
        py(p) = CDbl(coordXYZ(3 * p + 1))
        pz(p) = CDbl(coordXYZ(3 * p + 2))
    Next p
    ' Convert face connectivity from node IDs into cached coordinate indexes.
    For fi = 0 To faceTotal - 1
        For q = 0 To faceCount(fi) - 1
            key = CStr(faceNodes(q, fi))
            If Not coordIndex.Exists(key) Then
                MsgBox "Face node missing from coordinate set: " & key, vbExclamation
                Exit Sub
            End If
            faceNodes(q, fi) = CLng(coordIndex.Item(key))
        Next q
    Next fi
    rc = App.feAppMessage(FCM_NORMAL, "Surface coordinates ready: " & CStr(nt) & " nodes; " & CStr(faceTotal) & " faces.")

    ' Build a fast corner-to-quad-face incidence index once.
    ReDim firstInc(0 To nt - 1)
    ReDim incFace(1 To 4 * faceTotal)
    ReDim nextInc(1 To 4 * faceTotal)
    For fi = 0 To faceTotal - 1
        If faceCount(fi) = 4 Or faceCount(fi) = 8 Then
            For q = 0 To 3
                idx = faceNodes(q, fi)
                incN = incN + 1
                incFace(incN) = fi
                nextInc(incN) = firstInc(idx)
                firstInc(idx) = incN
            Next q
        End If
    Next fi
    rc = App.feAppMessage(FCM_NORMAL, "Quad adjacency ready.")

    ' Index existing rigid connectivity once, rather than asking FEMAP for
    ' every element connected to each candidate node.
    rc = App.feAppMessage(FCM_NORMAL, "Indexing existing rigid elements (one pass)...")
    rc = allRigids.AddRule(29, FGD_ELEM_BYTYPE)
    If rc <> FE_OK Or allRigids.Count = 0 Then
        MsgBox "Could not index existing rigid elements. No changes made.", vbExclamation
        Exit Sub
    End If
    nRigid = allRigids.Count
    rc = allRigids.Reset
    rid = allRigids.Next
    Do While rid > 0
        rc = rNodes.Clear
        rc = rNodes.AddRule(rid, FGD_NODE_ONELEM)
        If rc <> FE_OK Or rNodes.Count = 0 Then
            MsgBox "Cannot index RBE " & CStr(rid) & ". No changes made.", vbExclamation
            Exit Sub
        End If
        rc = rNodes.Reset
        nid = rNodes.Next
        Do While nid > 0
            key = CStr(nid)
            If owner.Exists(key) Then
                If CLng(owner.Item(key)) <> rid Then owner.Item(key) = -1
            Else
                owner.Add key, rid
            End If
            nid = rNodes.Next
        Loop
        readRigids = readRigids + 1
        If readRigids Mod 100 = 0 Then rc = App.feAppMessage(FCM_NORMAL, "  Rigids indexed: " & CStr(readRigids) & "/" & CStr(nRigid))
        rid = allRigids.Next
    Loop
    rc = App.feAppMessage(FCM_NORMAL, "Rigid index ready: " & CStr(readRigids) & " elements.")

    ne = es.Count
    ReDim planEid(0 To ne - 1)
    ReDim planCenter(0 To ne - 1)
    ReDim planType(0 To ne - 1)
    ReDim planNodes(0 To ne - 1)
    ReDim planOld(0 To ne - 1)
    ReDim planN(0 To ne - 1)
    ReDim planOldN(0 To ne - 1)
    ReDim planDof(0 To ne - 1)
    ReDim planWeight(0 To ne - 1)
    rc = App.feAppMessage(FCM_NORMAL, "--- 2x2 QUAD PATCH PREVIEW (no model changes) ---")
    rc = App.feAppMessage(FCM_NORMAL, "Preferred: four connected quad faces (2x2); fallback: one nearest face. No radius or count.")

    rc = es.Reset
    eid = es.Next
    Do While eid > 0
        reason = ""
        rc = App.feAppMessage(FCM_NORMAL, "Finding nearest 2x2 quad patch for RBE " & CStr(eid) & "...")
        newN = 0
        removedConnected = 0
        maxD2 = 0#
        If el.Get(eid) <> FE_OK Then
            reason = "Cannot read element"
            GoTo SkipElem
        End If
        If el.Type <> 29 Or el.topology <> 13 Then
            reason = "Not a standard RBE2/RBE3"
            GoTo SkipElem
        End If
        isRbe3 = el.RigidInterpolate
        elemLabel = "RBE2"
        If isRbe3 Then
            elemLabel = "RBE3"
            If Not asked Then
                allowRbe3 = (MsgBox("Confirm selected RBE3 elements use NO custom UM DOFs." & vbCrLf & _
                    "No skips RBE3. Non-uniform weights/DOFs are skipped automatically.", _
                    vbYesNo Or vbQuestion Or vbDefaultButton2, "RBE3 safety check") = vbYes)
                asked = True
            End If
            If Not allowRbe3 Then
                reason = "RBE3 without UM not confirmed"
                GoTo SkipElem
            End If
        End If
        center = el.node(0)
        If center <= 0 Then
            reason = "Missing center/reference node"
            GoTo SkipElem
        End If
        If App.feCoordOnNode(center, coord) <> FE_OK Then
            reason = "Cannot read center/reference coordinates"
            GoTo SkipElem
        End If
        cx = CDbl(coord(0))
        cy = CDbl(coord(1))
        cz = CDbl(coord(2))
        rc = el.GetNodeList(0, oldN, vNodes, vFace, vWeight, vDOF)
        If rc <> FE_OK Or oldN < 1 Then
            reason = "No readable old spider list"
            GoTo SkipElem
        End If
        uniformWeight = 0#
        If isRbe3 Then uniformWeight = CDbl(vWeight(0))
        ' Changing node count requires a common per-node DOF/weight template.
        For i = 0 To oldN - 1
            If isRbe3 Then
                If Abs(CDbl(vWeight(i)) - uniformWeight) > 0.000000000001 * (1# + Abs(uniformWeight)) Then
                    reason = "Non-uniform RBE3 weights; needs manual redistribution"
                    GoTo SkipElem
                End If
            End If
            For k = 0 To 5
                If CLng(vDOF(6 * i + k)) <> CLng(vDOF(k)) Then
                    reason = "Non-uniform spider DOFs; needs manual redistribution"
                    GoTo SkipElem
                End If
            Next k
        Next i

        ' Find closest geometric face, as triangles (flat approximation for
        ' curved midside geometry). Break edge/corner ties by corner centroid.
        bestFace = -1
        bestFaceD2 = 1.0E+30
        tieD2 = 1.0E+30
        For fi = 0 To faceTotal - 1
            a = faceNodes(0, fi)
            b = faceNodes(1, fi)
            c = faceNodes(2, fi)
            corners = 3
            If faceCount(fi) = 4 Or faceCount(fi) = 8 Then corners = 4
            d2 = TriDist2(cx, cy, cz, a, b, c, px, py, pz)
            If corners = 4 Then
                d = faceNodes(3, fi)
                dx = TriDist2(cx, cy, cz, a, c, d, px, py, pz)
                If dx < d2 Then d2 = dx
            End If
            If d2 <= bestFaceD2 + 0.0000000001 * (1# + bestFaceD2) Then
                tx = 0#: ty = 0#: tz = 0#
                For q = 0 To corners - 1
                    idx = faceNodes(q, fi)
                    tx = tx + px(idx)
                    ty = ty + py(idx)
                    tz = tz + pz(idx)
                Next q
                centroidD2 = (tx / corners - cx) ^ 2 + (ty / corners - cy) ^ 2 + (tz / corners - cz) ^ 2
                If bestFace < 0 Or d2 < bestFaceD2 - 0.0000000001 * (1# + bestFaceD2) Or centroidD2 < tieD2 Then
                    bestFace = fi
                    bestFaceD2 = d2
                    tieD2 = centroidD2
                End If
            End If
        Next fi
        If bestFace < 0 Then
            reason = "No supported nearest solid face"
            GoTo SkipElem
        End If

        ' A 2x2 quad patch contains the nearest face plus the three other
        ' quads incident to a shared interior corner. All four must be on
        ' the same selected CAD surface. Choose the shared corner closest
        ' to the RBE reference, among conflict-free valid patches.
        patchN = 1
        patchFaces(0) = bestFace
        patchMiddle = -1
        patchScore = 1.0E+30
        If faceCount(bestFace) = 4 Or faceCount(bestFace) = 8 Then
            For v = 0 To 3
                trialMiddle = faceNodes(v, bestFace)
                incPtr = firstInc(trialMiddle)
                trialN = 0
                Do While incPtr > 0
                    fi = incFace(incPtr)
                    If faceSurface(fi) = faceSurface(bestFace) Then
                        If trialN < 4 Then trialFaces(trialN) = fi
                        trialN = trialN + 1
                    End If
                    incPtr = nextInc(incPtr)
                Loop
                If trialN = 4 Then
                    If ValidQuadPatch(trialFaces, faceNodes, trialMiddle) Then
                        trialOK = True
                        For u = 0 To 3
                            fi = trialFaces(u)
                            For q = 0 To faceCount(fi) - 1
                                j = faceNodes(q, fi)
                                If ids(j) <> center Then
                                    ownerId = 0
                                    key = CStr(ids(j))
                                    If owner.Exists(key) Then ownerId = CLng(owner.Item(key))
                                    If reserved(j) <> 0 Or (ownerId <> 0 And ownerId <> eid) Then trialOK = False
                                End If
                            Next q
                        Next u
                        If trialOK Then
                            trialScore = (px(trialMiddle) - cx) ^ 2 + (py(trialMiddle) - cy) ^ 2 + (pz(trialMiddle) - cz) ^ 2
                            If trialScore < patchScore Then
                                patchScore = trialScore
                                patchN = 4
                                patchMiddle = trialMiddle
                                For u = 0 To 3
                                    patchFaces(u) = trialFaces(u)
                                Next u
                            End If
                        End If
                    End If
                End If
            Next v
        End If

        ' Use every distinct node of the chosen four-face patch, or single
        ' face fallback. Keep the center unchanged and never self-connect.
        conflict = False
        For u = 0 To patchN - 1
            fi = patchFaces(u)
            For q = 0 To faceCount(fi) - 1
                j = faceNodes(q, fi)
                If ids(j) <> center Then
                    ownerId = 0
                    key = CStr(ids(j))
                    If owner.Exists(key) Then ownerId = CLng(owner.Item(key))
                    If reserved(j) <> 0 Or (ownerId <> 0 And ownerId <> eid) Then
                        conflict = True
                    Else
                        sameNode = False
                        For k = 0 To newN - 1
                            If picked(k) = j Then sameNode = True
                        Next k
                        If Not sameNode Then
                            picked(newN) = j
                            newN = newN + 1
                            d2 = (px(j) - cx) ^ 2 + (py(j) - cy) ^ 2 + (pz(j) - cz) ^ 2
                            If d2 > maxD2 Then maxD2 = d2
                        End If
                    End If
                End If
            Next q
        Next u
        If conflict Then
            reason = "Nearest patch contains nodes used by another RBE. Check manually."
            GoTo SkipElem
        End If
        If newN < 1 Then
            reason = "No eligible node on nearest patch"
            GoTo SkipElem
        End If
        If isRbe3 Then
            If Not NonCollinear(picked, newN, px, py, pz) Then
                reason = "Nearest patch does not provide 3 non-collinear RBE3 targets"
                GoTo SkipElem
            End If
        End If

        ReDim newNodes(0 To newN - 1)
        For i = 0 To newN - 1
            newNodes(i) = ids(picked(i))
        Next i
        ' Note every old leg removed. Do not scan whole-model connectivity.
        ' Some may still be attached to another part; user must inspect preview.
        For i = 0 To oldN - 1
            sameNode = False
            For j = 0 To newN - 1
                If CLng(vNodes(i)) = newNodes(j) Then sameNode = True
            Next j
            If Not sameNode Then
                removedConnected = removedConnected + 1
            End If
        Next i
        If removedConnected > 0 Then
            warnCount = warnCount + removedConnected
            rc = App.feAppMessage(FCM_NORMAL, "WARNING " & CStr(eid) & ": " & CStr(removedConnected) & _
                " old entries will be removed (may include still-connected nodes).")
        End If
        For i = 0 To newN - 1
            reserved(picked(i)) = eid
        Next i
        nv = newNodes
        planEid(planCount) = eid
        planCenter(planCount) = center
        planType(planCount) = isRbe3
        planNodes(planCount) = nv
        planOld(planCount) = vNodes
        planN(planCount) = newN
        planOldN(planCount) = oldN
        planDof(planCount) = vDOF
        planWeight(planCount) = uniformWeight
        planCount = planCount + 1
        planned = planned + newN
        rc = App.feAppMessage(FCM_NORMAL, elemLabel & " " & CStr(eid) & ", center " & CStr(center) & _
            ": old " & CStr(oldN) & " -> new " & CStr(newN) & "; faces " & CStr(patchN) & "; nearest face " & CStr(faceElement(bestFace)) & "/" & CStr(faceNumber(bestFace)) & _
            "; face distance " & CStr(Sqr(bestFaceD2)) & "; farthest node " & CStr(Sqr(maxD2)))
        If patchN = 4 Then
            rc = App.feAppMessage(FCM_NORMAL, "  2x2 quad patch: 4 connected faces; middle mesh node " & CStr(ids(patchMiddle)))
        Else
            rc = App.feAppMessage(FCM_NORMAL, "  FALLBACK: 1 face (no valid conflict-free 2x2 block at nearest face).")
        End If
        GoTo NextElem
SkipElem:
        skipped = skipped + 1
        rc = App.feAppMessage(FCM_NORMAL, "SKIP " & CStr(eid) & ": " & reason)
NextElem:
        eid = es.Next
    Loop

    If planCount = 0 Then
        MsgBox "No spiders can be rebuilt. See FEMAP Messages for reasons.", vbInformation
        Exit Sub
    End If
    If MsgBox("Automatic 2x2 quad face patch (one-face fallback)." & vbCrLf & "RBE2/RBE3 to update: " & CStr(planCount) & vbCrLf & _
        "New spider entries: " & CStr(planned) & vbCrLf & "Skipped: " & CStr(skipped) & vbCrLf & _
        "Old entries being removed: " & CStr(warnCount) & vbCrLf & vbCrLf & _
        "ALL old spider ends of each listed RBE will be replaced. Other-side connections may be lost." & vbCrLf & _
        "Check FEMAP Messages. Apply changes?", vbYesNo Or vbQuestion Or vbDefaultButton2, _
        "Confirm 2x2 patch rebuild") <> vbYes Then Exit Sub

    For m = 0 To planCount - 1
        eid = planEid(m)
        If el.Get(eid) <> FE_OK Then GoTo WriteFailed
        If el.Type <> 29 Or el.topology <> 13 Then GoTo WriteFailed
        If CBool(el.RigidInterpolate) <> planType(m) Then GoTo WriteFailed
        If el.node(0) <> planCenter(m) Then GoTo WriteFailed
        rc = el.GetNodeList(0, oldN, vNodes, vFace, vWeight, vDOF)
        If rc <> FE_OK Or oldN <> planOldN(m) Then GoTo WriteFailed
        oldDof = planDof(m)
        expectedOld = planOld(m)
        For i = 0 To oldN - 1
            If CLng(vNodes(i)) <> CLng(expectedOld(i)) Then GoTo WriteFailed
            If planType(m) Then
                If Abs(CDbl(vWeight(i)) - planWeight(m)) > 0.000000000001 * (1# + Abs(planWeight(m))) Then GoTo WriteFailed
            End If
            For k = 0 To 5
                If CLng(vDOF(6 * i + k)) <> CLng(oldDof(6 * i + k)) Then GoTo WriteFailed
            Next k
        Next i
        newN = planN(m)
        nv = planNodes(m)
        ReDim newFace(0 To newN - 1)
        ReDim newWeight(0 To newN - 1)
        ReDim newDOF(0 To 6 * newN - 1)
        For i = 0 To newN - 1
            newFace(i) = 0
            newWeight(i) = planWeight(m)
            For k = 0 To 5
                newDOF(6 * i + k) = CLng(oldDof(k))
            Next k
        Next i
        fv = newFace
        wv = newWeight
        dv = newDOF
        rc = el.PutNodeList(0, newN, nv, fv, wv, dv)
        If rc <> FE_OK Then GoTo WriteFailed
        rc = el.Put(eid)
        If rc <> FE_OK Then GoTo WriteFailed
        written = written + 1
        If el.Get(eid) <> FE_OK Then GoTo WriteFailed
        If el.node(0) <> planCenter(m) Or CBool(el.RigidInterpolate) <> planType(m) Then GoTo WriteFailed
        rc = el.GetNodeList(0, rbCount, rbNodes, rbFace, rbWeight, rbDof)
        If rc <> FE_OK Or rbCount <> newN Then GoTo WriteFailed
        ' Order-independent readback: same IDs and preserved uniform settings.
        For i = 0 To rbCount - 1
            checkIndex = -1
            For j = 0 To newN - 1
                If CLng(rbNodes(i)) = CLng(nv(j)) Then checkIndex = j
            Next j
            If checkIndex < 0 Then GoTo WriteFailed
            If planType(m) Then
                If Abs(CDbl(rbWeight(i)) - planWeight(m)) > 0.000000000001 * (1# + Abs(planWeight(m))) Then GoTo WriteFailed
            End If
            For k = 0 To 5
                If CLng(rbDof(6 * i + k)) <> CLng(oldDof(k)) Then GoTo WriteFailed
            Next k
        Next i
        rc = App.feAppMessage(FCM_NORMAL, "UPDATED " & CStr(eid) & " with " & CStr(newN) & " nodes")
    Next m
    rc = App.feViewRegenerate(0)
    MsgBox CStr(written) & " spiders updated. " & CStr(skipped) & " skipped." & vbCrLf & _
        "Old floating nodes were not deleted. Inspect connections and run MPC/model checks.", vbInformation
    Exit Sub

WriteFailed:
    rc = App.feAppMessage(FCM_NORMAL, "STOP: write/readback failed at element " & CStr(eid))
    MsgBox "Stopped at RBE " & CStr(eid) & ". Writes completed: " & CStr(written) & _
        ". The current element may have changed. Inspect or restore backup.", vbCritical
    Exit Sub
Failed:
    MsgBox "API error " & CStr(Err.Number) & ": " & Err.Description & vbCrLf & _
        "Writes completed: " & CStr(written) & ". Check model backup.", vbCritical
End Sub

' Conservative geometric check: at least 3 non-collinear target locations.
Private Function NonCollinear(ByRef picked() As Long, ByVal n As Long, _
    ByRef x() As Double, ByRef y() As Double, ByRef z() As Double) As Boolean
    Dim i As Long, a As Long, b As Long, c As Long
    Dim ux As Double, uy As Double, uz As Double, vx As Double, vy As Double, vz As Double
    Dim d As Double, longest As Double, crossX As Double, crossY As Double, crossZ As Double
    NonCollinear = False
    If n < 3 Then Exit Function
    a = picked(0)
    For i = 1 To n - 1
        c = picked(i)
        d = (x(c) - x(a)) ^ 2 + (y(c) - y(a)) ^ 2 + (z(c) - z(a)) ^ 2
        If d > longest Then
            longest = d
            b = c
        End If
    Next i
    If longest <= 0# Then Exit Function
    ux = x(b) - x(a)
    uy = y(b) - y(a)
    uz = z(b) - z(a)
    For i = 1 To n - 1
        c = picked(i)
        vx = x(c) - x(a)
        vy = y(c) - y(a)
        vz = z(c) - z(a)
        crossX = uy * vz - uz * vy
        crossY = uz * vx - ux * vz
        crossZ = ux * vy - uy * vx
        If crossX * crossX + crossY * crossY + crossZ * crossZ > 0.000000000001 * longest * longest Then
            NonCollinear = True
            Exit Function
        End If
    Next i
End Function

' Squared distance from point to a triangle: projection if inside,
' otherwise the minimum of the three point-to-edge distances.
Private Function TriDist2(ByVal cx As Double, ByVal cy As Double, ByVal cz As Double, _
    ByVal a As Long, ByVal b As Long, ByVal c As Long, _
    ByRef x() As Double, ByRef y() As Double, ByRef z() As Double) As Double
    Dim ux As Double, uy As Double, uz As Double, vx As Double, vy As Double, vz As Double
    Dim wx As Double, wy As Double, wz As Double, uu As Double, uv As Double, vv As Double
    Dim wu As Double, wv As Double, den As Double, alpha As Double, beta As Double
    Dim dx As Double, dy As Double, dz As Double, best As Double, dd As Double
    ux = x(b) - x(a): uy = y(b) - y(a): uz = z(b) - z(a)
    vx = x(c) - x(a): vy = y(c) - y(a): vz = z(c) - z(a)
    wx = cx - x(a): wy = cy - y(a): wz = cz - z(a)
    uu = ux * ux + uy * uy + uz * uz
    uv = ux * vx + uy * vy + uz * vz
    vv = vx * vx + vy * vy + vz * vz
    wu = wx * ux + wy * uy + wz * uz
    wv = wx * vx + wy * vy + wz * vz
    den = uu * vv - uv * uv
    If den > 0.00000000000001 * uu * vv Then
        alpha = (wu * vv - wv * uv) / den
        beta = (wv * uu - wu * uv) / den
        If alpha >= 0# And beta >= 0# And alpha + beta <= 1# Then
            dx = wx - alpha * ux - beta * vx
            dy = wy - alpha * uy - beta * vy
            dz = wz - alpha * uz - beta * vz
            TriDist2 = dx * dx + dy * dy + dz * dz
            Exit Function
        End If
    End If
    best = EdgeDist2(cx, cy, cz, a, b, x, y, z)
    dd = EdgeDist2(cx, cy, cz, b, c, x, y, z)
    If dd < best Then best = dd
    dd = EdgeDist2(cx, cy, cz, c, a, x, y, z)
    If dd < best Then best = dd
    TriDist2 = best
End Function

Private Function EdgeDist2(ByVal cx As Double, ByVal cy As Double, ByVal cz As Double, _
    ByVal a As Long, ByVal b As Long, ByRef x() As Double, ByRef y() As Double, _
    ByRef z() As Double) As Double
    Dim ux As Double, uy As Double, uz As Double, wx As Double, wy As Double, wz As Double
    Dim len2 As Double, t As Double, dx As Double, dy As Double, dz As Double
    ux = x(b) - x(a): uy = y(b) - y(a): uz = z(b) - z(a)
    wx = cx - x(a): wy = cy - y(a): wz = cz - z(a)
    len2 = ux * ux + uy * uy + uz * uz
    t = 0#
    If len2 > 0# Then t = (wx * ux + wy * uy + wz * uz) / len2
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    dx = wx - t * ux: dy = wy - t * uy: dz = wz - t * uz
    EdgeDist2 = dx * dx + dy * dy + dz * dz
End Function

' True only for a topological 2x2 quad patch: 9 distinct corner nodes,
' one central corner in all four faces, four edge corners in two faces,
' four outer corners in one face. Also verify four actual internal edges.
Private Function ValidQuadPatch(ByRef f() As Long, ByRef fn() As Long, ByVal middle As Long) As Boolean
    Dim unique(0 To 15) As Long, hits(0 To 15) As Long
    Dim radial(0 To 7) As Long, radialHits(0 To 7) As Long
    Dim n As Long, rn As Long, i As Long, j As Long, k As Long, p As Long, loc As Long
    Dim id As Long, centerHits As Long, twice As Long, once As Long, r As Long
    ValidQuadPatch = False
    For i = 0 To 3
        loc = -1
        For j = 0 To 3
            id = fn(j, f(i))
            If id = middle Then loc = j
            p = -1
            For k = 0 To n - 1
                If unique(k) = id Then p = k
            Next k
            If p < 0 Then
                p = n
                unique(p) = id
                n = n + 1
            End If
            hits(p) = hits(p) + 1
        Next j
        If loc < 0 Then Exit Function
        For r = 1 To 3 Step 2
            id = fn((loc + r) Mod 4, f(i))
            If id = middle Then Exit Function
            p = -1
            For k = 0 To rn - 1
                If radial(k) = id Then p = k
            Next k
            If p < 0 Then
                p = rn
                radial(p) = id
                rn = rn + 1
            End If
            radialHits(p) = radialHits(p) + 1
        Next r
    Next i
    If n <> 9 Or rn <> 4 Then Exit Function
    For i = 0 To n - 1
        If unique(i) = middle Then
            centerHits = hits(i)
        ElseIf hits(i) = 2 Then
            twice = twice + 1
        ElseIf hits(i) = 1 Then
            once = once + 1
        Else
            Exit Function
        End If
    Next i
    If centerHits <> 4 Or twice <> 4 Or once <> 4 Then Exit Function
    For i = 0 To rn - 1
        If radialHits(i) <> 2 Then Exit Function
        For j = 0 To n - 1
            If unique(j) = radial(i) Then
                If hits(j) <> 2 Then Exit Function
            End If
        Next j
    Next i
    ValidQuadPatch = True
End Function
