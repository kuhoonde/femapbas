Option Explicit

' FEMAP WinWrap Basic - v1 experimental cylindrical glue array.
' Creates independent RBE2 centers at identical XYZ, RBE2/CBUSH/RBE2.
' K1/K4 point inward; K2/K5 tangential; K3/K6 along the cylinder axis.
' Existing six-DOF CBUSH Property must be created before running.
' Select ONLY meshed coaxial cylindrical attachment CAD surfaces.
' First glue center: pick a point or enter XYZ (sets height and angle).
' Glue diameter is in model units; surface distance uses arc length and height.
' Preview is read-only. No existing mesh, property or elements are edited.
' Never equivalence/merge the two coincident center nodes.
' 1. 글루바깥족선택(링-렌즈사이면 링내부)
' 2. 글루안쪽면선택(렌즈표면)
' 3. 글루 패드의 중심점선택(링의 글루홀이나 렌즈 표면의 글루중심점 등 상관없음. 어차피 자동으로 글루체적의 중심으로 zerogap coincident만들어줌. 단순히 어디에 몇개를 배치할지 확인하기 위한 시작점)
' 4. 글루 갯수입력
' 5. 글루 직경(반지름아님, mks면 m로, mm 단위계면 mm로 입력해야함)
' 6. cbush 프로퍼티 선택
' 7. yes 
' 8. 좌표계 포함 자동생성. 손계산을 통해 렌즈 1차모드가 일치함을 확인

Const PI As Double = 3.14159265358979

Sub Main
    Dim App As femap.model
    Set App = GetObject(, "femap.model")
    Dim inner As femap.Set, outer As femap.Set, props As femap.Set, swapSet As femap.Set
    Set inner = App.feSet
    Set outer = App.feSet
    Set props = App.feSet
    Dim pr As femap.Prop, nd As femap.Node, el As femap.Elem, cs As femap.CSys
    Set pr = App.feProp
    Set nd = App.feNode
    Set el = App.feElem
    Set cs = App.feCSys
    Dim axisBase As Variant, axis As Variant, first As Variant, radial As Variant, tangent As Variant
    Dim ringCenter As Variant, d As Variant, u As Variant, midXYZ As Variant
    Dim xPoint As Variant, yPoint As Variant, checkA As Variant, checkB As Variant
    Dim innerIDs As Variant, innerXYZ As Variant, outerIDs As Variant, outerXYZ As Variant
    Dim innerPolar As Variant, outerPolar As Variant, swapArray As Variant
    Dim planIn() As Variant, planOut() As Variant, planMid() As Variant, planU() As Variant
    Dim nIn As Long, nOut As Long, n As Long, i As Long, k As Long, pid As Long, rc As Long
    Dim sid As Long, rIn As Double, rOut As Double, diameter As Double, h As Double
    Dim theta As Double, rMid As Double, numericCount As Double, swapRadius As Double
    Dim swapCount As Long, aIDs As String, bIDs As String, innerLabel As String, outerLabel As String
    Dim s As String, stage As String, why As String, msg As String, orderText As String
    Dim occupied As Object, planned As Object, testNodes As Variant
    Dim nCreated As Long, eCreated As Long, cCreated As Long, done As Long
    Dim nodeA As Long, nodeB As Long, cid As Long, eid As Long
    Dim creating As Boolean
    On Error GoTo RuntimeError

    If inner.Select(FT_SURFACE, True, "Side A: select cylindrical attachment surface(s) on first part") <> FE_OK Then Exit Sub
    If inner.Count = 0 Then Exit Sub
    If outer.Select(FT_SURFACE, True, "Side B: select opposing cylindrical attachment surface(s) on second part") <> FE_OK Then Exit Sub
    If outer.Count = 0 Then Exit Sub
    aIDs = SurfaceIDs(inner)
    bIDs = SurfaceIDs(outer)
    rc = App.feAppMessage(FCM_NORMAL, "--- CYLINDRICAL GLUE V2: selected surfaces (no changes) ---")
    rc = App.feAppMessage(FCM_NORMAL, "Side A Surface IDs: " & aIDs)
    rc = App.feAppMessage(FCM_NORMAL, "Side B Surface IDs: " & bIDs)
    rc = inner.Reset
    sid = inner.Next
    stage = "Extracting cylinder axis from Surface " & CStr(sid)
    If App.feVectorAxisOfSurface(sid, axisBase, axis) <> FE_OK Then GoTo Rejected
    If LengthOf(axis) < 0.000000001 Then GoTo Rejected
    axis = ScaleV(axis, 1# / LengthOf(axis))
    ' Choose a stable sign for the optical axis independent of surface orientation.
    k = 0
    For i = 1 To 2
        If Abs(CDbl(axis(i))) > Abs(CDbl(axis(k))) Then k = i
    Next i
    If CDbl(axis(k)) < 0# Then axis = ScaleV(axis, -1#)
    rc = App.feAppMessage(FCM_NORMAL, "Reference axis base=" & VecText(axisBase) & "; direction=" & VecText(axis))

    stage = "Picking first glue center"
    If App.feCoordPick("First glue center: desired height and angle (pick point/node or enter XYZ)", first) <> FE_OK Then Exit Sub
    d = SubV(first, axisBase)
    h = DotV(d, axis)
    ringCenter = AddV(axisBase, ScaleV(axis, h))
    radial = SubV(first, ringCenter)
    If LengthOf(radial) < 0.000000001 Then
        why = "First glue point must be off the cylinder axis."
        GoTo Rejected
    End If
    radial = ScaleV(radial, 1# / LengthOf(radial))
    tangent = CrossV(axis, radial)

    s = InputBox("Number of equally spaced glue spots (360 degrees)." & vbCrLf & _
        "Use 1 to test a single connection.", "Glue count", "24")
    If Len(Trim(s)) = 0 Then Exit Sub
    If Not IsNumeric(s) Then GoTo BadInput
    numericCount = CDbl(s)
    If numericCount < 1# Or numericCount > 720# Then GoTo BadInput
    If numericCount <> Fix(numericCount) Then GoTo BadInput
    n = CLng(numericCount)
    s = InputBox("Glue DIAMETER (not radius), in model length units." & vbCrLf & _
        "Nodes are selected only on each chosen cylindrical surface.", "Glue diameter")
    If Len(Trim(s)) = 0 Then Exit Sub
    If Not IsNumeric(s) Then GoTo BadInput
    diameter = CDbl(s)
    If diameter <= 0# Then GoTo BadInput

    If props.Select(FT_PROP, False, "Select ONE existing 6-DOF CBUSH property (K1 through K6)") <> FE_OK Then Exit Sub
    If props.Count <> 1 Then Exit Sub
    rc = props.Reset
    pid = props.Next
    stage = "Reading selected property"
    If pr.Get(pid) <> FE_OK Then GoTo Rejected
    If pr.Type <> 6 Then
        why = "Property is not Spring/Damper."
        GoTo Rejected
    End If
    If MsgBox("Property " & CStr(pid) & ": " & pr.title & vbCrLf & _
        "Confirm its Type is CBUSH (NOT Other or CBUSH1D), with your K1-K6." & vbCrLf & _
        "The property will not be modified.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Check CBUSH property") <> vbYes Then Exit Sub

    stage = "Reading Side A solid faces and cylinder radius"
    rc = App.feAppMessage(FCM_NORMAL, "--- CYLINDRICAL GLUE PREVIEW (no model changes) ---")
    If Not CollectSide(App, inner, axisBase, axis, innerIDs, innerXYZ, nIn, rIn, why) Then GoTo Rejected
    rc = App.feAppMessage(FCM_NORMAL, "SIDE A radius=" & CStr(rIn) & "; Surface IDs=" & aIDs)
    stage = "Reading Side B solid faces and cylinder radius"
    If Not CollectSide(App, outer, axisBase, axis, outerIDs, outerXYZ, nOut, rOut, why) Then GoTo Rejected
    rc = App.feAppMessage(FCM_NORMAL, "SIDE B radius=" & CStr(rOut) & "; Surface IDs=" & bIDs)
    rc = App.feAppMessage(FCM_NORMAL, "Signed radius difference B-A=" & CStr(rOut - rIn))
    If rIn <= 0# Or rOut <= 0# Then GoTo Rejected
    stage = "Checking selected cylinder order"
    innerLabel = aIDs
    outerLabel = bIDs
    orderText = "A=inner, B=outer"
    If rOut < rIn Then
        rc = App.feAppMessage(FCM_NORMAL, "Selection order reversed: B has the smaller radius. Proposed inner=B; outer=A.")
        msg = "Side A radius: " & CStr(rIn) & vbCrLf & _
            "Side B radius: " & CStr(rOut) & vbCrLf & _
            "Absolute gap: " & CStr(rIn - rOut) & vbCrLf & _
            "Side B will become inner; Side A will become outer." & vbCrLf & _
            "Confirm these are the intended opposing glue faces (not another step or interference)." & vbCrLf & _
            "No = stop without model changes."
        If MsgBox(msg, vbYesNo Or vbQuestion Or vbDefaultButton2, "Reverse radius order?") <> vbYes Then Exit Sub
        swapRadius = rIn
        rIn = rOut
        rOut = swapRadius
        swapCount = nIn
        nIn = nOut
        nOut = swapCount
        swapArray = innerIDs
        innerIDs = outerIDs
        outerIDs = swapArray
        swapArray = innerXYZ
        innerXYZ = outerXYZ
        outerXYZ = swapArray
        Set swapSet = inner
        Set inner = outer
        Set outer = swapSet
        innerLabel = bIDs
        outerLabel = aIDs
        orderText = "B=inner, A=outer (reversed)"
    End If
    rc = App.feAppMessage(FCM_NORMAL, "FINAL order: " & orderText)
    rc = App.feAppMessage(FCM_NORMAL, "FINAL inner Surface IDs=" & innerLabel & "; outer Surface IDs=" & outerLabel)
    If diameter >= PI * rIn Then
        why = "Glue diameter too large for this cylinder."
        GoTo Rejected
    End If
    If n > 1 Then
        If diameter >= 2# * PI * rIn / n Then
            why = "Glue diameter too large for the circumferential spacing; footprints would touch/overlap."
            GoTo Rejected
        End If
    End If
    rMid = (rIn + rOut) / 2#
    rc = App.feAppMessage(FCM_NORMAL, "Inner radius=" & CStr(rIn) & "; outer radius=" & CStr(rOut) & "; gap=" & CStr(rOut - rIn))
    rc = App.feAppMessage(FCM_NORMAL, "CBUSH midpoint radius=" & CStr(rMid))
    rc = App.feAppMessage(FCM_NORMAL, "Axis=" & VecText(axis) & "; row center=" & VecText(ringCenter))
    rc = App.feAppMessage(FCM_NORMAL, "Surface nodes: inner=" & CStr(nIn) & ", outer=" & CStr(nOut))
    rc = App.feAppMessage(FCM_NORMAL, "Diameter=" & CStr(diameter) & "; pitch=" & CStr(360# / n) & " degrees; Property=" & CStr(pid))

    stage = "Caching cylindrical coordinates for fast footprint searches"
    innerPolar = PolarCache(innerXYZ, nIn, ringCenter, axis, radial, tangent)
    outerPolar = PolarCache(outerXYZ, nOut, ringCenter, axis, radial, tangent)

    stage = "Indexing existing rigid connections"
    Set occupied = CreateObject("Scripting.Dictionary")
    Set planned = CreateObject("Scripting.Dictionary")
    If Not ExistingRigidNodes(App, occupied, why) Then GoTo Rejected
    ReDim planIn(0 To n - 1)
    ReDim planOut(0 To n - 1)
    ReDim planMid(0 To n - 1)
    ReDim planU(0 To n - 1)
    For i = 0 To n - 1
        stage = "Planning glue " & CStr(i + 1) & ": inner footprint"
        theta = 2# * PI * i / n
        u = AddV(ScaleV(radial, Cos(theta)), ScaleV(tangent, Sin(theta)))
        midXYZ = AddV(ringCenter, ScaleV(u, rMid))
        planMid(i) = midXYZ
        planU(i) = u
        If Not FindPatch(innerIDs, innerPolar, nIn, theta, rIn, diameter / 2#, occupied, planned, testNodes, why) Then GoTo Rejected
        planIn(i) = testNodes
        stage = "Planning glue " & CStr(i + 1) & ": outer footprint"
        If Not FindPatch(outerIDs, outerPolar, nOut, theta, rOut, diameter / 2#, occupied, planned, testNodes, why) Then GoTo Rejected
        planOut(i) = testNodes
        rc = App.feAppMessage(FCM_NORMAL, "Glue " & CStr(i + 1) & ": center=" & VecText(midXYZ) & "; inner=" & CStr(UBound(planIn(i)) + 1) & "; outer=" & CStr(UBound(planOut(i)) + 1) & "; local X=" & VecText(ScaleV(u, -1#)))
    Next i
    rc = App.feAppMessage(FCM_NORMAL, "PREVIEW OK: " & CStr(2 * n) & " nodes, " & CStr(2 * n) & " RBE2, " & CStr(n) & " CBUSH, " & CStr(n) & " coordinate systems.")
    msg = "Create " & CStr(n) & " zero-length CBUSH connections?" & vbCrLf & _
        "Surface order: " & orderText & vbCrLf & _
        "Inner diameter: " & CStr(2# * rIn) & vbCrLf & _
        "Outer diameter: " & CStr(2# * rOut) & vbCrLf & _
        "Gap: " & CStr(rOut - rIn) & vbCrLf & _
        "CBUSH radius: " & CStr(rMid) & vbCrLf & _
        "Glue diameter: " & CStr(diameter) & vbCrLf & _
        "K1: inward, K2: tangent, K3: cylinder axis." & vbCrLf & _
        "RBE2 dependent DOFs: 123 (solid mesh)." & vbCrLf & _
        "Check Messages for all positions and node counts first." & vbCrLf & _
        "Test on a saved COPY. No merges allowed."
    If MsgBox(msg, vbYesNo Or vbQuestion Or vbDefaultButton2, "Preview - create connections?") <> vbYes Then Exit Sub

    creating = True
    For i = 0 To n - 1
        stage = "Glue " & CStr(i + 1) & ": coordinate system"
        midXYZ = planMid(i)
        u = planU(i)
        xPoint = AddV(midXYZ, ScaleV(u, -1#))
        yPoint = AddV(midXYZ, ScaleV(CrossV(axis, u), -1#))
        Set cs = App.feCSys
        cs.Type = 0
        cs.defCSys = 0
        cs.title = "Glue_ZeroCBUSH_" & CStr(i + 1)
        If cs.XYPoints(midXYZ, xPoint, yPoint) <> FE_OK Then GoTo WriteFailed
        cid = cs.NextEmptyID
        If cid <= 2 Then GoTo WriteFailed
        If cs.Put(cid) <> FE_OK Then GoTo WriteFailed
        cCreated = cCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created CSys " & CStr(cid))
        ' Verify that FEMAP really stored the requested axes.
        checkA = MakeV(1#, 0#, 0#)
        If App.feCoordTransform(cid, checkA, 0, checkB) <> FE_OK Then GoTo WriteFailed
        If LengthOf(SubV(checkB, xPoint)) > 0.000001 Then GoTo WriteFailed
        checkA = MakeV(0#, 1#, 0#)
        If App.feCoordTransform(cid, checkA, 0, checkB) <> FE_OK Then GoTo WriteFailed
        If LengthOf(SubV(checkB, yPoint)) > 0.000001 Then GoTo WriteFailed

        stage = "Glue " & CStr(i + 1) & ": inner center node"
        nodeA = CreateCenter(App, midXYZ)
        If nodeA <= 0 Then GoTo WriteFailed
        nCreated = nCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created inner center node " & CStr(nodeA))
        stage = "Glue " & CStr(i + 1) & ": outer center node"
        nodeB = CreateCenter(App, midXYZ)
        If nodeB <= 0 Or nodeB = nodeA Then GoTo WriteFailed
        nCreated = nCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created outer center node " & CStr(nodeB))
        If App.feCoordOnNode(nodeA, checkA) <> FE_OK Then GoTo WriteFailed
        If App.feCoordOnNode(nodeB, checkB) <> FE_OK Then GoTo WriteFailed
        If LengthOf(SubV(checkA, checkB)) > 0.000000001 Then GoTo WriteFailed
        If LengthOf(SubV(checkA, midXYZ)) > 0.000001 Then GoTo WriteFailed
        stage = "Glue " & CStr(i + 1) & ": inner RBE2"
        testNodes = planIn(i)
        eid = CreateRBE2(App, nodeA, testNodes)
        If eid <= 0 Then GoTo WriteFailed
        eCreated = eCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created inner RBE2 " & CStr(eid))
        If Not CheckRBE2(App, eid, nodeA, testNodes) Then GoTo WriteFailed
        stage = "Glue " & CStr(i + 1) & ": outer RBE2"
        testNodes = planOut(i)
        eid = CreateRBE2(App, nodeB, testNodes)
        If eid <= 0 Then GoTo WriteFailed
        eCreated = eCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created outer RBE2 " & CStr(eid))
        If Not CheckRBE2(App, eid, nodeB, testNodes) Then GoTo WriteFailed

        stage = "Glue " & CStr(i + 1) & ": zero-length CBUSH"
        Set el = App.feElem
        el.Type = 6
        el.topology = 0
        el.propID = pid
        ' Order outer -> inner for consistent local inward axial sign.
        el.node(0) = nodeB
        el.node(1) = nodeA
        el.SpringUsePropCID = False
        el.SpringNoOrient = False
        el.SpringUseCID = True
        el.SpringCID = cid
        el.SpringUsePropLocation = False
        el.SpringUseLocation = True
        el.SpringLocation = 0.5
        eid = el.NextEmptyID
        If eid <= 0 Then GoTo WriteFailed
        If el.Put(eid) <> FE_OK Then GoTo WriteFailed
        eCreated = eCreated + 1
        rc = App.feAppMessage(FCM_NORMAL, "Created CBUSH " & CStr(eid) & " with CID " & CStr(cid))
        If el.Get(eid) <> FE_OK Then GoTo WriteFailed
        If el.propID <> pid Or el.node(0) <> nodeB Or el.node(1) <> nodeA Then GoTo WriteFailed
        If Not CBool(el.SpringUseCID) Or CBool(el.SpringUsePropCID) Or CBool(el.SpringNoOrient) Or el.SpringCID <> cid Then GoTo WriteFailed
        done = done + 1
    Next i
    rc = App.feViewRegenerate(0)
    rc = App.feAppMessage(FCM_NORMAL, "DONE: " & CStr(done) & " complete zero-length glue connections. Do NOT merge coincident nodes.")
    MsgBox CStr(done) & " glue connections created." & vbCrLf & _
        "Check RBE2 DOF 123, CBUSH Orientation CSys, and exported CBUSH/PBUSH cards." & vbCrLf & _
        "Never merge the coincident center nodes.", vbInformation, "Complete"
    Exit Sub

BadInput:
    MsgBox "Invalid number or glue diameter. No model changes.", vbExclamation
    Exit Sub
Rejected:
    If Len(why) = 0 Then why = "FEMAP API returned a failure."
    rc = App.feAppMessage(FCM_NORMAL, "STOP: " & stage & " - " & why)
    MsgBox stage & vbCrLf & why & vbCrLf & "No model changes.", vbExclamation, "Preview stopped"
    Exit Sub
WriteFailed:
    why = "FEMAP API failed or readback did not match."
    GoTo ReportError
RuntimeError:
    why = "Error " & CStr(Err.Number) & ": " & Err.Description
ReportError:
    On Error Resume Next
    rc = App.feAppMessage(FCM_NORMAL, "STOP: " & stage & " - " & why)
    If creating Then
        MsgBox stage & vbCrLf & why & vbCrLf & _
            "Partial new entities may remain. Fully completed glues: " & CStr(done) & vbCrLf & _
            "Use the saved model copy before retrying; see Messages for created IDs.", vbExclamation, "Creation stopped"
    Else
        MsgBox stage & vbCrLf & why & vbCrLf & "No model changes.", vbExclamation
    End If
End Sub

' Preserve original side labels for diagnostics and final confirmation.
Function SurfaceIDs(ss As femap.Set) As String
    Dim rc As Long, sid As Long, s As String
    rc = ss.Reset
    sid = ss.Next
    Do While sid > 0
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(sid)
        sid = ss.Next
    Loop
    SurfaceIDs = s
End Function

' Collect solid face nodes and validate revolution axis/radius of every surface.
Function CollectSide(App As femap.model, faces As femap.Set, axisBase As Variant, axis As Variant, ids As Variant, xyz As Variant, count As Long, radius As Double, why As String) As Boolean
    Dim sf As femap.Surface, el As femap.Elem, nd As femap.Node, nodes As femap.Set
    Set sf = App.feSurface
    Set el = App.feElem
    Set nd = App.feNode
    Set nodes = App.feSet
    Dim sid As Long, rc As Long, nf As Long, j As Long, k As Long, nid As Long
    Dim ef As Variant, fi As Variant, fn As Variant, good As Variant
    Dim b As Variant, a As Variant, diff As Variant, point As Variant, projected As Variant, v As Variant
    Dim r As Double, tol As Double, stepF As Long, f As Long, nsample As Long
    Dim rawR As Double, rawMin As Double, rawMax As Double, cadMin As Double, cadMax As Double
    Dim rawCount As Long, surfaceCount As Long, axisOffset As Double
    rc = faces.Reset
    sid = faces.Next
    Do While sid > 0
        why = "Surface " & CStr(sid)
        If App.feVectorAxisOfSurface(sid, b, a) <> FE_OK Then
            why = why & ": not an extractable revolved surface."
            Exit Function
        End If
        If LengthOf(a) < 0.000000001 Then Exit Function
        a = ScaleV(a, 1# / LengthOf(a))
        diff = SubV(b, axisBase)
        diff = SubV(diff, ScaleV(axis, DotV(diff, axis)))
        axisOffset = LengthOf(diff)
        rc = App.feAppMessage(FCM_NORMAL, "Surface " & CStr(sid) & " axis base=" & VecText(b) & "; direction=" & VecText(a) & "; axis offset=" & CStr(axisOffset))
        If Abs(DotV(a, axis)) < 0.99999999 Or axisOffset > 0.00001 * (1# + radius) Then
            why = why & ": cylinder axes are not concentric."
            Exit Function
        End If
        If sf.Get(sid) <> FE_OK Then Exit Function
        nf = 0
        rc = sf.ElementFaces(False, True, False, nf, ef, fi)
        If rc <> FE_OK Or nf < 1 Then
            why = why & ": no associated solid mesh faces."
            Exit Function
        End If
        rc = App.feAppMessage(FCM_NORMAL, why & ": " & CStr(nf) & " solid faces")
        rawMin = 1.0E+30
        rawMax = 0#
        cadMin = 1.0E+30
        cadMax = 0#
        rawCount = 0
        surfaceCount = 0
        stepF = nf \ 12
        If stepF < 1 Then stepF = 1
        For f = 0 To nf - 1
            If el.Get(CLng(ef(f))) <> FE_OK Then Exit Function
            If el.GetFaceNodes(CLng(fi(f)), fn) <> FE_OK Then Exit Function
            For j = 0 To 7
                nid = CLng(fn(j))
                If nid > 0 Then rc = nodes.Add(nid)
            Next j
            ' Obtain actual CAD radius from projected mesh coordinates, not facet centroids.
            If f Mod stepF = 0 Or f = nf - 1 Then
                For k = 0 To 3
                    nid = CLng(fn(k))
                    If nid > 0 Then
                        If App.feCoordOnNode(nid, point) <> FE_OK Then Exit Function
                        If k < 3 Then
                            v = SubV(point, axisBase)
                            v = SubV(v, ScaleV(axis, DotV(v, axis)))
                            rawR = LengthOf(v)
                            If rawR < rawMin Then rawMin = rawR
                            If rawR > rawMax Then rawMax = rawR
                            rawCount = rawCount + 1
                        End If
                        If App.feCoordOntoSurface(sid, point, projected) <> FE_OK Then Exit Function
                        v = SubV(projected, axisBase)
                        v = SubV(v, ScaleV(axis, DotV(v, axis)))
                        r = LengthOf(v)
                        If r <= 0.000000001 Then Exit Function
                        If r < cadMin Then cadMin = r
                        If r > cadMax Then cadMax = r
                        surfaceCount = surfaceCount + 1
                        If nsample = 0 Then radius = r
                        tol = 0.00001 * (1# + radius)
                        If Abs(r - radius) > tol Then
                            rc = App.feAppMessage(FCM_NORMAL, "RADIUS MISMATCH: Surface " & CStr(sid) & "; expected=" & CStr(radius) & "; projected=" & CStr(r) & "; sample node=" & CStr(nid))
                            rc = App.feAppMessage(FCM_NORMAL, "Current Surface CAD samples min/max=" & CStr(cadMin) & " / " & CStr(cadMax))
                            If rawCount > 0 Then rc = App.feAppMessage(FCM_NORMAL, "Current Surface mesh corner radii min/max=" & CStr(rawMin) & " / " & CStr(rawMax))
                            why = why & ": sampled CAD radii vary. Select cylindrical faces of one radius only."
                            Exit Function
                        End If
                        nsample = nsample + 1
                    End If
                Next k
            End If
        Next f
        If surfaceCount > 0 Then rc = App.feAppMessage(FCM_NORMAL, "Surface " & CStr(sid) & " CAD projected radius min/max=" & CStr(cadMin) & " / " & CStr(cadMax))
        If rawCount > 0 Then rc = App.feAppMessage(FCM_NORMAL, "Surface " & CStr(sid) & " mesh corner radius min/max=" & CStr(rawMin) & " / " & CStr(rawMax))
        If surfaceCount > 0 And rawCount > 0 Then
            If Abs(cadMax - rawMax) > 0.00001 * (1# + radius) Then
                rc = App.feAppMessage(FCM_NORMAL, "WARNING: CAD projection and mesh corner radii differ. Verify selected Surface and mesh association before creating.")
            End If
        End If
        sid = faces.Next
    Loop
    count = nodes.Count
    If count < 3 Or nsample = 0 Then Exit Function
    rc = nd.GetCoordArray2(nodes.ID, count, ids, good, xyz)
    If rc <> FE_OK Then Exit Function
    For j = 0 To count - 1
        If Not CBool(good(j)) Then Exit Function
    Next j
    why = ""
    CollectSide = True
End Function

' Conservative: reject any patch using an already connected rigid node.
Function ExistingRigidNodes(App As femap.model, occupied As Object, why As String) As Boolean
    Dim rig As femap.Set, ns As femap.Set
    Set rig = App.feSet
    Set ns = App.feSet
    Dim rc As Long, eid As Long, nid As Long, key As String
    rc = rig.AddRule(29, FGD_ELEM_BYTYPE)
    If rc <> FE_OK And rig.Count > 0 Then
        why = "Unable to read existing rigid elements."
        Exit Function
    End If
    rc = rig.Reset
    eid = rig.Next
    Do While eid > 0
        rc = ns.Clear
        rc = ns.AddRule(eid, FGD_NODE_ONELEM)
        If rc <> FE_OK Or ns.Count = 0 Then
            why = "Unable to read existing rigid " & CStr(eid)
            Exit Function
        End If
        rc = ns.Reset
        nid = ns.Next
        Do While nid > 0
            key = CStr(nid)
            If Not occupied.Exists(key) Then occupied.Add key, eid
            nid = ns.Next
        Loop
        eid = rig.Next
    Loop
    ExistingRigidNodes = True
End Function

' Store [angle, relative axial height] once for each surface node.
Function PolarCache(xyz As Variant, n As Long, center As Variant, axis As Variant, u As Variant, t As Variant) As Variant
    Dim v() As Double, j As Long, x As Double, y As Double, z As Double
    Dim cc As Double, ss As Double
    ReDim v(0 To 2 * n - 1)
    For j = 0 To n - 1
        x = CDbl(xyz(3 * j)) - CDbl(center(0))
        y = CDbl(xyz(3 * j + 1)) - CDbl(center(1))
        z = CDbl(xyz(3 * j + 2)) - CDbl(center(2))
        cc = x * CDbl(u(0)) + y * CDbl(u(1)) + z * CDbl(u(2))
        ss = x * CDbl(t(0)) + y * CDbl(t(1)) + z * CDbl(t(2))
        v(2 * j) = Atn2(ss, cc)
        v(2 * j + 1) = x * CDbl(axis(0)) + y * CDbl(axis(1)) + z * CDbl(axis(2))
    Next j
    PolarCache = v
End Function

' Circular footprint on developed cylindrical surface: (R*deltaAngle)^2 + dz^2.
' Require at least three nodes, nondegenerate area, and points around the center.
' No nearest-node fallback that silently changes the requested glue footprint.
Function FindPatch(ids As Variant, polar As Variant, n As Long, theta As Double, r As Double, rad As Double, occupied As Object, planned As Object, result As Variant, why As String) As Boolean
    Dim j As Long, m As Long, k As Long
    Dim zz As Double, dd As Double, angle As Double, lim As Double
    Dim aa As Double, bb As Double, area2 As Double, tol As Double
    Dim dx() As Double, dy() As Double, chosen() As Long, angles() As Double
    Dim key As String, p As Long, minAngle As Double, gapAngle As Double, maxGap As Double
    lim = rad * rad
    ReDim chosen(0 To n - 1)
    ReDim dx(0 To n - 1)
    ReDim dy(0 To n - 1)
    ReDim angles(0 To n - 1)
    For j = 0 To n - 1
        zz = CDbl(polar(2 * j + 1))
        If Abs(zz) <= rad Then
            angle = CDbl(polar(2 * j)) - theta
            Do While angle < -PI
                angle = angle + 2# * PI
            Loop
            Do While angle > PI
                angle = angle - 2# * PI
            Loop
            dd = r * angle
            If dd * dd + zz * zz <= lim * (1# + 0.00000001) Then
                key = CStr(ids(j))
                If occupied.Exists(key) Then
                    why = "Footprint overlaps an existing rigid at node " & key & "."
                    Exit Function
                End If
                If planned.Exists(key) Then
                    why = "Footprint shares node " & key & " with another planned RBE2."
                    Exit Function
                End If
                chosen(m) = CLng(ids(j))
                dx(m) = dd
                dy(m) = zz
                m = m + 1
            End If
        End If
    Next j
    If m < 3 Then
        why = "Only " & CStr(m) & " surface nodes in glue diameter. Refine mesh or check diameter/height."
        Exit Function
    End If
    tol = 0.00000001 * lim
    ' Non-collinear dependent nodes are needed to transfer all 6 center DOFs.
    k = 1
    dd = 0#
    For j = 1 To m - 1
        aa = (dx(j) - dx(0)) ^ 2 + (dy(j) - dy(0)) ^ 2
        If aa > dd Then
            k = j
            dd = aa
        End If
    Next j
    area2 = 0#
    For j = 1 To m - 1
        aa = Abs((dx(k) - dx(0)) * (dy(j) - dy(0)) - (dy(k) - dy(0)) * (dx(j) - dx(0)))
        If aa > area2 Then area2 = aa
    Next j
    If area2 <= tol Then
        why = "Footprint nodes are collinear or degenerate. Refine mesh."
        Exit Function
    End If
    ' Center must lie within the patch's developed-surface convex hull.
    p = 0
    For j = 0 To m - 1
        If dx(j) * dx(j) + dy(j) * dy(j) > tol Then
            angles(p) = Atn2(dy(j), dx(j))
            p = p + 1
        End If
    Next j
    If p < 3 Then
        why = "Not enough nodes around the glue center. Refine mesh."
        Exit Function
    End If
    ' Insertion sort is cheap for small glue footprints.
    For j = 1 To p - 1
        angle = angles(j)
        k = j - 1
        Do While k >= 0
            If angles(k) <= angle Then Exit Do
            angles(k + 1) = angles(k)
            k = k - 1
        Loop
        angles(k + 1) = angle
    Next j
    maxGap = angles(0) + 2# * PI - angles(p - 1)
    For j = 1 To p - 1
        gapAngle = angles(j) - angles(j - 1)
        If gapAngle > maxGap Then maxGap = gapAngle
    Next j
    If maxGap >= PI - 0.00000001 Then
        why = "Nodes do not surround the glue center. Check axial row/selected surfaces or refine mesh."
        Exit Function
    End If
    ReDim Preserve chosen(0 To m - 1)
    For j = 0 To m - 1
        planned.Add CStr(chosen(j)), True
    Next j
    result = chosen
    why = ""
    FindPatch = True
End Function

Function CreateCenter(App As femap.model, xyz As Variant) As Long
    Dim nd As femap.Node, nid As Long
    Set nd = App.feNode
    nd.x = CDbl(xyz(0))
    nd.y = CDbl(xyz(1))
    nd.z = CDbl(xyz(2))
    nid = nd.NextEmptyID
    If nid <= 0 Then Exit Function
    If nd.Put(nid) <> FE_OK Then Exit Function
    CreateCenter = nid
End Function

Function CreateRBE2(App As femap.model, center As Long, nodes As Variant) As Long
    Dim el As femap.Elem, n As Long, i As Long, eid As Long
    Dim faces() As Long, weights() As Double, dofs() As Long
    Dim vf As Variant, vw As Variant, vd As Variant
    Set el = App.feElem
    n = UBound(nodes) + 1
    If n < 3 Then Exit Function
    ReDim faces(0 To n - 1)
    ReDim weights(0 To n - 1)
    ReDim dofs(0 To 6 * n - 1)
    ' Standard RBE2 uses release(0,0..5) for the common DOF mask;
    ' individual interpolation masks/weights are unused (zero).
    el.Type = 29
    el.topology = 13
    el.RigidInterpolate = False
    el.node(0) = center
    For i = 0 To 5
        el.release(0, i) = (i < 3)
    Next i
    vf = faces
    vw = weights
    vd = dofs
    If el.PutNodeList(0, n, nodes, vf, vw, vd) <> FE_OK Then Exit Function
    eid = el.NextEmptyID
    If eid <= 0 Then Exit Function
    If el.Put(eid) <> FE_OK Then Exit Function
    CreateRBE2 = eid
End Function

Function CheckRBE2(App As femap.model, eid As Long, center As Long, expected As Variant) As Boolean
    Dim el As femap.Elem, n As Long, j As Long, found As Object, key As String
    Dim actual As Variant, f As Variant, w As Variant, d As Variant
    Set el = App.feElem
    If el.Get(eid) <> FE_OK Then Exit Function
    If el.Type <> 29 Or el.topology <> 13 Or CBool(el.RigidInterpolate) Then Exit Function
    If el.node(0) <> center Then Exit Function
    For j = 0 To 5
        If CBool(el.release(0, j)) <> (j < 3) Then Exit Function
    Next j
    If el.GetNodeList(0, n, actual, f, w, d) <> FE_OK Then Exit Function
    If n <> UBound(expected) + 1 Then Exit Function
    Set found = CreateObject("Scripting.Dictionary")
    For j = 0 To n - 1
        key = CStr(actual(j))
        If found.Exists(key) Then Exit Function
        found.Add key, True
    Next j
    For j = 0 To n - 1
        If Not found.Exists(CStr(expected(j))) Then Exit Function
    Next j
    CheckRBE2 = True
End Function

Function MakeV(x As Double, y As Double, z As Double) As Variant
    Dim a(0 To 2) As Double
    a(0) = x
    a(1) = y
    a(2) = z
    MakeV = a
End Function

Function AddV(a As Variant, b As Variant) As Variant
    AddV = MakeV(CDbl(a(0)) + CDbl(b(0)), CDbl(a(1)) + CDbl(b(1)), CDbl(a(2)) + CDbl(b(2)))
End Function

Function SubV(a As Variant, b As Variant) As Variant
    SubV = MakeV(CDbl(a(0)) - CDbl(b(0)), CDbl(a(1)) - CDbl(b(1)), CDbl(a(2)) - CDbl(b(2)))
End Function

Function ScaleV(a As Variant, s As Double) As Variant
    ScaleV = MakeV(CDbl(a(0)) * s, CDbl(a(1)) * s, CDbl(a(2)) * s)
End Function

Function DotV(a As Variant, b As Variant) As Double
    DotV = CDbl(a(0)) * CDbl(b(0)) + CDbl(a(1)) * CDbl(b(1)) + CDbl(a(2)) * CDbl(b(2))
End Function

Function CrossV(a As Variant, b As Variant) As Variant
    CrossV = MakeV(CDbl(a(1)) * CDbl(b(2)) - CDbl(a(2)) * CDbl(b(1)), CDbl(a(2)) * CDbl(b(0)) - CDbl(a(0)) * CDbl(b(2)), CDbl(a(0)) * CDbl(b(1)) - CDbl(a(1)) * CDbl(b(0)))
End Function

Function LengthOf(a As Variant) As Double
    LengthOf = Sqr(DotV(a, a))
End Function

Function Atn2(y As Double, x As Double) As Double
    If x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then
            Atn2 = Atn(y / x) + PI
        Else
            Atn2 = Atn(y / x) - PI
        End If
    ElseIf y > 0# Then
        Atn2 = PI / 2#
    ElseIf y < 0# Then
        Atn2 = -PI / 2#
    Else
        Atn2 = 0#
    End If
End Function

Function VecText(a As Variant) As String
    VecText = "(" & Format(CDbl(a(0)), "0.000000") & ", " & Format(CDbl(a(1)), "0.000000") & ", " & Format(CDbl(a(2)), "0.000000") & ")"
End Function
