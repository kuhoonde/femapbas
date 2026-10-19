' FEMAP API / WinWrap Basic
' Move the two DISTINCT end nodes of each selected CBUSH to THEIR midpoint.
' No proximity search, merge, new node, deleted node, or element rewrite.
'
' USE:
'   1. Open a working COPY of the model in FEMAP.
'   2. Confirm every selected CBUSH has an explicit Orientation CSys (CID).
'      A former GA->GB / vector orientation alone is not valid at zero length.
'      This script does NOT change element/property orientations or offsets.
'   3. Open this file in the FEMAP API Programming window and run Main.
'   4. Select actual CBUSH elements ONLY. FEMAP type Spring/Damper also
'      includes other solver element types; solver formulation is not inferred.
'
' PREVIEW_ONLY=True lists the plan without moving nodes.
' PREVIEW_ONLY=False applies the complete, prechecked plan.
' The default is preview. Change this one constant to False to apply.
'
' Prechecks, completed before the first model write:
'   - selected elements are two-node Spring/Damper elements;
'   - distinct node IDs and existing ordinary GRID nodes;
'   - a shared endpoint cannot belong to different endpoint pairs;
'   - endpoints may connect only to selected springs, rigid/interpolation
'     elements, or concentrated masses. Direct solid/shell/beam endpoints
'     are rejected, to avoid changing the physical element mesh;
'   - permanent nodal constraints are rejected.
' All target coordinates are computed from the ORIGINAL coordinates.
' Multiple parallel CBUSH elements using the same pair are allowed.
' A different pair stays separate even if its coordinates are very close.
'
' Successful updates preserve node IDs, definition/output coordinate systems,
' loads, constraints, and all element connectivity. Coordinates are global XYZ.
' A failed write/verification attempts to restore every original coordinate.
' Never save over the source model before reviewing the changed geometry.
'
' Documented interfaces: FEMAP API Reference, Node / Elem / Set objects.
' FEMAP itself is required for integration/GUI verification.

Option Explicit

Const PREVIEW_ONLY As Boolean = False
Const MAX_PLAN_LINES As Long = 40

Sub Main
    Dim App As femap.model
    Dim pick As femap.Set
    Dim adjacent As femap.Set
    Dim el As femap.Elem
    Dim other As femap.Elem
    Dim nd As femap.Node
    Dim elements As Object
    Dim originals As Object
    Dim targets As Object
    Dim owners As Object
    Dim rc As Long
    Dim eid As Long
    Dim ga As Long, gb As Long, nid As Long
    Dim aeid As Long
    Dim i As Long, k As Long
    Dim moved As Long, already As Long, printed As Long
    Dim key As Variant, row As Variant, old As Variant
    Dim pa As Variant, pb As Variant, mid As Variant
    Dim nk As String, pk As String
    Dim message As String, failure As String
    Dim writing As Boolean, restoreOK As Boolean

    On Error GoTo Failed
    Set App = feFemap()
    Set pick = App.feSet
    Set adjacent = App.feSet
    Set el = App.feElem
    Set other = App.feElem
    Set nd = App.feNode
    Set elements = CreateObject("Scripting.Dictionary")
    Set originals = CreateObject("Scripting.Dictionary")
    Set targets = CreateObject("Scripting.Dictionary")
    Set owners = CreateObject("Scripting.Dictionary")

    rc = pick.Select(FT_ELEM, True, "Select CBUSH elements ONLY - pairwise midpoint")
    If rc <> FE_OK Then Exit Sub
    If pick.Count = 0 Then Exit Sub

    ' Snapshot ALL endpoint coordinates before calculating or writing a move.
    eid = pick.First
    Do While eid > 0
        If el.Get(eid) <> FE_OK Then
            Err.Raise 513, "CBUSH midpoint", "Cannot read element " & CStr(eid)
        End If
        If el.type <> FET_L_SPRING Then
            Err.Raise 513, "CBUSH midpoint", "Non-Spring/Damper element selected: " & CStr(eid)
        End If
        ga = el.node(0)
        gb = el.node(1)
        If ga <= 0 Or gb <= 0 Or ga = gb Then
            Err.Raise 513, "CBUSH midpoint", "Element needs two DISTINCT existing node IDs: " & CStr(eid)
        End If
        elements.Add CStr(eid), Array(ga, gb)
        For i = 0 To 1
            If i = 0 Then
                nid = ga
            Else
                nid = gb
            End If
            nk = CStr(nid)
            If Not originals.Exists(nk) Then
                If nd.Get(nid) <> FE_OK Then
                    Err.Raise 513, "CBUSH midpoint", "Cannot read node " & nk
                End If
                If nd.type <> 0 Then
                    Err.Raise 513, "CBUSH midpoint", "Endpoint is not an ordinary node: " & nk
                End If
                For k = 0 To 5
                    If nd.permBC(k) Then
                        Err.Raise 513, "CBUSH midpoint", "Endpoint has a permanent constraint: " & nk
                    End If
                Next k
                originals.Add nk, Array(nd.x, nd.y, nd.z)
            End If
        Next i
        eid = pick.Next
    Loop

    For Each key In elements.Keys
        row = elements.Item(key)
        ga = CLng(row(0))
        gb = CLng(row(1))
        If ga < gb Then
            pk = CStr(ga) & ":" & CStr(gb)
        Else
            pk = CStr(gb) & ":" & CStr(ga)
        End If
        pa = originals.Item(CStr(ga))
        pb = originals.Item(CStr(gb))
        mid = Array(pa(0) + (pb(0) - pa(0)) / 2#, _
                    pa(1) + (pb(1) - pa(1)) / 2#, _
                    pa(2) + (pb(2) - pa(2)) / 2#)
        For i = 0 To 1
            If i = 0 Then
                nk = CStr(ga)
            Else
                nk = CStr(gb)
            End If
            If owners.Exists(nk) Then
                If CStr(owners.Item(nk)) <> pk Then
                    Err.Raise 513, "CBUSH midpoint", "Node " & nk & _
                        " is shared by different CBUSH pairs. No nodes moved."
                End If
            Else
                owners.Add nk, pk
                targets.Add nk, mid
            End If
        Next i
        If printed < MAX_PLAN_LINES Then
            message = "CBUSH " & CStr(key) & ": nodes " & CStr(ga) & ", " & _
                CStr(gb) & " -> (" & CStr(mid(0)) & ", " & CStr(mid(1)) & _
                ", " & CStr(mid(2)) & ")"
            rc = App.feAppMessage(FCM_NORMAL, message)
            printed = printed + 1
        End If
    Next key

    ' Reject endpoints attached directly to real meshes or unselected springs.
    For Each key In targets.Keys
        nid = CLng(key)
        rc = adjacent.Clear
        rc = adjacent.AddRule(nid, FGD_ELEM_BYNODE)
        If adjacent.Count = 0 Then
            Err.Raise 513, "CBUSH midpoint", "Cannot establish connectivity for node " & CStr(key)
        End If
        aeid = adjacent.First
        Do While aeid > 0
            If other.Get(aeid) <> FE_OK Then
                Err.Raise 513, "CBUSH midpoint", "Cannot read adjacent element " & CStr(aeid)
            End If
            If other.type = FET_L_SPRING Then
                If Not elements.Exists(CStr(aeid)) Then
                    Err.Raise 513, "CBUSH midpoint", "Node " & CStr(key) & _
                        " also belongs to unselected spring " & CStr(aeid)
                End If
            ElseIf other.type <> FET_L_RIGID And _
                   other.type <> FET_L_MASS And _
                   other.type <> FET_L_MASS_MATRIX Then
                Err.Raise 513, "CBUSH midpoint", "Node " & CStr(key) & _
                    " is attached to a physical/unsupported element " & CStr(aeid) & _
                    " (FEMAP type " & CStr(other.type) & "). No nodes moved."
            End If
            aeid = adjacent.Next
        Loop
    Next key

    If PREVIEW_ONLY Then
        rc = App.feAppMessage(FCM_NORMAL, "PREVIEW ONLY: " & CStr(elements.Count) & _
            " selected elements, " & CStr(targets.Count) & _
            " endpoint nodes. No changes. Check explicit CBUSH Orientation CSys, " & _
            "then set PREVIEW_ONLY=False to apply.")
        Exit Sub
    End If

    writing = True
    For Each key In targets.Keys
        old = originals.Item(key)
        mid = targets.Item(key)
        If old(0) = mid(0) And old(1) = mid(1) And old(2) = mid(2) Then
            already = already + 1
        Else
            nid = CLng(key)
            If nd.Get(nid) <> FE_OK Then
                Err.Raise 513, "CBUSH midpoint", "Cannot reload node " & CStr(key)
            End If
            nd.x = CDbl(mid(0))
            nd.y = CDbl(mid(1))
            nd.z = CDbl(mid(2))
            If nd.Put(nid) <> FE_OK Then
                Err.Raise 513, "CBUSH midpoint", "Cannot write node " & CStr(key)
            End If
            moved = moved + 1
        End If
    Next key

    For Each key In targets.Keys
        mid = targets.Item(key)
        If nd.Get(CLng(key)) <> FE_OK Then
            Err.Raise 513, "CBUSH midpoint", "Cannot verify node " & CStr(key)
        End If
        If Not CloseCoordinate(nd.x, CDbl(mid(0))) Or _
           Not CloseCoordinate(nd.y, CDbl(mid(1))) Or _
           Not CloseCoordinate(nd.z, CDbl(mid(2))) Then
            Err.Raise 513, "CBUSH midpoint", "Coordinate verification failed at node " & CStr(key)
        End If
    Next key
    writing = False
    rc = App.feViewRegenerate(0)
    rc = App.feAppMessage(FCM_NORMAL, "DONE: " & CStr(elements.Count) & _
        " selected elements; " & CStr(moved) & " nodes moved; " & _
        CStr(already) & " nodes already at target. Node IDs/connectivity preserved.")
    Exit Sub

Failed:
    failure = Err.Description
    On Error Resume Next
    restoreOK = True
    If writing Then
        For Each key In originals.Keys
            old = originals.Item(key)
            Err.Clear
            rc = 0
            rc = nd.Get(CLng(key))
            If Err.Number = 0 And rc = FE_OK Then
                Err.Clear
                nd.x = CDbl(old(0))
                nd.y = CDbl(old(1))
                nd.z = CDbl(old(2))
                If Err.Number <> 0 Then restoreOK = False
                Err.Clear
                rc = 0
                rc = nd.Put(CLng(key))
                If Err.Number <> 0 Or rc <> FE_OK Then restoreOK = False
            Else
                restoreOK = False
            End If
        Next key
        For Each key In originals.Keys
            old = originals.Item(key)
            Err.Clear
            rc = 0
            rc = nd.Get(CLng(key))
            If Err.Number = 0 And rc = FE_OK Then
                If Not CloseCoordinate(nd.x, CDbl(old(0))) Or _
                   Not CloseCoordinate(nd.y, CDbl(old(1))) Or _
                   Not CloseCoordinate(nd.z, CDbl(old(2))) Then restoreOK = False
                If Err.Number <> 0 Then restoreOK = False
            Else
                restoreOK = False
            End If
        Next key
        rc = App.feViewRegenerate(0)
        If restoreOK Then
            failure = failure & " Original endpoint coordinates restored."
        Else
            failure = failure & " Restore incomplete: reload your saved working copy."
        End If
    Else
        failure = failure & " No coordinate writes were started."
    End If
    rc = App.feAppMessage(FCM_ERROR, failure)
End Sub

Function CloseCoordinate(ByVal a As Double, ByVal b As Double) As Boolean
    ' Verification tolerance only; never used for node pairing or merging.
    CloseCoordinate = (Abs(a - b) <= 0.0000000001 * (1# + Abs(a) + Abs(b)))
End Function
