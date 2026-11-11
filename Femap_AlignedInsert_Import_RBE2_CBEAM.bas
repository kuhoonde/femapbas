' ========================================================================
' FEMAP API / WinWrap Basic
' Aligned solid inserts: coincident GRID + RBE2 at each end of a CBEAM.
' Data: insert_aligned_to_base1.bdf + base1_insert_beam_node_pairs.csv
' Use the Base FEMAP model that supplied the node IDs in the pair CSV.
' The beam property must already exist in this model.
' ========================================================================
Option Explicit

Const FE_SUCCESS As Long = -1
Const MAX_PAIR_COUNT As Long = 50000
Const MAX_AXIAL_ERR_MM As Double = 0.002
Const MAX_TANG_ERR_MM As Double = 0.02
Const MIN_BEAM_LEN_MM As Double = 0.05
Const MAX_BEAM_LEN_MM As Double = 10.0
Const PI_NUM As Double = 3.14159265358979
Const NODE_ENTITY As Long = 7
Const ELEMENT_ENTITY As Long = 8
Const MATERIAL_ENTITY As Long = 10
Const PROPERTY_ENTITY As Long = 11
Const GROUP_ENTITY As Long = 24
Const ELEMENT_BY_TYPE As Long = 24
Const LINEAR_BEAM As Long = 5
Const PARABOLIC_BEAM As Long = 37
Const RIGID_ELEMENT As Long = 29
Const RIGID_LIST As Long = 13
Const LINE2 As Long = 0
Const MAX_ENTITY_ID As Long = 99999999

Sub Main()
    On Error GoTo Fail

    Dim App As femap.model
    Set App = feFemap()

    Dim bdfPath As String, pairPath As String
    Dim scratchName As String, detail As String
    Dim reply As Integer, propText As String, beamPropID As Long
    Dim present As Long
    Dim beamProp As femap.Prop
    Dim baseNd As femap.Node, insertNd As femap.Node, auxiliary As femap.Node
    Dim beam As femap.Elem, rigid As femap.Elem
    Dim partners As Object, insertIDs As Object
    Dim nBase() As Long, nInsert() As Long, angleDeg() As Double, xMm() As Double
    Dim nPairs As Long, i As Long
    Dim lineText As String, fields As Variant
    Dim fNum As Integer, rowNum As Long, s As String
    Dim bid As Long, iid As Long, angle As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim tangErr As Double, distanceMM As Double
    Dim worstAx As Double, worstTang As Double
    Dim shortest As Double, longest As Double
    Dim firstNode As Long, firstElem As Long, groupID As Long
    Dim auxBase As Long, auxInsert As Long, rigidID As Long
    Dim addedNodes As Long, addedElems As Long, addedGroup As Boolean, committed As Boolean
    Dim nodeCountBefore As Long, elemCountBefore As Long, groupCountBefore As Long

    scratchName = App.FullScratchName
    If scratchName = "" Or App.Info_Count(NODE_ENTITY) = 0 Then
        Err.Raise 2200, , "Activate the Base FEMAP model used to create the pair CSV."
    End If

    Set beamProp = App.feProp
    propText = InputBox("Enter the existing Glue BEAM property ID.", "CBEAM property ID", "0")
    If Trim$(propText) = "" Then Exit Sub
    If Not IsNumeric(propText) Then
        MsgBox "Not a numeric property ID.", vbExclamation
        Exit Sub
    End If
    beamPropID = ReadID(propText, "beam property ID")
    If beamPropID < 1 Then Exit Sub
    If beamProp.Get(beamPropID) <> FE_SUCCESS Then
        MsgBox "Property " & beamPropID & " does not exist. Create a BEAM property first.", vbCritical
        Exit Sub
    End If
    If beamProp.Type <> LINEAR_BEAM And beamProp.Type <> PARABOLIC_BEAM Then
        MsgBox "Select a CBEAM section property (FEMAP type 5 or 37).", vbCritical
        Exit Sub
    End If

    pairPath = GetFilePath("base1_insert_beam_node_pairs.csv", "csv", , "Select base1_insert_beam_node_pairs.csv", 0)
    If pairPath = "" Then Exit Sub

    ReDim nBase(0 To MAX_PAIR_COUNT - 1)
    ReDim nInsert(0 To MAX_PAIR_COUNT - 1)
    ReDim angleDeg(0 To MAX_PAIR_COUNT - 1)
    ReDim xMm(0 To MAX_PAIR_COUNT - 1)
    Set partners = CreateObject("Scripting.Dictionary")
    Set insertIDs = CreateObject("Scripting.Dictionary")
    fNum = FreeFile
    Open pairPath For Input As #fNum
    rowNum = 0
    nPairs = 0
    Do While Not EOF(fNum)
        Line Input #fNum, lineText
        rowNum = rowNum + 1
        s = Trim$(lineText)
        If Len(s) = 0 Then GoTo NextCSVRow
        If Left$(LCase$(s), 9) = "insert_no" Then GoTo NextCSVRow
        fields = Split(s, ",")
        If UBound(fields) < 5 Then Err.Raise 2201, , "CSV has too few columns at row " & rowNum
        If nPairs >= MAX_PAIR_COUNT Then Err.Raise 2202, , "Increase MAX_PAIR_COUNT in script."
        bid = ReadID(fields(1), "Base GRID at CSV row " & rowNum)
        iid = ReadID(fields(2), "Insert GRID at CSV row " & rowNum)
        angle = Val(Trim$(fields(5)))
        If bid <= 0 Or iid <= 0 Or bid = iid Then Err.Raise 2203, , "Invalid GRID pair at row " & rowNum
        If partners.Exists(bid) Or partners.Exists(iid) Then Err.Raise 2204, , "Repeated GRID at CSV row " & rowNum
        partners.Add bid, iid
        partners.Add iid, bid
        insertIDs.Add iid, True
        nBase(nPairs) = bid
        nInsert(nPairs) = iid
        angleDeg(nPairs) = angle
        xMm(nPairs) = Val(Trim$(fields(3)))
        nPairs = nPairs + 1
NextCSVRow:
    Loop
    Close #fNum
    fNum = 0
    If nPairs = 0 Then Err.Raise 2206, , "No node pairs found in CSV."

    Set baseNd = App.feNode
    Set insertNd = App.feNode
    CheckActive App, scratchName
    For i = 0 To nPairs - 1
        If baseNd.Get(nBase(i)) <> FE_SUCCESS Then
            Err.Raise 2207, , "Base GRID " & nBase(i) & " is missing. Wrong base model or CSV."
        End If
        If Abs(baseNd.x * 1000# - xMm(i)) > MAX_AXIAL_ERR_MM Then
            Err.Raise 2208, , "Base GRID " & nBase(i) & " has a different X coordinate from the pair CSV."
        End If
        If insertNd.Get(nInsert(i)) = FE_SUCCESS Then present = present + 1
    Next i

    CheckExistingLinks App, partners
    If present = 0 Then
        bdfPath = GetFilePath("insert_aligned_to_base1.bdf", "bdf", , "Select insert_aligned_to_base1.bdf", 0)
        If bdfPath = "" Then Exit Sub
        reply = MsgBox("Import the selected insert BDF into this FEMAP model?", _
                       vbYesNo + vbQuestion, "Import HEX8 insert mesh")
        If reply <> vbYes Then Exit Sub
        CheckActive App, scratchName
        ImportInsertMesh App, bdfPath, insertIDs
    ElseIf present <> nPairs Then
        Err.Raise 2209, , CStr(present) & " / " & CStr(nPairs) & " insert GRIDs already exist. Check the model and pair CSV."
    End If

    CheckActive App, scratchName
    ' Check all pairs before adding their coincident GRIDs and connectors.
    worstAx = 0#: worstTang = 0#
    shortest = 1E+30: longest = 0#
    For i = 0 To nPairs - 1
        If baseNd.Get(nBase(i)) <> FE_SUCCESS Then
            Err.Raise 2210, , "Base GRID missing at pair " & (i + 1)
        End If
        If insertNd.Get(nInsert(i)) <> FE_SUCCESS Then
            Err.Raise 2211, , "Insert GRID " & nInsert(i) & " missing; import problem."
        End If
        ' Original deck and inserted mesh use METRES. Convert to mm for QA.
        dx = (insertNd.x - baseNd.x) * 1000#
        dy = (insertNd.y - baseNd.y) * 1000#
        dz = (insertNd.z - baseNd.z) * 1000#
        angle = angleDeg(i) * PI_NUM / 180#
        tangErr = Abs(dy * Sin(angle) - dz * Cos(angle))
        distanceMM = Sqr(dx * dx + dy * dy + dz * dz)
        If Abs(dx) > worstAx Then worstAx = Abs(dx)
        If tangErr > worstTang Then worstTang = tangErr
        If distanceMM < shortest Then shortest = distanceMM
        If distanceMM > longest Then longest = distanceMM
        If Abs(dx) > MAX_AXIAL_ERR_MM Or tangErr > MAX_TANG_ERR_MM Then
            Err.Raise 2212, , "Off-axis beam pair " & (i + 1) & _
                      "; base=" & nBase(i) & "; insert=" & nInsert(i) & _
                      "; axial/tangent errors mm=" & Format$(Abs(dx), "0.000000") & _
                      "/" & Format$(tangErr, "0.000000")
        End If
        If distanceMM < MIN_BEAM_LEN_MM Or distanceMM > MAX_BEAM_LEN_MM Then
            Err.Raise 2213, , "Beam length outside limits at pair " & (i + 1) & _
                      ", length=" & Format$(distanceMM, "0.000") & " mm."
        End If
    Next i

    reply = MsgBox(CStr(nPairs) & " UNIQUE 1:1 PAIRS validated." & vbCrLf & _
                   "Worst axial error: " & Format$(worstAx, "0.000000") & " mm" & vbCrLf & _
                   "Worst radial tangent error: " & Format$(worstTang, "0.000000") & " mm" & vbCrLf & _
                   "Beam lengths: " & Format$(shortest, "0.000") & " to " & _
                   Format$(longest, "0.000") & " mm" & vbCrLf & vbCrLf & _
                   "Create " & CStr(2 * nPairs) & " coincident GRIDs, " & CStr(2 * nPairs) & " RBE2 and " & _
                   CStr(nPairs) & " CBEAM with property " & beamPropID & "?", _
                   vbYesNo + vbQuestion, "Validation OK")
    If reply <> vbYes Then Exit Sub

    CheckActive App, scratchName
    nodeCountBefore = App.Info_Count(NODE_ENTITY)
    elemCountBefore = App.Info_Count(ELEMENT_ENTITY)
    groupCountBefore = App.Info_Count(GROUP_ENTITY)
    firstNode = App.Info_MaxID(NODE_ENTITY) + 1
    firstElem = App.Info_MaxID(ELEMENT_ENTITY) + 1
    groupID = 1
    If groupCountBefore > 0 Then groupID = App.Info_MaxID(GROUP_ENTITY) + 1
    If firstNode + 2 * nPairs - 1 > MAX_ENTITY_ID Or firstElem + 3 * nPairs - 1 > MAX_ENTITY_ID Or _
       groupID > MAX_ENTITY_ID Then Err.Raise 2214, , "Insufficient free FEMAP IDs for the glue connections."

    Set auxiliary = App.feNode
    auxiliary.Type = 0
    auxiliary.defCSys = 0
    Set rigid = App.feElem
    rigid.Type = RIGID_ELEMENT
    rigid.Topology = RIGID_LIST
    rigid.propID = 0
    rigid.RigidInterpolate = False
    rigid.RigidThermalExpansion = 0#
    For i = 0 To 5
        auxiliary.permBC(i) = False
        rigid.release(0, i) = True
    Next i

    Set beam = App.feElem
    beam.Type = LINEAR_BEAM
    beam.Topology = LINE2
    beam.propID = beamPropID
    beam.orientID = 0
    ' Beam axis is radial in YZ; global X is perpendicular to every beam.
    beam.orient(0) = 1#
    beam.orient(1) = 0#
    beam.orient(2) = 0#

    For i = 0 To nPairs - 1
        auxBase = firstNode + 2 * i
        auxInsert = auxBase + 1
        rigidID = firstElem + 3 * i

        addedNodes = 2 * i + 1
        AddCoincidentNode baseNd, auxiliary, nBase(i), auxBase
        addedNodes = 2 * i + 2
        AddCoincidentNode insertNd, auxiliary, nInsert(i), auxInsert

        addedElems = 3 * i + 1
        AddRBE2 rigid, rigidID, auxBase, nBase(i)
        addedElems = 3 * i + 2
        AddRBE2 rigid, rigidID + 1, auxInsert, nInsert(i)

        beam.Node(0) = auxBase
        beam.Node(1) = auxInsert
        addedElems = 3 * i + 3
        CheckStatus beam.Put(rigidID + 2), "write CBEAM " & CStr(rigidID + 2)
        If i = 0 Or i = nPairs - 1 Then
            VerifyConnection App, nBase(i), nInsert(i), auxBase, auxInsert, rigidID, beamPropID
        End If
        If (i + 1) Mod 500 = 0 Or i = nPairs - 1 Then
            App.feAppMessage 0, "Solid glue: added " & (i + 1) & " / " & nPairs & " pairs."
        End If
    Next i

    addedGroup = True
    AddGlueGroup App, groupID, firstNode, addedNodes, firstElem, addedElems
    If App.Info_Count(NODE_ENTITY) <> nodeCountBefore + addedNodes Or _
       App.Info_Count(ELEMENT_ENTITY) <> elemCountBefore + addedElems Or _
       App.Info_Count(GROUP_ENTITY) <> groupCountBefore + 1 Then
        Err.Raise 2216, , "The FEMAP counts differ from the expected solid glue addition."
    End If
    committed = True
    detail = "Added " & CStr(addedNodes) & " coincident GRIDs, " & CStr(2 * nPairs) & " RBE2 and " & _
             CStr(nPairs) & " CBEAM. Group " & CStr(groupID) & " (Solid Glue RBE2 " & CStr(firstElem) & ")."
    App.feAppMessage 0, detail
    App.feViewRegenerate 0
    MsgBox detail, vbInformation, "Solid glue completed"
    Exit Sub
Fail:
    detail = Err.Description
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If committed Then
        detail = detail & vbCrLf & "The solid glue group was stored."
    ElseIf addedNodes > 0 Or addedElems > 0 Or addedGroup Then
        If App.FullScratchName = scratchName Then
            If addedGroup Then DeleteRange App, GROUP_ENTITY, groupID, 1
            DeleteRange App, ELEMENT_ENTITY, firstElem, addedElems
            DeleteRange App, NODE_ENTITY, firstNode, addedNodes
            If App.Info_Count(NODE_ENTITY) = nodeCountBefore And _
               App.Info_Count(ELEMENT_ENTITY) = elemCountBefore And _
               App.Info_Count(GROUP_ENTITY) = groupCountBefore Then
                detail = detail & vbCrLf & "The glue entities from this run were removed."
            Else
                detail = detail & vbCrLf & "Review the new glue IDs before saving. Rollback counts differ."
            End If
            App.feViewRegenerate 0
        Else
            detail = detail & vbCrLf & "The active model changed. Review the original model."
        End If
    End If
    If Not (App Is Nothing) Then App.feAppMessage 0, "Solid glue stopped. " & detail
    MsgBox detail, vbCritical, "Solid glue stopped"
End Sub

Private Function ReadID(ByVal text As String, ByVal fieldName As String) As Long
    Dim digits As String
    digits = Trim$(text)
    If Len(digits) = 0 Or Len(digits) > 8 Or digits Like "*[!0-9]*" Then
        Err.Raise 2220, , "Invalid " & fieldName & "."
    End If
    ReadID = CLng(digits)
End Function

Private Sub CheckStatus(ByVal rc As Long, ByVal action As String)
    If rc <> FE_SUCCESS Then Err.Raise 2221, , "FEMAP could not " & action & " (code " & CStr(rc) & ")."
End Sub

Private Sub CheckActive(ByVal App As femap.model, ByVal scratchName As String)
    Dim active As femap.model
    Set active = feFemap()
    If App.FullScratchName <> scratchName Or active.FullScratchName <> scratchName Then
        Err.Raise 2222, , "Return to the Base model that was active when the macro started."
    End If
End Sub

Private Sub CheckExistingLinks(ByVal App As femap.model, ByVal partners As Object)
    Dim elements As femap.Set, existing As femap.Elem
    Dim dependentIDs As Variant, faces As Variant, weights As Variant, dofs As Variant
    Dim dependentCount As Long, i As Long, first As Long, gid As Long
    Set elements = App.feSet
    CheckStatus elements.AddRule(RIGID_ELEMENT, ELEMENT_BY_TYPE), "select existing rigid elements"
    CheckStatus elements.AddRule(LINEAR_BEAM, ELEMENT_BY_TYPE), "select existing beam elements"
    CheckStatus elements.AddRule(PARABOLIC_BEAM, ELEMENT_BY_TYPE), "select existing parabolic beam elements"
    If elements.Count = 0 Then Exit Sub
    Set existing = App.feElem
    Do While existing.NextInSet(elements.ID) = FE_SUCCESS
        If existing.Type = LINEAR_BEAM Or existing.Type = PARABOLIC_BEAM Then
            gid = existing.Node(0)
            If partners.Exists(gid) Then
                If partners.Item(gid) = existing.Node(1) Then
                    Err.Raise 2223, , "Pair already has direct BEAM " & CStr(existing.ID) & "."
                End If
            End If
        ElseIf existing.Topology = RIGID_LIST Then
            If existing.RigidInterpolate Then
                gid = existing.Node(0)
                If partners.Exists(gid) Then
                    Err.Raise 2224, , "GRID " & CStr(gid) & " is dependent in RBE3 " & CStr(existing.ID) & "."
                End If
            Else
                CheckStatus existing.GetNodeList(0, dependentCount, dependentIDs, faces, weights, dofs), _
                            "read existing RBE2 " & CStr(existing.ID)
                first = LBound(dependentIDs)
                For i = first To first + dependentCount - 1
                    gid = CLng(dependentIDs(i))
                    If partners.Exists(gid) Then
                        Err.Raise 2224, , "GRID " & CStr(gid) & " is already dependent in RBE2 " & CStr(existing.ID) & "."
                    End If
                Next i
            End If
        End If
    Loop
End Sub

Private Sub ImportInsertMesh(ByVal App As femap.model, ByVal bdfPath As String, ByVal insertIDs As Object)
    On Error GoTo Failed
    Dim ids(0 To 3) As Object, incoming(0 To 3) As femap.Set, live As femap.Set
    Dim kind(0 To 3) As Long, before(0 To 3) As Long, idArray() As Long
    Dim fNum As Integer, slot As Long, i As Long, j As Long, entityID As Long
    Dim text As String, card As String, idText As String, fields As Variant, key As Variant
    Dim scratchName As String, detail As String, attempted As Boolean, restored As Boolean

    scratchName = App.FullScratchName
    kind(0) = MATERIAL_ENTITY
    kind(1) = PROPERTY_ENTITY
    kind(2) = NODE_ENTITY
    kind(3) = ELEMENT_ENTITY
    For i = 0 To 3
        Set ids(i) = CreateObject("Scripting.Dictionary")
    Next i

    fNum = FreeFile
    Open bdfPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, text
        card = UCase$(Trim$(text))
        If card <> "" And card <> "BEGIN BULK" And card <> "CEND" And card <> "ENDDATA" And _
           Left$(card, 1) <> "$" And Left$(card, 1) <> "+" And Left$(card, 1) <> "*" Then
            If InStr(text, ",") > 0 Then
                fields = Split(text, ",")
                card = UCase$(Trim$(fields(0)))
                idText = fields(1)
            Else
                card = UCase$(Trim$(Left$(text, 8)))
                If Right$(card, 1) = "*" Then
                    card = Left$(card, Len(card) - 1)
                    idText = Mid$(text, 9, 16)
                Else
                    idText = Mid$(text, 9, 8)
                End If
            End If
            Select Case card
                Case "MAT1": slot = 0
                Case "PSOLID": slot = 1
                Case "GRID": slot = 2
                Case "CHEXA": slot = 3
                Case Else
                    Err.Raise 2225, , "Unexpected card " & card & ". Select the insert-only BDF from the Python generator."
            End Select
            entityID = ReadID(idText, card & " ID in insert BDF")
            If entityID <= 0 Or ids(slot).Exists(entityID) Then
                Err.Raise 2226, , "Invalid or duplicate " & card & " ID " & CStr(entityID) & " in insert BDF."
            End If
            ids(slot).Add entityID, True
        End If
    Loop
    Close #fNum
    fNum = 0

    For Each key In insertIDs.Keys
        If Not ids(2).Exists(key) Then
            Err.Raise 2227, , "Insert BDF is missing CSV GRID " & CStr(key) & "."
        End If
    Next key
    For i = 0 To 3
        If ids(i).Count = 0 Then Err.Raise 2228, , "The insert BDF is missing mesh or section data."
        ReDim idArray(0 To ids(i).Count - 1)
        j = 0
        For Each key In ids(i).Keys
            idArray(j) = CLng(key)
            j = j + 1
        Next key
        Set incoming(i) = App.feSet
        CheckStatus incoming(i).AddArray(j, idArray), "collect the insert BDF IDs"
        before(i) = App.Info_Count(kind(i))
        If before(i) > 0 Then
            Set live = App.feSet
            CheckStatus live.AddAll(kind(i)), "read current FEMAP IDs"
            If live.CountCommon(incoming(i).ID) > 0 Then
                Err.Raise 2229, , "Insert BDF has IDs already in use (entity type " & CStr(kind(i)) & ")."
            End If
        End If
    Next i

    attempted = True
    CheckStatus App.feFileReadNastran(8, bdfPath), "import the insert-only Nastran BDF"
    CheckActive App, scratchName
    For i = 0 To 3
        If App.Info_Count(kind(i)) <> before(i) + incoming(i).Count Then
            Err.Raise 2230, , "Imported entity count differs from the insert BDF (entity type " & CStr(kind(i)) & ")."
        End If
        Set live = App.feSet
        CheckStatus live.AddAll(kind(i)), "read imported FEMAP IDs"
        If live.CountCommon(incoming(i).ID) <> incoming(i).Count Then
            Err.Raise 2231, , "FEMAP did not preserve all insert BDF IDs."
        End If
    Next i
    Exit Sub

Failed:
    detail = Err.Description
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    If attempted Then
        If App.FullScratchName = scratchName Then
            For i = 3 To 0 Step -1
                App.feDelete kind(i), incoming(i).ID
            Next i
            restored = True
            For i = 0 To 3
                If App.Info_Count(kind(i)) <> before(i) Then restored = False
            Next i
            If restored Then
                detail = detail & vbCrLf & "The insert import was rolled back."
            Else
                detail = detail & vbCrLf & "Review the import before saving. Rollback counts differ."
            End If
        Else
            detail = detail & vbCrLf & "The active model changed during the insert import."
        End If
    End If
    On Error GoTo 0
    Err.Raise 2232, , detail
End Sub

Private Sub AddCoincidentNode(ByVal solidNode As femap.Node, ByVal newNode As femap.Node, _
                              ByVal solidID As Long, ByVal newID As Long)
    CheckStatus solidNode.Get(solidID), "read solid GRID " & CStr(solidID)
    newNode.x = solidNode.x
    newNode.y = solidNode.y
    newNode.z = solidNode.z
    newNode.outCSys = solidNode.outCSys
    CheckStatus newNode.Put(newID), "write coincident GRID " & CStr(newID)
End Sub

Private Sub AddRBE2(ByVal rigid As femap.Elem, ByVal elementID As Long, _
                    ByVal independentID As Long, ByVal dependentID As Long)
    Dim nodeIDs(0 To 0) As Long, dependentList As Variant
    nodeIDs(0) = dependentID
    dependentList = nodeIDs
    rigid.ID = elementID
    rigid.Node(0) = independentID
    CheckStatus rigid.PutNodeList(0, 1, dependentList, Null, Null, Null), _
                "set dependent GRID for RBE2 " & CStr(elementID)
    CheckStatus rigid.Put(elementID), "write RBE2 " & CStr(elementID)
End Sub

Private Sub VerifyConnection(ByVal App As femap.model, ByVal baseID As Long, ByVal insertID As Long, _
                             ByVal auxBase As Long, ByVal auxInsert As Long, ByVal rigidID As Long, _
                             ByVal propertyID As Long)
    Dim original As femap.Node, auxiliary As femap.Node, stored As femap.Elem
    Set original = App.feNode
    Set auxiliary = App.feNode
    Set stored = App.feElem
    VerifyCoincident original, auxiliary, baseID, auxBase
    VerifyCoincident original, auxiliary, insertID, auxInsert
    VerifyRBE2 stored, rigidID, auxBase, baseID
    VerifyRBE2 stored, rigidID + 1, auxInsert, insertID
    CheckStatus stored.Get(rigidID + 2), "read back CBEAM " & CStr(rigidID + 2)
    If stored.Type <> LINEAR_BEAM Or stored.Topology <> LINE2 Or stored.propID <> propertyID Or _
       stored.Node(0) <> auxBase Or stored.Node(1) <> auxInsert Then
        Err.Raise 2233, , "CBEAM connectivity could not be verified at element " & CStr(rigidID + 2) & "."
    End If
End Sub

Private Sub VerifyCoincident(ByVal original As femap.Node, ByVal auxiliary As femap.Node, _
                             ByVal originalID As Long, ByVal auxiliaryID As Long)
    CheckStatus original.Get(originalID), "read original GRID " & CStr(originalID)
    CheckStatus auxiliary.Get(auxiliaryID), "read coincident GRID " & CStr(auxiliaryID)
    If original.x <> auxiliary.x Or original.y <> auxiliary.y Or original.z <> auxiliary.z Then
        Err.Raise 2234, , "GRID " & CStr(auxiliaryID) & " differs from the original solid GRID coordinates."
    End If
End Sub

Private Sub VerifyRBE2(ByVal stored As femap.Elem, ByVal elementID As Long, _
                       ByVal independentID As Long, ByVal dependentID As Long)
    Dim n As Long, i As Long, ids As Variant, faces As Variant, weights As Variant, dofs As Variant
    CheckStatus stored.Get(elementID), "read back RBE2 " & CStr(elementID)
    If stored.Type <> RIGID_ELEMENT Or stored.Topology <> RIGID_LIST Or stored.RigidInterpolate Or _
       stored.Node(0) <> independentID Then
        Err.Raise 2235, , "RBE2 type or independent GRID differs at element " & CStr(elementID) & "."
    End If
    CheckStatus stored.GetNodeList(0, n, ids, faces, weights, dofs), "read back RBE2 dependent GRID"
    If n <> 1 Then Err.Raise 2236, , "RBE2 " & CStr(elementID) & " has a different dependent GRID count."
    If CLng(ids(LBound(ids))) <> dependentID Then
        Err.Raise 2236, , "RBE2 " & CStr(elementID) & " has a different dependent GRID."
    End If
    For i = 0 To 5
        If Not stored.release(0, i) Then
            Err.Raise 2237, , "RBE2 " & CStr(elementID) & " is missing a selected DOF."
        End If
    Next i
End Sub

Private Sub AddGlueGroup(ByVal App As femap.model, ByVal groupID As Long, _
                         ByVal firstNode As Long, ByVal nNodes As Long, _
                         ByVal firstElem As Long, ByVal nElems As Long)
    Dim nodes As femap.Set, elems As femap.Set, glue As femap.Group
    Set nodes = App.feSet
    Set elems = App.feSet
    Set glue = App.feGroup
    CheckStatus nodes.AddRange(firstNode, firstNode + nNodes - 1, 1), "select new coincident GRIDs"
    CheckStatus elems.AddRange(firstElem, firstElem + nElems - 1, 1), "select new glue elements"
    glue.title = "Solid Glue RBE2 " & CStr(firstElem)
    CheckStatus glue.SetAdd(NODE_ENTITY, nodes.ID), "group the coincident GRIDs"
    CheckStatus glue.SetAdd(ELEMENT_ENTITY, elems.ID), "group the RBE2 and CBEAM elements"
    CheckStatus glue.Put(groupID), "store the solid glue group"
End Sub

Private Sub DeleteRange(ByVal App As femap.model, ByVal entityType As Long, _
                        ByVal firstID As Long, ByVal count As Long)
    Dim ids As femap.Set
    If count = 0 Then Exit Sub
    Set ids = App.feSet
    CheckStatus ids.AddRange(firstID, firstID + count - 1, 1), "select the newly reserved IDs"
    App.feDelete entityType, ids.ID
End Sub
