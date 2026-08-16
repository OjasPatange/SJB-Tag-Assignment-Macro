Attribute VB_Name = "SJB_Tag_Assignment"
'==============================================================================
'  SJB TAG ASSIGNMENT MACRO  -  v14
'------------------------------------------------------------------------------
'  Auto-assigns Card No. / Channel No. to hardwired tags inside Smart Junction
'  Boxes (SJBs) and produces per-system assignment reports.
'
'  v14 change (over v13) -- the ONLY behavioural change in this release:
'    - SYSTEM IS NOW DERIVED FROM THE SJB NAME (Column AZ), NOT Column I.
'      The old Column-I "one-System-per-JB" segregation rule is SCRAPPED
'      (no more skipping mixed DCS/SIS JBs; no SegregationReport). Every SJB is
'      single-system by definition of its name. Mapping (case-insensitive):
'          name contains "PDSSB"  -> PDS
'          name contains "SSB"    -> SIS
'          name contains "DSB"    -> DCS
'          (anything else)        -> DCS   (fallback)
'      Order matters: "PDSSB" contains "SSB" as a substring, so PDSSB is tested
'      FIRST. ("DSB" does NOT appear inside "PDSSB".) See SystemFromSJB().
'      Column I is no longer used for system. A stale SegregationReport sheet
'      from a prior run is deleted at the start of AssignTags.
'      GDS: no SJB-name token was supplied, so GDS is not produced. If GDS boxes
'      exist, add their token to SystemFromSJB().
'
'  v13 change (over v12):
'    - PURE-INDEX card grouping for NON-TRAIN groups: key = GB|<index> (index
'      only), so tags that share the same _Gn share a card regardless of loop
'      base or plant area (group index -> card, loop -> channel).
'
'  v12 change (over v11):
'    - STANDALONE TRAIN TAGS may share a grouped card of the same train
'      (first-fit into the earliest same-train card with room). Three placement
'      phases: grouped -> standalone-train -> pure independents.
'
'  v11 change (over v10):
'    - GENERALISED TRAIN / FAMILY DETECTION: a train token is any suffix
'      _<Letter><digits> where <Letter> <> "G"; the FULL TOKEN is the identity
'      (ETH_T1, VRC_C1, N2_C1 ...). Compound TOKEN/loop_G<n> split on "/".
'
'  v10: GAP_BACKFILL toggle (Option A vs B).
'  v9 : main sheet READ-ONLY; reports-only output; bulk array I/O; "Tag Details".
'  v8 : header auto-detect; PDS own report; SIS/GDS class rules.
'
'  Placement / capacity rules:
'    SIS/GDS: 16 ch/card, 6/4/3 cards, usable 86/57/42, 12-AO cap.
'    DCS/PDS: 4 ch/card, 24/16/12 cards, usable 87/58/43, 2-AO cap, Rule 3b.
'    Rule 2 IS/NIS never share a card (ALL systems).
'    AssignedReport = SUMMARY only (per-SJB + TOTAL).
'
'  DORMANT: High-current DO id, 50 mA DO limit, CWA grouping,
'           SIS/GDS physical left/right side split.
'==============================================================================
Option Explicit

'==============================================================================
'  >>> GAP-BACKFILL TOGGLE  (flip this one value to switch strategy) <<<
'      True  = OPTION B : independent tags backfill gaps on grouped cards
'                         (denser packing; capacity - used = spares).  [DEFAULT]
'      False = OPTION A : grouped cards keep their gaps as SPARE.
'==============================================================================
Private Const GAP_BACKFILL As Boolean = True

'--- Main sheet resolution ---------------------------------------------------
Private Const MAIN_SHEET      As String = "Tag Details"
Private Const MAIN_SHEET_ALT  As String = "Sheet1"
Private HEADER_ROW As Long          ' auto-detected
Private FIRST_DATA As Long          ' = HEADER_ROW + 1

'--- Column map --------------------------------------------------------------
Private Const COL_TAG  As Long = 2    ' B  Tag_Number
Private Const COL_IO   As Long = 8    ' H  Io_Type
Private Const COL_SYS  As Long = 9    ' I  System (NO LONGER used - v14 derives from name)
Private Const COL_EXEC As Long = 18   ' R  CS_Execution (IS/NIS)
Private Const COL_GRP  As Long = 22   ' V  Grouping_Remarks
Private Const COL_SJB  As Long = 52   ' AZ SJB_Name  (system is derived from this)
Private Const COL_CARD As Long = 57   ' BE Card No.
Private Const COL_CH   As Long = 58   ' BF Channel No.

'--- Report sheet names ------------------------------------------------------
Private Const SH_CONFIG As String = "SJBConfig"
Private Const SH_ASSIGN As String = "AssignedReport"
Private Const SH_UNPLAC As String = "UnplacedReport"
Private Const SH_SEGREG As String = "SegregationReport"    ' deleted if present (v14)
Private Const SH_DCS    As String = "DCS Tag Assignment Report"
Private Const SH_SIS    As String = "SIS Tag Assignment Report"
Private Const SH_GDS    As String = "GDS Tag Assignment Report"
Private Const SH_PDS    As String = "PDS Tag Assignment Report"

Private Const SPARE_TAG As String = "SPARE"

'--- Per-tag record ----------------------------------------------------------
Private Type TagRec
    ro       As Long        ' 1-based offset into gData (row - FIRST_DATA + 1)
    tag      As String
    io       As String
    system   As String      ' v14: derived from SJB name, not Column I
    exec     As String
    grp      As String
    sjb      As String
    train    As String
    baseRem  As String
    grpIdx   As Long
    hasGrp   As Boolean
    IsAO     As Boolean
    isNHAO   As Boolean
    isMand   As Boolean
    grpKey   As String
    card     As Long
    ch       As Long
    placed   As Boolean
End Type

Private Type CardRec
    key       As String
    used      As Long
    aoCount   As Long
    hcCount   As Long
    execLock  As String
    trainLock As String
End Type

'--- Module-level source buffers (read once) ---------------------------------
Private gData As Variant     ' (1..nRows, 1..lastCol) data block, from FIRST_DATA
Private gHdr  As Variant     ' (1..1, 1..lastCol) header row
Private gLastCol As Long

'==============================================================================
'  ENTRY POINT
'==============================================================================
Public Sub AssignTags()
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = ResolveMainSheet()
    If ws Is Nothing Then
        MsgBox "Could not find a '" & MAIN_SHEET & "' (or '" & MAIN_SHEET_ALT & "') sheet.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    DetectHeaderRow ws
    ReadSourceBuffers ws            ' fills gData / gHdr / gLastCol (read-only)

    '--- v14: segregation scrapped; remove any stale SegregationReport --------
    DeleteSheetIfExists SH_SEGREG

    Dim tags() As TagRec, nTags As Long
    ReadTags tags, nTags
    If nTags = 0 Then
        MsgBox "No data rows found on '" & ws.Name & "'.", vbExclamation
        GoTo CleanExit
    End If

    Dim sjbOrder As Collection: Set sjbOrder = New Collection
    Dim firstSys As Object: Set firstSys = CreateObject("Scripting.Dictionary")
    BuildSjbMaps tags, nTags, sjbOrder, firstSys

    Dim capMap As Object: Set capMap = CreateObject("Scripting.Dictionary")
    ReadConfig capMap

    Dim spareRows As Collection: Set spareRows = New Collection   ' Array(sjb, card, ch)
    Dim unplaced As Collection: Set unplaced = New Collection
    Dim summary As Collection: Set summary = New Collection

    Dim i As Long, sjbName As String
    For i = 1 To sjbOrder.Count
        sjbName = sjbOrder(i)
        AssignOneSJB tags, nTags, sjbName, CStr(firstSys(sjbName)), _
                     CLng(capMap(sjbName)), spareRows, unplaced, summary
    Next i

    '--- Reports only. The main sheet is NOT modified. ----------------------
    BuildAssignedSummary summary
    BuildSystemReports tags, nTags, spareRows, firstSys
    BuildUnplacedReport unplaced

CleanExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "AssignTags (v14) complete." & vbCrLf & _
           "Source sheet: " & ws.Name & "  (read-only)" & vbCrLf & _
           "Header row: " & HEADER_ROW & vbCrLf & _
           "SJBs: " & sjbOrder.Count & vbCrLf & _
           "Unplaced tags: " & unplaced.Count, vbInformation
    Exit Sub
EH:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in AssignTags: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'==============================================================================
'  SHEET / HEADER / BUFFER SETUP
'==============================================================================
Private Function ResolveMainSheet() As Worksheet
    On Error Resume Next
    Set ResolveMainSheet = ThisWorkbook.Worksheets(MAIN_SHEET)
    If ResolveMainSheet Is Nothing Then Set ResolveMainSheet = ThisWorkbook.Worksheets(MAIN_SHEET_ALT)
    On Error GoTo 0
End Function

Private Sub DetectHeaderRow(ws As Worksheet)
    Dim r As Long
    HEADER_ROW = 4
    For r = 1 To 15
        If UCase$(Trim$(CStr(ws.Cells(r, COL_TAG).Value))) = "TAG_NUMBER" _
        And UCase$(Trim$(CStr(ws.Cells(r, COL_SJB).Value))) = "SJB_NAME" Then
            HEADER_ROW = r
            Exit For
        End If
    Next r
    FIRST_DATA = HEADER_ROW + 1
End Sub

Private Sub ReadSourceBuffers(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TAG).End(xlUp).Row
    gLastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If gLastCol < COL_CH Then gLastCol = COL_CH

    gHdr = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, gLastCol)).Value
    If lastRow >= FIRST_DATA Then
        gData = ws.Range(ws.Cells(FIRST_DATA, 1), ws.Cells(lastRow, gLastCol)).Value
    Else
        gData = Empty
    End If
End Sub

'--- safe string read from gData ---------------------------------------------
Private Function Gv(ByVal ro As Long, ByVal c As Long) As String
    Dim v As Variant
    v = gData(ro, c)
    If IsError(v) Then
        Gv = ""
    ElseIf IsEmpty(v) Then
        Gv = ""
    Else
        Gv = Trim$(CStr(v))
    End If
End Function

'==============================================================================
'  v14: SYSTEM FROM SJB NAME (Column AZ)
'    name contains "PDSSB" -> PDS   (checked FIRST: PDSSB contains "SSB")
'    name contains "SSB"   -> SIS
'    name contains "DSB"   -> DCS
'    else                  -> DCS   (fallback)
'==============================================================================
Private Function SystemFromSJB(ByVal sjbName As String) As String
    Dim s As String: s = UCase$(Trim$(sjbName))
    If InStr(s, "PDSSB") > 0 Then
        SystemFromSJB = "PDS"
    ElseIf InStr(s, "SSB") > 0 Then
        SystemFromSJB = "SIS"
    ElseIf InStr(s, "DSB") > 0 Then
        SystemFromSJB = "DCS"
    Else
        SystemFromSJB = "DCS"
    End If
End Function

'==============================================================================
'  READ / PARSE TAGS (from gData)
'==============================================================================
Private Sub ReadTags(ByRef tags() As TagRec, ByRef n As Long)
    n = 0
    If IsEmpty(gData) Then Exit Sub
    Dim nRows As Long: nRows = UBound(gData, 1)
    ReDim tags(1 To nRows)
    Dim ro As Long
    For ro = 1 To nRows
        Dim t As String, sj As String
        t = Gv(ro, COL_TAG)
        sj = Gv(ro, COL_SJB)
        If Len(t) = 0 And Len(sj) = 0 Then GoTo NextR
        If UCase$(t) = SPARE_TAG Then GoTo NextR      ' ignore any stale spares

        n = n + 1
        With tags(n)
            .ro = ro
            .tag = t
            .io = Gv(ro, COL_IO)
            .system = SystemFromSJB(sj)               ' v14: from name, not Col I
            .exec = UCase$(Gv(ro, COL_EXEC))
            .grp = Gv(ro, COL_GRP)
            .sjb = sj
            .IsAO = IsAO(.io)
            .isNHAO = IsNonHartAO(.io)
            ParseGrouping .grp, .train, .baseRem, .grpIdx, .hasGrp
            .isMand = (.hasGrp Or Len(.train) > 0)
            .grpKey = MakeGroupKey(tags(n))
            .card = 0: .ch = 0: .placed = False
        End With
NextR:
    Next ro
    If n = 0 Then Erase tags
End Sub

'--- PURE-INDEX card grouping for non-train groups (v13) ----------------------
'    Train group    -> "GT|<train>|<index>"  (per-train, index -> card)
'    Non-train group-> "GB|<index>"          (index -> card, ANY loop/area)
'    Train only     -> "TRAIN|<token>"
'    Independent    -> ""
Private Function MakeGroupKey(ByRef t As TagRec) As String
    If t.hasGrp Then
        If Len(t.train) > 0 Then
            MakeGroupKey = "GT|" & t.train & "|" & t.grpIdx        ' per-train + index
        Else
            MakeGroupKey = "GB|" & t.grpIdx                        ' index only (pure index)
        End If
    ElseIf Len(t.train) > 0 Then
        MakeGroupKey = "TRAIN|" & t.train
    Else
        MakeGroupKey = ""
    End If
End Function

Private Sub ParseGrouping(ByVal remark As String, ByRef train As String, _
                          ByRef baseRem As String, ByRef grpIdx As Long, _
                          ByRef hasGrp As Boolean)
    '--------------------------------------------------------------------------
    ' Two clean steps -- strip the group index first, then classify the
    ' remainder as a train/family token (compound via "/" or standalone).
    '--------------------------------------------------------------------------
    Dim s As String
    s = Trim$(remark)
    train = "": baseRem = s: grpIdx = 0: hasGrp = False
    If Len(s) = 0 Then Exit Sub

    '--- STEP 1: group index = last "_G" or "-G" followed by ALL DIGITS -------
    Dim up As String: up = UCase$(s)
    Dim delims(1) As String: delims(0) = "_G": delims(1) = "-G"
    Dim d As Long, p As Long, suffix As String
    For d = 0 To 1
        p = InStrRev(up, delims(d))
        If p > 0 Then
            suffix = Mid$(s, p + 2)
            If IsAllDigits(suffix) Then
                grpIdx = CLng(suffix)
                baseRem = Left$(s, p - 1)
                hasGrp = True
                Exit For
            End If
        End If
    Next d

    '--- STEP 2: train / family token on whatever is left ---------------------
    Dim rest As String
    rest = IIf(hasGrp, baseRem, s)

    Dim slashPos As Long: slashPos = InStr(rest, "/")
    If slashPos > 0 Then
        Dim lft As String, rgt As String
        lft = Trim$(Left$(rest, slashPos - 1))
        rgt = Trim$(Mid$(rest, slashPos + 1))
        If IsTrainToken(lft) Then
            train = UCase$(lft)         ' FULL token, e.g. ETH_T1 / VRC_C1
            baseRem = rgt               ' loop -> CHANNEL
        Else
            baseRem = rest              ' not a family: keep the whole remainder
        End If
    ElseIf IsTrainToken(rest) Then
        train = UCase$(rest)
        If Not hasGrp Then baseRem = ""  ' standalone token -> keyed TRAIN|<tok>
    End If
End Sub

'--- helper: True if s ends with "_<Letter><digits>" and Letter <> "G" ---------
'    ETH_T1 -> True (T)   VRC_C1 -> True (C)   N2_C1 -> True (C)
'    6121-TD-602 -> False (no underscore)   xxx_G1 -> False (G reserved)
Private Function IsTrainToken(ByVal s As String) As Boolean
    Dim t As String: t = Trim$(s)
    Dim p As Long: p = InStrRev(t, "_")
    If p = 0 Or p >= Len(t) Then Exit Function
    Dim letter As String: letter = UCase$(Mid$(t, p + 1, 1))
    If letter < "A" Or letter > "Z" Then Exit Function   ' must be a letter
    If letter = "G" Then Exit Function                   ' G = group index
    IsTrainToken = IsAllDigits(Mid$(t, p + 2))           ' "" -> False
End Function

'==============================================================================
'  SJB MAP  (v14: system derived from SJB name)
'==============================================================================
Private Sub BuildSjbMaps(ByRef tags() As TagRec, ByVal n As Long, _
                         ByRef sjbOrder As Collection, ByRef firstSys As Object)
    Dim i As Long, sjb As String
    For i = 1 To n
        sjb = tags(i).sjb
        If Len(sjb) = 0 Then GoTo NextI
        If Not firstSys.Exists(sjb) Then
            firstSys(sjb) = SystemFromSJB(sjb)
            sjbOrder.Add sjb, sjb
        End If
NextI:
    Next i
End Sub

Private Sub ReadConfig(ByRef capMap As Object)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        Dim sj As String, capv As Variant
        sj = Trim$(CStr(ws.Cells(r, 1).Value))
        capv = ws.Cells(r, 2).Value
        If Len(sj) > 0 Then
            If IsNumeric(capv) Then capMap(sj) = CLng(capv) Else capMap(sj) = 0
        End If
    Next r
End Sub

'==============================================================================
'  CORE ASSIGNMENT (per SJB)  -- in memory only
'    Three placement phases --
'      Phase 1  grouped tags (hasGrp)          -> PlaceGrouped   (index -> card)
'      Phase 2  standalone-train (train, !grp) -> PlaceStandaloneTrain (first-fit)
'      Phase 3  pure independents              -> PlaceIndependent (GAP_BACKFILL)
'==============================================================================
Private Sub AssignOneSJB(ByRef tags() As TagRec, ByVal n As Long, _
                         ByVal sjb As String, ByVal sys As String, _
                         ByVal capOverride As Long, _
                         ByRef spareRows As Collection, ByRef unplaced As Collection, _
                         ByRef summary As Collection)

    Dim cpc As Long: cpc = ChannelsPerCard(sys)
    Dim aoCap As Long: aoCap = AoCapPerCard(sys)

    Dim idx() As Long, m As Long
    ReDim idx(1 To n): m = 0
    Dim i As Long
    For i = 1 To n
        If tags(i).sjb = sjb Then m = m + 1: idx(m) = i
    Next i
    If m = 0 Then Exit Sub

    Dim cards() As CardRec
    Dim nCards As Long: nCards = 0
    ReDim cards(1 To m)

    ' Phase 1: grouped tags (group index -> card, loop -> channel)
    For i = 1 To m
        Dim t1 As Long: t1 = idx(i)
        If tags(t1).hasGrp Then PlaceGrouped tags, t1, sys, cpc, aoCap, cards, nCards
    Next i
    ' Phase 2: standalone-train tags -> first-fit into any same-train card
    For i = 1 To m
        Dim t2 As Long: t2 = idx(i)
        If (Not tags(t2).hasGrp) And Len(tags(t2).train) > 0 Then _
            PlaceStandaloneTrain tags, t2, sys, cpc, aoCap, cards, nCards
    Next i
    ' Phase 3: pure independent tags
    For i = 1 To m
        Dim t3 As Long: t3 = idx(i)
        If (Not tags(t3).hasGrp) And Len(tags(t3).train) = 0 Then _
            PlaceIndependent tags, t3, sys, cpc, aoCap, cards, nCards
    Next i

    Dim usedCount As Long: usedCount = m
    Dim capacity As Long
    capacity = ChooseCapacity(sys, capOverride, usedCount, nCards, cpc)
    Dim maxCards As Long: maxCards = capacity \ cpc

    Dim placedCount As Long: placedCount = 0
    For i = 1 To m
        Dim ti As Long: ti = idx(i)
        If tags(ti).card > maxCards Then
            tags(ti).card = 0: tags(ti).ch = 0: tags(ti).placed = False
            unplaced.Add Array(tags(ti).tag, sjb, tags(ti).grp, "Capacity exceeded")
        ElseIf tags(ti).card > 0 Then
            tags(ti).placed = True
            placedCount = placedCount + 1
        End If
    Next i

    ' SPARE fill: any channel still unoccupied after placement becomes SPARE.
    Dim occupied As Object: Set occupied = CreateObject("Scripting.Dictionary")
    For i = 1 To m
        ti = idx(i)
        If tags(ti).placed Then occupied(tags(ti).card & "-" & tags(ti).ch) = True
    Next i
    Dim spareCount As Long: spareCount = 0
    Dim c As Long, ch As Long
    For c = 1 To maxCards
        For ch = 1 To cpc
            If Not occupied.Exists(c & "-" & ch) Then
                spareRows.Add Array(sjb, c, ch)
                spareCount = spareCount + 1
            End If
        Next ch
    Next c

    Dim unplacedHere As Long: unplacedHere = usedCount - placedCount
    Dim status As String, pctSpare As Double
    pctSpare = IIf(capacity > 0, spareCount / capacity, 0)
    Dim reserve As Long: reserve = capacity - Usable(sys, capacity)
    If unplacedHere > 0 Then
        status = "OVER CAPACITY by " & unplacedHere
    ElseIf spareCount < reserve Then
        status = "Below 10% spare"
    Else
        status = "OK"
    End If

    summary.Add Array(sjb, sys, capacity, maxCards, placedCount, spareCount, pctSpare, status)
End Sub

Private Sub PlaceGrouped(ByRef tags() As TagRec, ByVal ti As Long, ByVal sys As String, _
                         ByVal cpc As Long, ByVal aoCap As Long, _
                         ByRef cards() As CardRec, ByRef nCards As Long)
    Dim key As String: key = tags(ti).grpKey
    Dim c As Long
    For c = 1 To nCards
        If cards(c).key = key Then
            If CanPlace(tags(ti), cards(c), sys, cpc, aoCap) Then
                DoPlace tags, ti, cards(c), c
                Exit Sub
            End If
        End If
    Next c
    nCards = nCards + 1
    cards(nCards).key = key
    cards(nCards).used = 0: cards(nCards).aoCount = 0: cards(nCards).hcCount = 0
    cards(nCards).execLock = "": cards(nCards).trainLock = ""
    DoPlace tags, ti, cards(nCards), nCards
End Sub

'--- standalone-train tag -> first-fit into any SAME-TRAIN card with room ------
Private Sub PlaceStandaloneTrain(ByRef tags() As TagRec, ByVal ti As Long, ByVal sys As String, _
                                 ByVal cpc As Long, ByVal aoCap As Long, _
                                 ByRef cards() As CardRec, ByRef nCards As Long)
    Dim c As Long
    For c = 1 To nCards
        If cards(c).trainLock = tags(ti).train Then
            If CanPlace(tags(ti), cards(c), sys, cpc, aoCap) Then
                DoPlace tags, ti, cards(c), c
                Exit Sub
            End If
        End If
    Next c
    nCards = nCards + 1
    cards(nCards).key = tags(ti).grpKey       ' "TRAIN|<token>"
    cards(nCards).used = 0: cards(nCards).aoCount = 0: cards(nCards).hcCount = 0
    cards(nCards).execLock = "": cards(nCards).trainLock = ""
    DoPlace tags, ti, cards(nCards), nCards
End Sub

Private Sub PlaceIndependent(ByRef tags() As TagRec, ByVal ti As Long, ByVal sys As String, _
                             ByVal cpc As Long, ByVal aoCap As Long, _
                             ByRef cards() As CardRec, ByRef nCards As Long)
    Dim c As Long
    For c = 1 To nCards
        ' OPTION A (GAP_BACKFILL=False): only fill independent cards (key = "").
        ' OPTION B (GAP_BACKFILL=True) : also backfill gaps on GROUPED cards.
        ' CanPlace still enforces AO cap, IS/NIS lock and train lock either way.
        If GAP_BACKFILL Or cards(c).key = "" Then
            If CanPlace(tags(ti), cards(c), sys, cpc, aoCap) Then
                DoPlace tags, ti, cards(c), c
                Exit Sub
            End If
        End If
    Next c
    nCards = nCards + 1
    cards(nCards).key = ""
    cards(nCards).used = 0: cards(nCards).aoCount = 0: cards(nCards).hcCount = 0
    cards(nCards).execLock = "": cards(nCards).trainLock = ""
    DoPlace tags, ti, cards(nCards), nCards
End Sub

Private Function CanPlace(ByRef t As TagRec, ByRef cd As CardRec, ByVal sys As String, _
                          ByVal cpc As Long, ByVal aoCap As Long) As Boolean
    CanPlace = False
    If cd.used >= cpc Then Exit Function
    If t.IsAO And cd.aoCount >= aoCap Then Exit Function
    If Len(t.train) > 0 And Len(cd.trainLock) > 0 Then
        If cd.trainLock <> t.train Then Exit Function
    End If
    If Len(t.exec) > 0 And Len(cd.execLock) > 0 Then
        If cd.execLock <> t.exec Then Exit Function
    End If
    If Not IsSisClass(sys) Then
        If (t.isNHAO Or IsHighCurrentDO(t.io)) And cd.hcCount >= 2 Then Exit Function
    End If
    CanPlace = True
End Function

Private Sub DoPlace(ByRef tags() As TagRec, ByVal ti As Long, ByRef cd As CardRec, ByVal cardNo As Long)
    cd.used = cd.used + 1
    tags(ti).card = cardNo
    tags(ti).ch = cd.used
    If tags(ti).IsAO Then cd.aoCount = cd.aoCount + 1
    If tags(ti).isNHAO Or IsHighCurrentDO(tags(ti).io) Then cd.hcCount = cd.hcCount + 1
    If Len(tags(ti).exec) > 0 And Len(cd.execLock) = 0 Then cd.execLock = tags(ti).exec
    If Len(tags(ti).train) > 0 And Len(cd.trainLock) = 0 Then cd.trainLock = tags(ti).train
End Sub

Private Function ChooseCapacity(ByVal sys As String, ByVal override As Long, _
                                ByVal usedCount As Long, ByVal cardsNeeded As Long, _
                                ByVal cpc As Long) As Long
    If override = 48 Or override = 64 Or override = 96 Then
        ChooseCapacity = override
        Exit Function
    End If
    Dim caps(2) As Long: caps(0) = 48: caps(1) = 64: caps(2) = 96
    Dim k As Long
    For k = 0 To 2
        If (caps(k) \ cpc) >= cardsNeeded And Usable(sys, caps(k)) >= usedCount Then
            ChooseCapacity = caps(k)
            Exit Function
        End If
    Next k
    ChooseCapacity = 96
End Function

'==============================================================================
'  REPORTS
'==============================================================================
Private Sub BuildAssignedSummary(ByRef summary As Collection)
    Dim ws As Worksheet: Set ws = FreshSheet(SH_ASSIGN)
    ws.Range("A1").Value = "SUMMARY - SJB Capacity & Spare Count"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:H3").Value = Array("SJB_Name", "System", "IO Capacity", "Total Cards", _
                                    "Used IOs", "Spare Count", "% Spare", "Status")
    ws.Range("A3:H3").Font.Bold = True
    Dim r As Long: r = 4
    Dim tUsed As Long, tSpare As Long, tCap As Long, tCards As Long
    Dim i As Long
    For i = 1 To summary.Count
        Dim a As Variant: a = summary(i)
        ws.Cells(r, 1).Value = a(0): ws.Cells(r, 2).Value = a(1)
        ws.Cells(r, 3).Value = a(2): ws.Cells(r, 4).Value = a(3)
        ws.Cells(r, 5).Value = a(4): ws.Cells(r, 6).Value = a(5)
        ws.Cells(r, 7).Value = Format$(a(6), "0.0%"): ws.Cells(r, 8).Value = a(7)
        tCap = tCap + a(2): tCards = tCards + a(3)
        tUsed = tUsed + a(4): tSpare = tSpare + a(5)
        r = r + 1
    Next i
    ws.Cells(r, 1).Value = "TOTAL"
    ws.Cells(r, 3).Value = tCap: ws.Cells(r, 4).Value = tCards
    ws.Cells(r, 5).Value = tUsed: ws.Cells(r, 6).Value = tSpare
    ws.Cells(r, 7).Value = IIf(tCap > 0, Format$(tSpare / tCap, "0.0%"), "")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Font.Bold = True
    ws.Columns("A:H").AutoFit
End Sub

Private Sub BuildSystemReports(ByRef tags() As TagRec, ByVal nTags As Long, _
                               ByRef spareRows As Collection, ByRef firstSys As Object)
    BuildOneSystemReport "DCS", SH_DCS, tags, nTags, spareRows, firstSys, True
    BuildOneSystemReport "SIS", SH_SIS, tags, nTags, spareRows, firstSys, True
    BuildOneSystemReport "GDS", SH_GDS, tags, nTags, spareRows, firstSys, False
    BuildOneSystemReport "PDS", SH_PDS, tags, nTags, spareRows, firstSys, False
End Sub

'--- one per-system report: DETAILED list only (no summary block) ------------
Private Sub BuildOneSystemReport(ByVal sys As String, ByVal sheetName As String, _
                                 ByRef tags() As TagRec, ByVal nTags As Long, _
                                 ByRef spareRows As Collection, _
                                 ByRef firstSys As Object, _
                                 ByVal alwaysBuild As Boolean)
    '--- count matching rows (assigned tags + spares) for this system --------
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = 1 To nTags
        If tags(i).placed Then
            If sysOf(firstSys, tags(i).sjb) = sys Then cnt = cnt + 1
        End If
    Next i
    Dim k As Long
    For k = 1 To spareRows.Count
        Dim sjbK As String: sjbK = spareRows(k)(0)
        If sysOf(firstSys, sjbK) = sys Then cnt = cnt + 1
    Next k

    If cnt = 0 And Not alwaysBuild Then
        DeleteSheetIfExists sheetName
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = FreshSheet(sheetName)
    ws.Range("A1").Value = sys & " TAG ASSIGNMENT REPORT (Card -> Channel, incl. SPARE)"
    ws.Range("A1").Font.Bold = True

    ' header row (full width) at row 3
    ws.Range(ws.Cells(3, 1), ws.Cells(3, gLastCol)).Value = gHdr
    ws.Rows(3).Font.Bold = True

    If cnt = 0 Then
        ws.Cells(4, 1).Value = "(no " & sys & " tags)"
        ws.Columns.AutoFit
        Exit Sub
    End If

    '--- fill output block ---------------------------------------------------
    Dim outArr() As Variant
    ReDim outArr(1 To cnt, 1 To gLastCol)
    Dim rIdx As Long: rIdx = 0
    Dim c As Long

    For i = 1 To nTags
        If tags(i).placed Then
            If sysOf(firstSys, tags(i).sjb) = sys Then
                rIdx = rIdx + 1
                For c = 1 To gLastCol
                    outArr(rIdx, c) = gData(tags(i).ro, c)
                Next c
                outArr(rIdx, COL_CARD) = tags(i).card
                outArr(rIdx, COL_CH) = tags(i).ch
            End If
        End If
    Next i
    For k = 1 To spareRows.Count
        Dim a As Variant: a = spareRows(k)
        If sysOf(firstSys, CStr(a(0))) = sys Then
            rIdx = rIdx + 1
            For c = 1 To gLastCol
                outArr(rIdx, c) = Empty
            Next c
            outArr(rIdx, COL_TAG) = SPARE_TAG
            outArr(rIdx, COL_SJB) = a(0)
            outArr(rIdx, COL_CARD) = a(1)
            outArr(rIdx, COL_CH) = a(2)
        End If
    Next k

    ws.Cells(4, 1).Resize(cnt, gLastCol).Value = outArr

    '--- sort by SJB -> Card -> Channel --------------------------------------
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(4, COL_SJB), order:=xlAscending
        .SortFields.Add key:=ws.Cells(4, COL_CARD), order:=xlAscending
        .SortFields.Add key:=ws.Cells(4, COL_CH), order:=xlAscending
        .SetRange ws.Range(ws.Cells(4, 1), ws.Cells(3 + cnt, gLastCol))
        .Header = xlNo
        .Apply
        .SortFields.Clear
    End With

    ws.Columns.AutoFit
End Sub

Private Function sysOf(ByRef firstSys As Object, ByVal sjb As String) As String
    If firstSys.Exists(sjb) Then sysOf = UCase$(CStr(firstSys(sjb))) Else sysOf = ""
End Function

Private Sub BuildUnplacedReport(ByRef unplaced As Collection)
    Dim ws As Worksheet: Set ws = FreshSheet(SH_UNPLAC)
    ws.Range("A1").Value = "UNPLACED REPORT - tags not assigned (collision / capacity)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:D3").Value = Array("Tag_Number", "SJB_Name", "Grouping_Remarks", "Reason")
    ws.Range("A3:D3").Font.Bold = True
    Dim r As Long: r = 4
    Dim i As Long
    For i = 1 To unplaced.Count
        ws.Cells(r, 1).Value = unplaced(i)(0)
        ws.Cells(r, 2).Value = unplaced(i)(1)
        ws.Cells(r, 3).Value = unplaced(i)(2)
        ws.Cells(r, 4).Value = unplaced(i)(3)
        r = r + 1
    Next i
    If unplaced.Count = 0 Then ws.Cells(4, 1).Value = "(none)"
    ws.Columns("A:D").AutoFit
End Sub

'==============================================================================
'  SJBConfig BUILDER  (v14: SYSTEM column derived from SJB name)
'==============================================================================
Public Sub BuildSJBConfig()
    Dim src As Worksheet: Set src = ResolveMainSheet()
    If src Is Nothing Then
        MsgBox "Could not find '" & MAIN_SHEET & "' (or '" & MAIN_SHEET_ALT & "').", vbCritical
        Exit Sub
    End If
    DetectHeaderRow src
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(SH_CONFIG)

    ' Column B (IO_Capacity) is left BLANK on refresh => AssignTags auto-sizes.
    ' Type 48/64/96 in column B ONLY to force a size for a specific SJB.
    Dim r As Long
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("SJB_Name", "IO_Capacity (48/64/96)", "Used IOs", "Auto Suggestion", "SYSTEM")
    ws.Range("A1:E1").Font.Bold = True

    Dim lastRow As Long: lastRow = src.Cells(src.Rows.Count, COL_TAG).End(xlUp).Row
    Dim used As Object: Set used = CreateObject("Scripting.Dictionary")
    Dim order As Collection: Set order = New Collection

    For r = FIRST_DATA To lastRow
        Dim s As String: s = Trim$(CStr(src.Cells(r, COL_SJB).Value))
        If Len(s) = 0 Then GoTo NextR
        If UCase$(Trim$(CStr(src.Cells(r, COL_TAG).Value))) = SPARE_TAG Then GoTo NextR
        If Not used.Exists(s) Then used(s) = 0: order.Add s, s
        used(s) = used(s) + 1
NextR:
    Next r

    Dim outR As Long: outR = 2
    Dim k
    For Each k In order
        Dim nm As String: nm = CStr(k)
        ws.Cells(outR, 1).Value = nm
        ws.Cells(outR, 3).Value = used(nm)
        ws.Cells(outR, 4).Value = AutoSuggest(SystemFromSJB(nm), CLng(used(nm)))
        ws.Cells(outR, 5).Value = SystemFromSJB(nm)         ' v14: system from name
        outR = outR + 1
    Next k
    ws.Columns("A:E").AutoFit
    MsgBox "SJBConfig refreshed (" & order.Count & " SJBs).", vbInformation
End Sub

Private Function AutoSuggest(ByVal sys As String, ByVal usedCount As Long) As String
    Dim caps(2) As Long: caps(0) = 48: caps(1) = 64: caps(2) = 96
    Dim k As Long
    For k = 0 To 2
        If Usable(sys, caps(k)) >= usedCount Then
            AutoSuggest = CStr(caps(k))
            Exit Function
        End If
    Next k
    AutoSuggest = "96 - OVER by " & (usedCount - Usable(sys, 96))
End Function

'==============================================================================
'  SYSTEM PARAMETER FUNCTIONS  (SIS/GDS = one class; DCS/PDS/other = DCS class)
'==============================================================================
Private Function IsSisClass(ByVal sys As String) As Boolean
    IsSisClass = (sys = "SIS" Or sys = "GDS")
End Function

Private Function ChannelsPerCard(ByVal sys As String) As Long
    ChannelsPerCard = IIf(IsSisClass(sys), 16, 4)
End Function

Private Function AoCapPerCard(ByVal sys As String) As Long
    AoCapPerCard = IIf(IsSisClass(sys), 12, 2)
End Function

Private Function Usable(ByVal sys As String, ByVal capacity As Long) As Long
    If IsSisClass(sys) Then
        Select Case capacity
            Case 96: Usable = 86
            Case 64: Usable = 57
            Case 48: Usable = 42
            Case Else: Usable = capacity
        End Select
    Else
        Select Case capacity
            Case 96: Usable = 87
            Case 64: Usable = 58
            Case 48: Usable = 43
            Case Else: Usable = capacity
        End Select
    End If
End Function

'==============================================================================
'  IO-TYPE PREDICATES
'==============================================================================
Private Function IsAO(ByVal io As String) As Boolean
    IsAO = (Left$(UCase$(Trim$(io)), 2) = "AO")
End Function

Private Function IsNonHartAO(ByVal io As String) As Boolean
    Dim s As String: s = UCase$(Trim$(io))
    If Left$(s, 2) = "AO" Then IsNonHartAO = (Right$(s, 1) <> "H")
End Function

Private Function IsHighCurrentDO(ByVal io As String) As Boolean
    IsHighCurrentDO = False      ' DORMANT stub
End Function

'==============================================================================
'  STRING / SHEET UTILITIES
'==============================================================================
Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        If Not IsDigit(Mid$(s, i, 1)) Then IsAllDigits = False: Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsDigit(ByVal c As String) As Boolean
    IsDigit = (c >= "0" And c <= "9")
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = nm
    End If
End Function

Private Function FreshSheet(ByVal nm As String) As Worksheet
    DeleteSheetIfExists nm
    Set FreshSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    FreshSheet.Name = nm
End Function

Private Sub DeleteSheetIfExists(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
