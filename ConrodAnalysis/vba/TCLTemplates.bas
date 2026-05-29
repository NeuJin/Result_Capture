Attribute VB_Name = "TCLTemplates"
Option Explicit

' ============================================================
' TCLTemplates — Build complete TCL scripts from embedded
' template strings with token substitution.
'
' Tokens: <<RESULT_FILE>>, <<OUTPUT_FOLDER>>, <<SET_IDS_LIST>>,
'         <<SET_LABELS_DICT>>, <<LEGEND_*>>, <<VIEWS_*>>,
'         <<SCR_W>>, <<SCR_H>>, <<MANIFEST_PATH>>
' ============================================================

' ---- SHARED: screenshot + legend block (appended to both scripts) ----
Private Function GetScreenshotBlock() As String
    Dim t As String
    t = "" & vbLf
    t = t & "# ============================================================" & vbLf
    t = t & "# SCREENSHOT BLOCK — 6-view auto-fit onto node region" & vbLf
    t = t & "# ============================================================" & vbLf
    t = t & "hwi GetSessionHandle hvSession_s" & vbLf
    t = t & "hvSession_s GetProjectHandle hvProj_s" & vbLf
    t = t & "hvProj_s GetPageHandle hvPage_s 0" & vbLf
    t = t & "hvPage_s GetWindowHandle hvWin_s 0" & vbLf
    t = t & "hvWin_s GetClientHandle hvClient_s" & vbLf
    t = t & "hvWin_s GetModelHandle hvModel_s 0" & vbLf
    t = t & "hvClient_s GetResultCtrlHandle hvRC_s" & vbLf
    t = t & "" & vbLf
    t = t & "set manifest_path {<<MANIFEST_PATH>>}" & vbLf
    t = t & "set mfh [open $manifest_path w]" & vbLf
    t = t & "" & vbLf
    t = t & "# Node set IDs to process" & vbLf
    t = t & "set set_ids [list <<SET_IDS_LIST>>]" & vbLf
    t = t & "" & vbLf
    t = t & "# Map set_id -> display label" & vbLf
    t = t & "set set_labels [dict create <<SET_LABELS_DICT>>]" & vbLf
    t = t & "" & vbLf
    t = t & "# Map set_id -> summary value (filled by data extraction above)" & vbLf
    t = t & "# set_summary_values is a dict populated in the extraction loop" & vbLf
    t = t & "" & vbLf
    t = t & "foreach set_id $set_ids {" & vbLf
    t = t & "    set set_label [dict get $set_labels $set_id]" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 1. ISOLATE: hide all, then show only nodes of this set ---" & vbLf
    t = t & "    $hvModel_s SetMaskState 0 ; after 200" & vbLf
    t = t & "    set hvSSet [$hvModel_s GetSelectionSetHandle $set_id]" & vbLf
    t = t & "    set nodeIDs [$hvSSet GetNodeList]" & vbLf
    t = t & "    foreach nid $nodeIDs {" & vbLf
    t = t & "        $hvModel_s SetNodeMaskState $nid 1" & vbLf
    t = t & "    }" & vbLf
    t = t & "    after 300" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 2. APPLY LEGEND (user-defined, no auto-range) ---" & vbLf
    t = t & "    $hvRC_s SetLegendRange <<LEGEND_MIN>> <<LEGEND_MAX>>" & vbLf
    t = t & "    $hvRC_s SetLegendNumLevels <<LEGEND_LEVELS>>" & vbLf
    t = t & "    $hvRC_s SetLegendScale <<LEGEND_SCALE_INT>>" & vbLf
    t = t & "    $hvRC_s SetColorMap {<<LEGEND_PALETTE>>}" & vbLf
    t = t & "    $hvRC_s SetLegendShowValues <<LEGEND_SHOW_VAL>>" & vbLf
    t = t & "    $hvRC_s SetLegendStyle <<LEGEND_STYLE_INT>>" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 3. PREPARE output folder ---" & vbLf
    t = t & "    set safe_label [regsub -all {[^A-Za-z0-9_-]} $set_label {_}]" & vbLf
    t = t & "    set shot_dir {<<OUTPUT_FOLDER>>}/screenshots/$safe_label" & vbLf
    t = t & "    file mkdir $shot_dir" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 4. Get summary value for this set ---" & vbLf
    t = t & "    set sumval 0.0" & vbLf
    t = t & "    if {[dict exists $set_summary_values $set_id]} {" & vbLf
    t = t & "        set sumval [dict get $set_summary_values $set_id]" & vbLf
    t = t & "    }" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 5. CAPTURE 6 VIEWS (each: SetViewPoint -> Fit -> CaptureScreen) ---" & vbLf
    t = t & "    #   view_defs: {enabled vname eye_x eye_y eye_z up_x up_y up_z}" & vbLf
    t = t & "    set view_defs [list \\" & vbLf
    t = t & "        <<VIEWS_ISO>>   {ISO}    1  1  1  0  0  1 \\" & vbLf
    t = t & "        <<VIEWS_FRONT>> {Front}  0 -1  0  0  0  1 \\" & vbLf
    t = t & "        <<VIEWS_BACK>>  {Back}   0  1  0  0  0  1 \\" & vbLf
    t = t & "        <<VIEWS_LEFT>>  {Left}  -1  0  0  0  0  1 \\" & vbLf
    t = t & "        <<VIEWS_RIGHT>> {Right}  1  0  0  0  0  1 \\" & vbLf
    t = t & "        <<VIEWS_TOP>>   {Top}    0  0  1  0  1  0  \\" & vbLf
    t = t & "    ]" & vbLf
    t = t & "" & vbLf
    t = t & "    foreach {enabled vname ex ey ez ux uy uz} $view_defs {" & vbLf
    t = t & "        if {$enabled eq {1}} {" & vbLf
    t = t & "            $hvClient_s SetViewPoint $ex $ey $ez $ux $uy $uz ; after 150" & vbLf
    t = t & "            if {<<AUTO_FIT>> eq {1}} {" & vbLf
    t = t & "                $hvClient_s Fit ; after 300" & vbLf
    t = t & "            }" & vbLf
    t = t & "            set shot_path $shot_dir/$vname.png" & vbLf
    t = t & "            hwi CaptureScreen $shot_path <<SCR_W>> <<SCR_H>>" & vbLf
    t = t & "            # manifest line: set_id TAB label TAB view TAB path TAB value" & vbLf
    t = t & "            puts $mfh [format {%d\t%s\t%s\t%s\t%s} \\" & vbLf
    t = t & "                $set_id $set_label $vname $shot_path $sumval]" & vbLf
    t = t & "        }" & vbLf
    t = t & "    }" & vbLf
    t = t & "" & vbLf
    t = t & "    # --- 6. RESTORE full model visibility ---" & vbLf
    t = t & "    $hvModel_s SetMaskState 1 ; after 100" & vbLf
    t = t & "}" & vbLf
    t = t & "" & vbLf
    t = t & "close $mfh" & vbLf
    t = t & "puts {CONROD_DONE}" & vbLf
    t = t & "exit" & vbLf
    GetScreenshotBlock = t
End Function

' ============================================================
' GetStressTCL — Von Mises stress extraction + screenshots
' Based on TCL_StressExport.tcl by Nguyen Tan Loc
' ============================================================
Public Function GetStressTCL(cfg As PipelineCfg) As String
    Dim t As String

    t = "# ============================================================" & vbLf
    t = t & "# AUTO-GENERATED by ConrodAnalysis.xlsm" & vbLf
    t = t & "# Analysis: Von Mises Stress | Max per Node Set" & vbLf
    t = t & "# ============================================================" & vbLf
    t = t & "hwi OpenStack" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Session handles ---" & vbLf
    t = t & "hwi GetSessionHandle hvSession" & vbLf
    t = t & "hvSession GetProjectHandle hvProj" & vbLf
    t = t & "hvProj GetPageHandle hvPage 0" & vbLf
    t = t & "hvPage GetWindowHandle hvWin 0" & vbLf
    t = t & "hvWin GetClientHandle hvClient" & vbLf
    t = t & "hvWin GetModelHandle hvModel 0" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Load result file ---" & vbLf
    t = t & "hvModel SetResultFile {<<RESULT_FILE>>}" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Result controls: S-Stress Mises ---" & vbLf
    t = t & "hvClient GetResultCtrlHandle hvRC" & vbLf
    t = t & "hvRC SetDataType {S-Stress components}" & vbLf
    t = t & "hvRC SetDataComponent {Mises}" & vbLf
    t = t & "hvRC SetAverageMode 1" & vbLf
    t = t & "hvRC SetCornerDataEnabled 0" & vbLf
    t = t & "hvRC SetNumericPrecision 4" & vbLf
    t = t & "hvRC SetAvgAcrossPartsEnable 0" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Create derived load case from all subcases ---" & vbLf
    t = t & "hvRC GetSubcaseList subcaseList" & vbLf
    t = t & "set derivedCaseID [hvRC AddSubcase {Derived_Case}]" & vbLf
    t = t & "hvRC GetSubcaseHandle hvDerived $derivedCaseID" & vbLf
    t = t & "foreach sc $subcaseList {" & vbLf
    t = t & "    hvDerived AppendSimulation $sc" & vbLf
    t = t & "}" & vbLf
    t = t & "hvRC SetCurrentSubcase $derivedCaseID" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Summary dict: set_id -> max stress value ---" & vbLf
    t = t & "set set_summary_values [dict create]" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Query loop: iterate frames, track max stress per set ---" & vbLf
    t = t & "hvRC GetDerivedSimulationList simList" & vbLf
    t = t & "set numFrames [llength $simList]" & vbLf
    t = t & "set set_ids [list <<SET_IDS_LIST>>]" & vbLf
    t = t & "" & vbLf
    t = t & "# Init max trackers" & vbLf
    t = t & "set max_stress_per_set [dict create]" & vbLf
    t = t & "set max_node_per_set   [dict create]" & vbLf
    t = t & "foreach sid $set_ids {" & vbLf
    t = t & "    dict set max_stress_per_set $sid -1e30" & vbLf
    t = t & "    dict set max_node_per_set   $sid 0" & vbLf
    t = t & "}" & vbLf
    t = t & "" & vbLf
    t = t & "set frame_num 0" & vbLf
    t = t & "foreach sim $simList {" & vbLf
    t = t & "    incr frame_num" & vbLf
    t = t & "    hvRC SetCurrentSubcase $derivedCaseID" & vbLf
    t = t & "    hvRC SetCurrentSimulation $sim" & vbLf
    t = t & "" & vbLf
    t = t & "    # Per-frame CSV" & vbLf
    t = t & "    set fpath {<<OUTPUT_FOLDER>>}/csv/Stress_Frame[format %03d $frame_num].csv" & vbLf
    t = t & "    set fh [open $fpath w]" & vbLf
    t = t & "    puts $fh {SetID,NodeID,VonMises_MPa}" & vbLf
    t = t & "" & vbLf
    t = t & "    foreach sid $set_ids {" & vbLf
    t = t & "        hvRC GetQueryCtrlHandle hvQuery" & vbLf
    t = t & "        set hvSSet [hvModel GetSelectionSetHandle $sid]" & vbLf
    t = t & "        hvQuery SetSelectionSet $hvSSet" & vbLf
    t = t & "        hvQuery GetQuery queryResult" & vbLf
    t = t & "        set it [queryResult GetIteratorHandle]" & vbLf
    t = t & "        $it First" & vbLf
    t = t & "        while {![$it AtEnd]} {" & vbLf
    t = t & "            set nid  [$it GetNodeID]" & vbLf
    t = t & "            set val  [$it GetValue]" & vbLf
    t = t & "            puts $fh \"$sid,$nid,$val\"" & vbLf
    t = t & "            if {$val > [dict get $max_stress_per_set $sid]} {" & vbLf
    t = t & "                dict set max_stress_per_set $sid $val" & vbLf
    t = t & "                dict set max_node_per_set   $sid $nid" & vbLf
    t = t & "            }" & vbLf
    t = t & "            $it Next" & vbLf
    t = t & "        }" & vbLf
    t = t & "        hvQuery ReleaseHandle" & vbLf
    t = t & "    }" & vbLf
    t = t & "    close $fh" & vbLf
    t = t & "}" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Write summary CSV ---" & vbLf
    t = t & "set sumfh [open {<<OUTPUT_FOLDER>>}/csv/Stress_Summary.csv w]" & vbLf
    t = t & "puts $sumfh {SetID,Label,MaxVonMises_MPa,CriticalNodeID}" & vbLf
    t = t & "set set_labels [dict create <<SET_LABELS_DICT>>]" & vbLf
    t = t & "foreach sid $set_ids {" & vbLf
    t = t & "    set lbl [dict get $set_labels $sid]" & vbLf
    t = t & "    set val [dict get $max_stress_per_set $sid]" & vbLf
    t = t & "    set nid [dict get $max_node_per_set $sid]" & vbLf
    t = t & "    puts $sumfh \"$sid,$lbl,$val,$nid\"" & vbLf
    t = t & "    dict set set_summary_values $sid $val" & vbLf
    t = t & "}" & vbLf
    t = t & "close $sumfh" & vbLf
    t = t & "" & vbLf

    ' Append screenshot block
    t = t & GetScreenshotBlock()

    ' Substitute all tokens
    t = SubstituteTokens(t, cfg)
    GetStressTCL = t
End Function

' ============================================================
' GetSafetyFactorTCL — Endurance Safety Factor extraction
' Based on Conrod_SF_Find_NodeSet.tcl by Nguyen Tan Loc
' ============================================================
Public Function GetSafetyFactorTCL(cfg As PipelineCfg) As String
    Dim t As String

    t = "# ============================================================" & vbLf
    t = t & "# AUTO-GENERATED by ConrodAnalysis.xlsm" & vbLf
    t = t & "# Analysis: Endurance Safety Factor (Endure_SF_A) | Min per Node Set" & vbLf
    t = t & "# ============================================================" & vbLf
    t = t & "hwi OpenStack" & vbLf
    t = t & "" & vbLf
    t = t & "hwi GetSessionHandle hvSession" & vbLf
    t = t & "hvSession GetProjectHandle hvProj" & vbLf
    t = t & "hvProj GetPageHandle hvPage 0" & vbLf
    t = t & "hvPage GetWindowHandle hvWin 0" & vbLf
    t = t & "hvWin GetClientHandle hvClient" & vbLf
    t = t & "hvWin GetModelHandle hvModel 0" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Load result file ---" & vbLf
    t = t & "hvModel SetResultFile {<<RESULT_FILE>>}" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Contour: Endure_SF_A on Load Case 1 ---" & vbLf
    t = t & "hvClient GetResultCtrlHandle hvRC" & vbLf
    t = t & "hvRC GetContourCtrlHandle hvContour" & vbLf
    t = t & "hvContour SetDataType {Endure_SF_A}" & vbLf
    t = t & "hvRC SetCurrentSubcase 1" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Summary dict: set_id -> min SF value ---" & vbLf
    t = t & "set set_summary_values [dict create]" & vbLf
    t = t & "set set_ids [list <<SET_IDS_LIST>>]" & vbLf
    t = t & "set set_labels [dict create <<SET_LABELS_DICT>>]" & vbLf
    t = t & "" & vbLf
    t = t & "# --- Open summary CSV ---" & vbLf
    t = t & "set sumfh [open {<<OUTPUT_FOLDER>>}/csv/SafetyFactor_Summary.csv w]" & vbLf
    t = t & "puts $sumfh {SetID,Label,MinSafetyFactor,CriticalNodeID}" & vbLf
    t = t & "" & vbLf
    t = t & "foreach sid $set_ids {" & vbLf
    t = t & "    set hvSSet [hvModel GetSelectionSetHandle $sid]" & vbLf
    t = t & "    hvRC GetQueryCtrlHandle hvQuery" & vbLf
    t = t & "    hvQuery SetSelectionSet $hvSSet" & vbLf
    t = t & "    hvQuery GetQuery queryResult" & vbLf
    t = t & "" & vbLf
    t = t & "    set min_sf  1e30" & vbLf
    t = t & "    set min_nid 0" & vbLf
    t = t & "" & vbLf
    t = t & "    set it [queryResult GetIteratorHandle]" & vbLf
    t = t & "    $it First" & vbLf
    t = t & "    while {![$it AtEnd]} {" & vbLf
    t = t & "        set nid [$it GetNodeID]" & vbLf
    t = t & "        set val [$it GetValue]" & vbLf
    t = t & "        if {$val < $min_sf} {" & vbLf
    t = t & "            set min_sf  $val" & vbLf
    t = t & "            set min_nid $nid" & vbLf
    t = t & "        }" & vbLf
    t = t & "        $it Next" & vbLf
    t = t & "    }" & vbLf
    t = t & "    hvQuery ReleaseHandle" & vbLf
    t = t & "" & vbLf
    t = t & "    set lbl [dict get $set_labels $sid]" & vbLf
    t = t & "    puts $sumfh \"$sid,$lbl,$min_sf,$min_nid\"" & vbLf
    t = t & "    dict set set_summary_values $sid $min_sf" & vbLf
    t = t & "}" & vbLf
    t = t & "close $sumfh" & vbLf
    t = t & "" & vbLf

    ' Append screenshot block
    t = t & GetScreenshotBlock()

    t = SubstituteTokens(t, cfg)
    GetSafetyFactorTCL = t
End Function

' ============================================================
' SubstituteTokens — replace all <<TOKEN>> placeholders
' ============================================================
Private Function SubstituteTokens(t As String, cfg As PipelineCfg) As String
    ' Paths
    t = Replace(t, "<<RESULT_FILE>>",   PathToTCL(cfg.ResultFile))
    t = Replace(t, "<<OUTPUT_FOLDER>>", cfg.OutputFolder)  ' already forward-slashed
    t = Replace(t, "<<MANIFEST_PATH>>", cfg.OutputFolder & "/screenshot_manifest.txt")

    ' Set IDs list and labels dict
    t = Replace(t, "<<SET_IDS_LIST>>",    BuildSetIDsList(cfg))
    t = Replace(t, "<<SET_LABELS_DICT>>", BuildSetLabelsDict(cfg))

    ' Legend
    t = Replace(t, "<<LEGEND_MIN>>",       CStr(cfg.Leg.MinVal))
    t = Replace(t, "<<LEGEND_MAX>>",       CStr(cfg.Leg.MaxVal))
    t = Replace(t, "<<LEGEND_LEVELS>>",    CStr(cfg.Leg.NumLevels))
    t = Replace(t, "<<LEGEND_SCALE_INT>>", IIf(cfg.Leg.ScaleType = "Log", "1", "0"))
    t = Replace(t, "<<LEGEND_PALETTE>>",   cfg.Leg.Palette)
    t = Replace(t, "<<LEGEND_SHOW_VAL>>",  IIf(cfg.Leg.ShowValues, "1", "0"))
    t = Replace(t, "<<LEGEND_STYLE_INT>>", IIf(cfg.Leg.Discrete, "0", "1"))

    ' Views
    t = Replace(t, "<<VIEWS_ISO>>",   BoolToTCL(cfg.Views.ShowISO))
    t = Replace(t, "<<VIEWS_FRONT>>", BoolToTCL(cfg.Views.ShowFront))
    t = Replace(t, "<<VIEWS_BACK>>",  BoolToTCL(cfg.Views.ShowBack))
    t = Replace(t, "<<VIEWS_LEFT>>",  BoolToTCL(cfg.Views.ShowLeft))
    t = Replace(t, "<<VIEWS_RIGHT>>", BoolToTCL(cfg.Views.ShowRight))
    t = Replace(t, "<<VIEWS_TOP>>",   BoolToTCL(cfg.Views.ShowTop))
    t = Replace(t, "<<AUTO_FIT>>",    BoolToTCL(cfg.Views.AutoFit))

    ' Screenshot dimensions
    t = Replace(t, "<<SCR_W>>", CStr(cfg.Views.ScrW))
    t = Replace(t, "<<SCR_H>>", CStr(cfg.Views.ScrH))

    SubstituteTokens = t
End Function

' Build TCL list string: "1 2 3"
Private Function BuildSetIDsList(cfg As PipelineCfg) As String
    Dim ids As String : ids = ""
    Dim i As Integer
    For i = 0 To cfg.NodeSetCount - 1
        If ids = "" Then
            ids = CStr(cfg.NodeSets(i).SetID)
        Else
            ids = ids & " " & CStr(cfg.NodeSets(i).SetID)
        End If
    Next i
    BuildSetIDsList = ids
End Function

' Build TCL dict string: "1 {BigEnd} 2 {SmallEnd}"
Private Function BuildSetLabelsDict(cfg As PipelineCfg) As String
    Dim d As String : d = ""
    Dim i As Integer
    For i = 0 To cfg.NodeSetCount - 1
        d = d & CStr(cfg.NodeSets(i).SetID) & " {" & cfg.NodeSets(i).Label & "} "
    Next i
    BuildSetLabelsDict = Trim(d)
End Function
