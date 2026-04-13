' Attribute VB_Name = "ZoteroPMID"
Option Explicit

' ===========================================================================
' ZoteroPMID - 基于文档中的PMID标记自动插入Zotero引用
' ===========================================================================
'
' ===== 方案A：混合方案（推荐，Zotero原生集成） =====
'   Step1_ExtractAndDownload   → 提取PMID + 下载MEDLINE
'   (手动) Zotero导入           → 通过标识符添加
'   Step3_InsertNextCitation   → 逐个处理：删除占位符 + 调用Zotero对话框
'   （重复运行 Step3 直到全部完成）
'
' ===== 方案B：全自动方案（速度快，但Zotero集成度较低） =====
'   Step1_ExtractAndDownload   → 提取PMID + 下载MEDLINE
'   (手动) Zotero导入           → 通过标识符添加
'   (Zotero JS) 生成映射        → 运行 ZoteroPMID_Zotero.js
'   Step3_AutoInsertAll        → 一键插入所有Zotero域代码
'   (手动) Zotero Refresh       → 格式化引用
'
' ===========================================================================

' ========================= 可配置项 =========================
Private Const WORK_FOLDER   As String = "ZoteroPMID"
Private Const FILE_MAPPING  As String = "pmid_mapping.txt"
Private Const FILE_NBIB     As String = "citations.nbib"
Private Const FILE_PMIDS    As String = "pmid_list.txt"

' 输出根目录："Downloads" = 下载文件夹, "Desktop" = 桌面
Private Const OUTPUT_BASE   As String = "Downloads"

' 替换时是否删除 [PMID] 前的空格（Vancouver 等上标样式应设为 True）
Private Const REMOVE_SPACE_BEFORE_CITE As Boolean = True

' PMID匹配规则：方括号内7-9位纯数字，支持逗号分隔多个
Private Const RE_BRACKET    As String = "\[(\d{7,9}(?:,\s*\d{7,9})*)\]"
Private Const RE_SINGLE     As String = "\d{7,9}"
' ============================================================


' ============================================================
'  步骤1：扫描文档提取PMID，下载MEDLINE数据
' ============================================================
Public Sub ExtractAndDownload()

    If Documents.Count = 0 Then
        MsgBox "Please open a Word document first.", vbExclamation
        Exit Sub
    End If

    Dim wf As String
    wf = WorkPath()
    EnsureDir wf

    Dim reB As Object: Set reB = NewRE(RE_BRACKET)
    Dim reS As Object: Set reS = NewRE(RE_SINGLE)

    ' --- Scan every paragraph for [PMID] patterns ---
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        Dim txt As String: txt = para.Range.Text
        If reB.Test(txt) Then
            Dim bms As Object: Set bms = reB.Execute(txt)
            Dim bm As Object
            For Each bm In bms
                Dim singles As Object: Set singles = reS.Execute(bm.Value)
                Dim s As Object
                For Each s In singles
                    If Not dict.Exists(s.Value) Then dict.Add s.Value, True
                Next s
            Next bm
        End If
    Next para

    If dict.Count = 0 Then
        MsgBox "No PMID found (pattern: [7-9 digit number]).", vbExclamation
        Exit Sub
    End If

    ' --- Build array ---
    Dim arr() As String: ReDim arr(dict.Count - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In dict.Keys
        arr(i) = CStr(k): i = i + 1
    Next k

    ' --- Save PMID list ---
    PutFile wf & FILE_PMIDS, Join(arr, vbCrLf)

    ' --- Copy to clipboard (for Zotero "Add Item by Identifier") ---
    SetClip Join(arr, vbCrLf)

    ' --- Download MEDLINE from PubMed E-utilities (backup import method) ---
    Dim dlOK As Boolean: dlOK = False
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", _
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=" & _
        Join(arr, ",") & "&rettype=medline&retmode=text", False
    http.send
    If Err.Number = 0 And http.Status = 200 Then
        PutFileUTF8 wf & FILE_NBIB, http.responseText
        dlOK = True
    End If
    On Error GoTo 0

    ' --- Report ---
    Dim msg As String
    msg = "Step 1 Done! Extracted " & dict.Count & " unique PMIDs." & vbCrLf & vbCrLf
    msg = msg & "Output folder: " & wf & vbCrLf
    msg = msg & IIf(dlOK, "MEDLINE file downloaded (citations.nbib).", _
                          "MEDLINE download failed (can be ignored).") & vbCrLf
    msg = msg & "PMIDs copied to clipboard." & vbCrLf & vbCrLf
    msg = msg & "===== NEXT STEPS =====" & vbCrLf
    msg = msg & "Method A (Recommended):" & vbCrLf
    msg = msg & "  1. Open Zotero, select your collection (e.g. Research > Citation)" & vbCrLf
    msg = msg & "  2. Click the magic wand icon (Add Item by Identifier)" & vbCrLf
    msg = msg & "  3. Ctrl+V to paste PMIDs, then press Enter" & vbCrLf & vbCrLf
    msg = msg & "Method B (Backup):" & vbCrLf
    msg = msg & "  Zotero > File > Import > select citations.nbib" & vbCrLf & vbCrLf
    msg = msg & "After import:" & vbCrLf
    msg = msg & "  Option A: Run Step3_InsertNextCitation (recommended)" & vbCrLf
    msg = msg & "  Option B: Run Zotero JS + Step3_AutoInsertAll (full-auto)"

    MsgBox msg, vbInformation, "ZoteroPMID"

End Sub


' ============================================================
'  步骤3（方案A）：逐个插入 - 调用Zotero原生对话框
' ============================================================
' 每次运行处理一个引用，重复运行直到全部完成。
' 可分配快捷键（如 Ctrl+Shift+Z）加速操作。
' ============================================================
Public Sub InsertNextCitation()

    ActiveWindow.View.ShowFieldCodes = False

    ' --- Use Word's native Find to locate next [PMID] ---
    ' Word Find skips field codes automatically, avoiding false matches.
    Selection.HomeKey Unit:=wdStory

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "\[[0-9, ]{7,}\]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    If Not Selection.Find.Execute Then
        Application.StatusBar = ""
        MsgBox "All done! No more [PMID] placeholders found." & vbCrLf & vbCrLf & _
               "If you haven't yet: click Zotero > Add/Edit Bibliography" & vbCrLf & _
               "to insert the reference list at the end.", _
               vbInformation, "ZoteroPMID"
        Exit Sub
    End If

    ' --- Scroll to show the match ---
    ActiveWindow.ScrollIntoView Selection.Range, True

    ' --- Extract and validate PMIDs via regex ---
    Dim rawText As String: rawText = Selection.Text
    Dim reS As Object: Set reS = NewRE(RE_SINGLE)
    Dim pmids As Object: Set pmids = reS.Execute(rawText)

    If pmids.Count = 0 Then Exit Sub

    ' --- Multi-PMID: split [A, B] into [A][B], then re-find first one ---
    If pmids.Count > 1 Then
        Dim newText As String: newText = ""
        Dim i As Long
        For i = 0 To pmids.Count - 1
            newText = newText & "[" & pmids(i).Value & "]"
        Next i
        Selection.Text = newText
        Selection.Collapse Direction:=wdCollapseStart
        Selection.Find.Execute
    End If

    ' --- Handle preceding space ---
    If REMOVE_SPACE_BEFORE_CITE Then
        Dim s As Long: s = Selection.Start
        If s > 0 Then
            Dim chk As Range: Set chk = ActiveDocument.Range(s - 1, s)
            If chk.Text = " " Then Selection.SetRange s - 1, Selection.End
        End If
    End If

    ' --- Copy PMID to clipboard ---
    Dim pmid As String: pmid = pmids(0).Value
    SetClip pmid

    ' --- Delete placeholder ---
    Selection.Delete

    ' --- Status bar hint ---
    Application.StatusBar = "PMID " & pmid & _
        " copied to clipboard. Paste in Zotero search bar, select the item, press Enter."

    ' --- Call Zotero's native citation dialog ---
    On Error Resume Next
    Application.Run "ZoteroAddEditCitation"
    If Err.Number <> 0 Then
        MsgBox "ZoteroAddEditCitation failed." & vbCrLf & _
               "Make sure Zotero is running with the Word plugin.", vbExclamation
        Err.Clear
    End If
    On Error GoTo 0

End Sub



' ============================================================
'  步骤3（方案B）：一键自动插入所有引用域代码
' ============================================================
' 需要先运行 Zotero JS 映射脚本生成 pmid_mapping.txt。
' 速度快但Zotero集成度较低（切换样式可能需手动清除格式）。
' ============================================================
Public Sub AutoInsertAll()

    Dim wf As String: wf = WorkPath()
    Dim mapPath As String: mapPath = wf & FILE_MAPPING

    If Dir(mapPath) = "" Then
        MsgBox "Mapping file not found:" & vbCrLf & mapPath & vbCrLf & vbCrLf & _
               "Please run the Zotero JS mapping script first.", vbExclamation
        Exit Sub
    End If

    ' --- Load PMID -> (itemID, itemKey, URI) mapping ---
    Dim pMap As Object: Set pMap = LoadMapping(mapPath)
    If pMap.Count = 0 Then
        MsgBox "Mapping file is empty or has an invalid format.", vbExclamation
        Exit Sub
    End If

    ' Show field results (not codes) before scanning
    ActiveWindow.View.ShowFieldCodes = False

    Dim reB As Object: Set reB = NewRE(RE_BRACKET)
    Dim reS As Object: Set reS = NewRE(RE_SINGLE)

    ' --- Collect all citation locations ---
    Dim sPos() As Long, ePos() As Long, bText() As String
    Dim n As Long: n = 0
    Dim skipped As Long: skipped = 0

    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        Dim txt As String: txt = para.Range.Text
        If reB.Test(txt) Then
            Dim bms As Object: Set bms = reB.Execute(txt)
            Dim bm As Object
            For Each bm In bms
                Dim singles As Object: Set singles = reS.Execute(bm.Value)
                Dim allMapped As Boolean: allMapped = True
                Dim s As Object
                For Each s In singles
                    If Not pMap.Exists(s.Value) Then
                        allMapped = False
                        skipped = skipped + 1
                        Exit For
                    End If
                Next s
                If allMapped Then
                    n = n + 1
                    ReDim Preserve sPos(1 To n)
                    ReDim Preserve ePos(1 To n)
                    ReDim Preserve bText(1 To n)
                    sPos(n) = para.Range.Start + bm.FirstIndex
                    ePos(n) = sPos(n) + bm.Length
                    bText(n) = bm.Value
                End If
            Next bm
        End If
    Next para

    If n = 0 Then
        Dim noMsg As String
        noMsg = "No replaceable PMID citations found."
        If skipped > 0 Then noMsg = noMsg & vbCrLf & skipped & " skipped (missing from mapping)."
        MsgBox noMsg, vbExclamation
        Exit Sub
    End If

    ' --- Confirm before modifying ---
    Dim cfm As String
    cfm = "Will replace " & n & " PMID citation(s) with Zotero field codes."
    If skipped > 0 Then cfm = cfm & vbCrLf & skipped & " location(s) skipped (missing from mapping)."
    cfm = cfm & vbCrLf & vbCrLf & "Please save a backup of your document first." & vbCrLf & "Continue?"

    If MsgBox(cfm, vbYesNo + vbQuestion, "ZoteroPMID") = vbNo Then Exit Sub

    ' --- Insert fields from end to start (avoids position shifts) ---
    Application.ScreenUpdating = False
    Dim ci As Long: ci = 0
    Dim j As Long

    For j = n To 1 Step -1
        Dim rangeStart As Long: rangeStart = sPos(j)
        Dim rangeEnd As Long: rangeEnd = ePos(j)

        ' Optionally remove the space immediately before [PMID]
        If REMOVE_SPACE_BEFORE_CITE And rangeStart > 0 Then
            Dim chk As Range
            Set chk = ActiveDocument.Range(rangeStart - 1, rangeStart)
            If chk.Text = " " Then rangeStart = rangeStart - 1
        End If

        Dim rng As Range
        Set rng = ActiveDocument.Range(rangeStart, rangeEnd)

        Set singles = reS.Execute(bText(j))
        ci = ci + 1
        Dim citeID As String: citeID = "zcite_" & Format(ci, "0000")
        Dim jsonStr As String: jsonStr = BuildCiteJSON(citeID, singles, pMap)

        rng.Text = ""

        Dim fld As Field
        Set fld = ActiveDocument.Fields.Add( _
            Range:=rng, _
            Type:=wdFieldEmpty, _
            Text:="ADDIN ZOTERO_ITEM CSL_CITATION " & jsonStr, _
            PreserveFormatting:=False)
    Next j

    ' --- Insert bibliography field at end of document ---
    InsertBibField

    ' --- Force all fields to display results instead of codes ---
    Dim f As Field
    For Each f In ActiveDocument.Fields
        f.ShowCodes = False
    Next f
    ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = True

    MsgBox "Step 3 Done!" & vbCrLf & _
           "Inserted " & n & " citation(s) + bibliography field." & vbCrLf & vbCrLf & _
           "FINAL STEP:" & vbCrLf & _
           "Click Zotero tab > Refresh in the Word toolbar." & vbCrLf & _
           "When prompted, select the Vancouver citation style.", _
           vbInformation, "ZoteroPMID"

End Sub


' ============================================================
'  Internal Helpers
' ============================================================

Private Function WorkPath() As String
    WorkPath = Environ("USERPROFILE") & "\" & OUTPUT_BASE & "\" & WORK_FOLDER & "\"
End Function

Private Function NewRE(pat As String) As Object
    Set NewRE = CreateObject("VBScript.RegExp")
    NewRE.Global = True
    NewRE.Pattern = pat
End Function

Private Sub EnsureDir(p As String)
    Dim clean As String: clean = p
    If Right(clean, 1) = "\" Then clean = Left(clean, Len(clean) - 1)
    If Dir(clean, vbDirectory) = "" Then MkDir clean
End Sub

' Build Zotero CSL_CITATION JSON for one citation location
Private Function BuildCiteJSON(cid As String, pmids As Object, pMap As Object) As String
    Dim js As String
    js = "{""citationID"":""" & cid & """,""properties"":{},""citationItems"":["

    Dim idx As Long: idx = 0
    Dim pm As Object
    For Each pm In pmids
        If idx > 0 Then js = js & ","
        Dim v As Variant: v = pMap(pm.Value)
        js = js & "{""id"":" & CStr(v(0)) & _
                   ",""uris"":[""" & CStr(v(2)) & """]" & _
                   ",""itemData"":{""id"":" & CStr(v(0)) & ",""type"":""article-journal""}}"
        idx = idx + 1
    Next pm

    js = js & "],""schema"":""https://github.com/citation-style-language/schema/raw/master/csl-citation.json""}"
    BuildCiteJSON = js
End Function

' Insert ZOTERO_BIBL field at the end of the document
Private Sub InsertBibField()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseEnd
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    ActiveDocument.Fields.Add _
        Range:=rng, _
        Type:=wdFieldEmpty, _
        Text:="ADDIN ZOTERO_BIBL {""uncited"":[],""omitted"":[],""custom"":[]} CSL_BIBLIOGRAPHY", _
        PreserveFormatting:=False
End Sub

' Load tab-separated mapping file: PMID  itemID  itemKey  URI
Private Function LoadMapping(fp As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim f As Integer: f = FreeFile
    Dim ln As String

    Open fp For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim(ln)
        If Len(ln) > 0 And Left(ln, 4) <> "PMID" Then
            Dim pt() As String: pt = Split(ln, vbTab)
            If UBound(pt) >= 3 Then
                d(pt(0)) = Array(pt(1), pt(2), pt(3))
            End If
        End If
    Loop
    Close #f

    Set LoadMapping = d
End Function

Private Sub PutFile(fp As String, s As String)
    Dim f As Integer: f = FreeFile
    Open fp For Output As #f
    Print #f, s
    Close #f
End Sub

Private Sub PutFileUTF8(fp As String, s As String)
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText s
        .SaveToFile fp, 2
        .Close
    End With
End Sub

Private Sub SetClip(s As String)
    On Error Resume Next
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText s
        .PutInClipboard
    End With
    On Error GoTo 0
End Sub
