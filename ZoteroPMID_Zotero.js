// ===========================================================================
// ZoteroPMID - Zotero Mapping Script (v5)
// Run in: Zotero > Tools > Developer > Run JavaScript
// ===========================================================================
// Strategy: Match PMIDs to Zotero items via DOI (since Zotero 8 does not
// store PMID in Extra field when importing by identifier).
//
// Data sources:
//   - pmid_list.txt    : PMID list (created by VBA Step1)
//   - citations.nbib   : MEDLINE data with PMID->DOI pairs (created by VBA Step1)
//   - Zotero collection: items with DOI fields
// ===========================================================================

// ========================= Configuration =========================
var WORK_FOLDER     = "ZoteroPMID";
var MAPPING_FILE    = "pmid_mapping.txt";
var PMID_LIST_FILE  = "pmid_list.txt";
var NBIB_FILE       = "citations.nbib";
var COLLECTION_PATH = ["Research", "Citation"];
var OUTPUT_BASE     = "Downloads";       // "Downloads" or "Desktop"
// =================================================================

function normDoi(doi) {
    if (!doi) return "";
    return doi.toLowerCase().trim()
              .replace(/^https?:\/\/doi\.org\//, "")
              .replace(/^doi:\s*/i, "");
}

function makeNsIFile(path) {
    var f = Components.classes["@mozilla.org/file/local;1"]
                .createInstance(Components.interfaces.nsIFile);
    f.initWithPath(path);
    return f;
}

var _log = [];

try {

    // ---- 1. Navigate to target collection ----
    var libraryID = Zotero.Libraries.userLibraryID;
    var topCols   = Zotero.Collections.getByLibrary(libraryID);
    var targetCol = null;

    for (let c of topCols) {
        if (c.name === COLLECTION_PATH[0] && !c.parentID) { targetCol = c; break; }
    }
    if (!targetCol) throw new Error("Collection '" + COLLECTION_PATH[0] + "' not found.");
    _log.push("[1] Found: " + COLLECTION_PATH[0]);

    for (let ci = 1; ci < COLLECTION_PATH.length; ci++) {
        let children = targetCol.getChildCollections();
        let found = false;
        for (let ch of children) {
            if (ch.name === COLLECTION_PATH[ci]) { targetCol = ch; found = true; break; }
        }
        if (!found) throw new Error("'" + COLLECTION_PATH[ci] + "' not found under '" + targetCol.name + "'.");
        _log.push("[1] Found: " + COLLECTION_PATH[ci]);
    }

    // ---- 2. Get items and build DOI -> item map ----
    var allItems = targetCol.getChildItems();
    var doiToItem = {};
    var itemCount = 0;

    for (let item of allItems) {
        if (item.isNote() || item.isAttachment()) continue;
        itemCount++;
        var doi = "";
        try { doi = normDoi(item.getField("DOI")); } catch(e) {}
        if (doi) doiToItem[doi] = item;
    }
    _log.push("[2] Items in collection: " + itemCount + ",  with DOI: " + Object.keys(doiToItem).length);

    // ---- 3. Resolve output base path ----
    var homePath = Services.dirsvc.get("Home", Ci.nsIFile).path;
    var basePath = PathUtils.join(homePath, OUTPUT_BASE);
    var wfPath   = PathUtils.join(basePath, WORK_FOLDER);
    _log.push("[3] Work folder: " + wfPath);

    // ---- 4. Read PMID list ----
    var pmidListPath = PathUtils.join(wfPath, PMID_LIST_FILE);
    var pmidFile = makeNsIFile(pmidListPath);
    if (!pmidFile.exists()) throw new Error("PMID list not found: " + pmidListPath);

    var pmids = Zotero.File.getContents(pmidFile)
                    .trim().split(/[\r\n]+/)
                    .filter(function(s) { return /^\d+$/.test(s.trim()); })
                    .map(function(s) { return s.trim(); });
    _log.push("[4] PMIDs from document: " + pmids.length);

    // ---- 5. Build PMID -> DOI mapping ----
    var pmidToDoi = {};
    var nbibPath = PathUtils.join(wfPath, NBIB_FILE);
    var nbibFile = makeNsIFile(nbibPath);

    if (nbibFile.exists()) {
        var nbibText  = Zotero.File.getContents(nbibFile);
        var curPmid   = null;
        var curDoi    = null;
        var nbibLines = nbibText.split("\n");

        for (let line of nbibLines) {
            let pm = line.match(/^PMID-\s+(\d+)/);
            let dm = line.match(/^AID\s*-\s+(10\.\S+)\s+\[doi\]/);

            if (pm) {
                if (curPmid && curDoi) pmidToDoi[curPmid] = normDoi(curDoi);
                curPmid = pm[1];
                curDoi  = null;
            } else if (dm && !curDoi) {
                curDoi = dm[1];
            }
        }
        if (curPmid && curDoi) pmidToDoi[curPmid] = normDoi(curDoi);
        _log.push("[5] PMID->DOI from MEDLINE: " + Object.keys(pmidToDoi).length);
    } else {
        _log.push("[5] MEDLINE file not found, trying PubMed API...");
        try {
            var apiUrl = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&id=" +
                         pmids.join(",") + "&retmode=json";
            var xhr = new XMLHttpRequest();
            xhr.open("GET", apiUrl, false);
            xhr.send();
            if (xhr.status === 200) {
                var data = JSON.parse(xhr.responseText);
                for (let pmid of pmids) {
                    var art = data.result && data.result[pmid];
                    if (art && art.articleids) {
                        for (let aid of art.articleids) {
                            if (aid.idtype === "doi") {
                                pmidToDoi[pmid] = normDoi(aid.value);
                                break;
                            }
                        }
                    }
                }
                _log.push("[5] PMID->DOI from PubMed API: " + Object.keys(pmidToDoi).length);
            } else {
                _log.push("[5] PubMed API HTTP " + xhr.status);
            }
        } catch(apiErr) {
            _log.push("[5] PubMed API error: " + apiErr.message);
        }
    }

    // ---- 6. Match PMIDs to Zotero items via DOI ----
    var outLines  = ["PMID\titemID\titemKey\tURI"];
    var matched   = 0;
    var unmatched = [];
    var noDoi     = [];

    for (let pmid of pmids) {
        var doi  = pmidToDoi[pmid];
        var item = null;

        if (doi && doiToItem[doi]) {
            item = doiToItem[doi];
        }

        if (item) {
            var uri = Zotero.URI.getItemURI(item);
            outLines.push(pmid + "\t" + item.id + "\t" + item.key + "\t" + uri);
            matched++;
        } else if (!doi) {
            noDoi.push(pmid);
        } else {
            unmatched.push(pmid + " (doi:" + doi + ")");
        }
    }

    _log.push("[6] Matched: " + matched + " / " + pmids.length);
    if (noDoi.length > 0)     _log.push("    No DOI in MEDLINE (" + noDoi.length + "): " + noDoi.join(", "));
    if (unmatched.length > 0) _log.push("    DOI not in Zotero (" + unmatched.length + "): " + unmatched.join(", "));

    if (matched === 0) throw new Error("No PMID could be matched to a Zotero item.");

    // ---- 7. Write mapping file ----
    var content = outLines.join("\r\n");
    var outDir  = makeNsIFile(wfPath);
    if (!outDir.exists()) outDir.create(Components.interfaces.nsIFile.DIRECTORY_TYPE, 0o755);

    var outPath = PathUtils.join(wfPath, MAPPING_FILE);
    Zotero.File.putContents(makeNsIFile(outPath), content);

    _log.push("[7] Written to: " + outPath);
    _log.push("\n===== DONE =====");
    _log.push(matched + " / " + pmids.length + " PMID(s) mapped.");

} catch (ex) {
    _log.push("\n===== EXCEPTION =====");
    _log.push(ex.message);
}

_log.join("\n");
