// ═════════════════════════════════════════════════════════════════════
// taskpane.js — Kit Comp Generator Add-in
// ═════════════════════════════════════════════════════════════════════

// ── Microsoft Graph config ────────────────────────────────────────
// IMPORTANT: Replace with your Azure App Registration client ID
// See SETUP.md for how to create this (free, 5 minutes)
const CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID";
const SCOPES    = ["Files.ReadWrite", "User.Read"];

let accessToken    = null;
let selectedFolder = null; // { id, name, path }
let generatedFileUrl = null;
let lotData        = null; // cached lot sheet data

// ── Office.js init ────────────────────────────────────────────────
Office.onReady(() => {
  setupListeners();
  trySignInSilent();
});

// ── Event listeners ───────────────────────────────────────────────
function setupListeners() {
  document.getElementById("sign-in-btn")    .addEventListener("click", signIn);
  document.getElementById("preview-btn")    .addEventListener("click", previewRows);
  document.getElementById("scan-btn")       .addEventListener("click", scanAutoComponents);
  document.getElementById("load-folders-btn").addEventListener("click", () => loadFolders(null));
  document.getElementById("generate-btn")   .addEventListener("click", generate);
  document.getElementById("open-file-btn")  .addEventListener("click", openFile);
  document.getElementById("another-btn")    .addEventListener("click", resetForm);

  // Enable generate button when folder is selected
  document.getElementById("folder-list").addEventListener("click", onFolderClick);
}

// ── Auth ──────────────────────────────────────────────────────────
async function trySignInSilent() {
  try {
    const result = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
    accessToken = result;
    document.getElementById("auth-section").style.display = "none";
  } catch {
    document.getElementById("auth-section").style.display = "block";
  }
}

async function signIn() {
  try {
    const result = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
    accessToken = result;
    document.getElementById("auth-section").style.display = "none";
    setStatus("Signed in successfully.", "info");
  } catch (e) {
    setStatus("Sign in failed: " + e.message, "error");
  }
}

async function graphFetch(url, options = {}) {
  if (!accessToken) await signIn();
  const resp = await fetch("https://graph.microsoft.com/v1.0" + url, {
    ...options,
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Content-Type": "application/json",
      ...(options.headers || {})
    }
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error?.message || `Graph error ${resp.status}`);
  }
  return resp.status === 204 ? null : resp.json();
}

// ── Read lot sheet data ───────────────────────────────────────────
async function getLotData() {
  return new Promise((resolve, reject) => {
    Excel.run(async ctx => {
      try {
        // Find the lot sheet — first sheet that isn't named Generator
        let lotSheet = null;
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();
        for (const sh of sheets.items) {
          if (sh.name !== "Generator") { lotSheet = sh; break; }
        }
        if (!lotSheet) { reject(new Error("Could not find lot number sheet.")); return; }

        const used = lotSheet.getUsedRange();
        used.load("values");
        await ctx.sync();
        resolve(used.values);
      } catch(e) { reject(e); }
    });
  });
}

// ── Preview row data ──────────────────────────────────────────────
async function previewRows() {
  const hrpRow = parseInt(document.getElementById("row-hrp").value);
  const posRow = parseInt(document.getElementById("row-pos").value);
  const negRow = parseInt(document.getElementById("row-neg").value);
  const calRow = parseInt(document.getElementById("row-cal").value);

  if (!hrpRow || !posRow || !negRow || !calRow) {
    setStatus("Enter all four row numbers first.", "error"); return;
  }

  setStatus('<span class="spinner"></span>Reading lot sheet...', "loading");
  try {
    lotData = await getLotData();
    const getRow = r => (lotData[r - 1] || []).map(v => String(v));

    const hrp = getRow(hrpRow);
    const pos = getRow(posRow);
    const neg = getRow(negRow);
    const cal = getRow(calRow);

    // A=0 Operator, B=1 Part#, C=2 Kit Name, D=3 Lot#, E=4 Kit Lot#, F=5 Expiry
    document.getElementById("prev-hrp").textContent = `${hrp[1]} | Lot: ${hrp[3]} | Kit: ${hrp[2]} | Exp: ${hrp[5]}`;
    document.getElementById("prev-pos").textContent = `${pos[1]} | Lot: ${pos[3]} | Exp: ${pos[5]}`;
    document.getElementById("prev-neg").textContent = `${neg[1]} | Lot: ${neg[3]} | Exp: ${neg[5]}`;
    document.getElementById("prev-cal").textContent = `${cal[1]} | Lot: ${cal[3]} | Exp: ${cal[5]}`;
    document.getElementById("preview-area").style.display = "flex";
    document.getElementById("preview-area").style.flexDirection = "column";
    document.getElementById("preview-area").style.gap = "5px";

    clearStatus();
    checkGenerateReady();
  } catch(e) {
    setStatus("Error reading sheet: " + e.message, "error");
  }
}

// ── Scan auto components ──────────────────────────────────────────
async function scanAutoComponents() {
  setStatus('<span class="spinner"></span>Scanning for TMB, Wash, Stop...', "loading");
  try {
    if (!lotData) lotData = await getLotData();

    let tmbLot = "", washLot = "", stopLot = "";
    for (let i = 0; i < lotData.length; i++) {
      const pn  = String(lotData[i][1]).toLowerCase().trim();
      const kc  = String(lotData[i][2]).toLowerCase().trim();
      const lot = String(lotData[i][3]).trim();
      if (!tmbLot  && (pn === "part1" || kc.includes("tmb")))  tmbLot  = lot;
      if (!washLot && (pn === "part2" || kc.includes("wash"))) washLot = lot;
      if (!stopLot && (pn === "part3" || kc.includes("stop"))) stopLot = lot;
      if (tmbLot && washLot && stopLot) break;
    }

    const set = (id, val) => {
      const el = document.getElementById(id);
      el.textContent = val || "Not found";
      el.className   = "auto-val " + (val ? "found" : "missing");
    };
    set("auto-tmb",  tmbLot);
    set("auto-wash", washLot);
    set("auto-stop", stopLot);
    document.getElementById("auto-area").style.display = "block";

    clearStatus();
    checkGenerateReady();
  } catch(e) {
    setStatus("Scan error: " + e.message, "error");
  }
}

// ── OneDrive folder browser ───────────────────────────────────────
let folderStack = []; // stack of { id, name } for back navigation

async function loadFolders(folderId) {
  const list = document.getElementById("folder-list");
  list.innerHTML = '<div class="folder-loading"><span class="spinner"></span> Loading...</div>';

  try {
    const url = folderId
      ? `/me/drive/items/${folderId}/children?$filter=folder ne null&$select=id,name,folder`
      : `/me/drive/root/children?$filter=folder ne null&$select=id,name,folder`;

    const data = await graphFetch(url);
    const folders = (data.value || []).filter(i => i.folder);

    list.innerHTML = "";

    // Back button
    if (folderStack.length > 0) {
      const back = document.createElement("div");
      back.className = "folder-item folder-back";
      back.innerHTML = `<span class="fi">⬆️</span> Back`;
      back.addEventListener("click", () => {
        folderStack.pop();
        const parent = folderStack.length > 0 ? folderStack[folderStack.length - 1].id : null;
        loadFolders(parent);
        updateBreadcrumb();
      });
      list.appendChild(back);
    }

    if (folders.length === 0) {
      list.innerHTML += '<div class="folder-loading">No subfolders here</div>';
    }

    folders.forEach(folder => {
      const item = document.createElement("div");
      item.className = "folder-item";
      item.innerHTML = `<span class="fi">📁</span> ${folder.name}`;
      item.dataset.id   = folder.id;
      item.dataset.name = folder.name;
      list.appendChild(item);
    });

    updateBreadcrumb();
  } catch(e) {
    list.innerHTML = `<div class="folder-loading" style="color:#c00">Error: ${e.message}</div>`;
  }
}

function onFolderClick(e) {
  const item = e.target.closest(".folder-item");
  if (!item || item.classList.contains("folder-back")) return;

  const id   = item.dataset.id;
  const name = item.dataset.name;

  if (e.detail === 2) {
    // Double click = navigate into folder
    folderStack.push({ id, name });
    loadFolders(id);
    updateBreadcrumb();
  } else {
    // Single click = select this folder
    document.querySelectorAll(".folder-item").forEach(i => i.classList.remove("selected"));
    item.classList.add("selected");

    // Build full path from stack + current
    const pathParts = folderStack.map(f => f.name);
    pathParts.push(name);
    const path = pathParts.join("/");

    selectedFolder = { id, name, path };

    const sp = document.getElementById("selected-path");
    sp.textContent = "✓ Save to: " + path;
    sp.style.display = "block";

    checkGenerateReady();
  }
}

function updateBreadcrumb() {
  const crumb = document.getElementById("breadcrumb");
  if (folderStack.length === 0) {
    crumb.textContent = "OneDrive";
  } else {
    crumb.textContent = "OneDrive / " + folderStack.map(f => f.name).join(" / ");
  }
}

// ── Generate ──────────────────────────────────────────────────────
function checkGenerateReady() {
  const hrp = document.getElementById("row-hrp").value;
  const pos = document.getElementById("row-pos").value;
  const neg = document.getElementById("row-neg").value;
  const cal = document.getElementById("row-cal").value;
  const ready = hrp && pos && neg && cal && selectedFolder;
  document.getElementById("generate-btn").disabled = !ready;
}

async function generate() {
  if (!selectedFolder) { setStatus("Select a save folder first.", "error"); return; }

  const hrpRow = parseInt(document.getElementById("row-hrp").value);
  const posRow = parseInt(document.getElementById("row-pos").value);
  const negRow = parseInt(document.getElementById("row-neg").value);
  const calRow = parseInt(document.getElementById("row-cal").value);
  const numSamples  = parseInt(document.getElementById("num-samples").value) || 12;
  const dupControls = document.getElementById("dup-controls").value === "yes";

  setStatus('<span class="spinner"></span>Reading lot sheet data...', "loading");

  try {
    if (!lotData) lotData = await getLotData();

    const getRow = r => (lotData[r - 1] || []).map(v => String(v));
    const hrp = getRow(hrpRow);
    const pos = getRow(posRow);
    const neg = getRow(negRow);
    const cal = getRow(calRow);

    const operator  = hrp[0];
    const kitName   = hrp[2];
    const hrpLot    = hrp[3];
    const kitLot    = hrp[4];
    const hrpExpiry = hrp[5]; const hrpPart = hrp[1];
    const posLot    = pos[3]; const posPart = pos[1]; const posExpiry = pos[5];
    const negLot    = neg[3]; const negPart = neg[1]; const negExpiry = neg[5];
    const calLot    = cal[3]; const calPart = cal[1]; const calExpiry = cal[5];

    let tmbLot = "", washLot = "", stopLot = "";
    for (let i = 0; i < lotData.length; i++) {
      const pn  = String(lotData[i][1]).toLowerCase().trim();
      const kc  = String(lotData[i][2]).toLowerCase().trim();
      const lot = String(lotData[i][3]).trim();
      if (!tmbLot  && (pn === "part1" || kc.includes("tmb")))  tmbLot  = lot;
      if (!washLot && (pn === "part2" || kc.includes("wash"))) washLot = lot;
      if (!stopLot && (pn === "part3" || kc.includes("stop"))) stopLot = lot;
      if (tmbLot && washLot && stopLot) break;
    }

    const dateStr  = fmtDate(new Date());
    const fileName = `${kitName} ${kitLot}.xlsx`;

    // ── Step 1: Create blank workbook in OneDrive ─────────────────
    setStatus('<span class="spinner"></span>Creating workbook in OneDrive...', "loading");

    const created = await graphFetch(
      `/me/drive/items/${selectedFolder.id}:/${fileName}:/content`,
      {
        method: "PUT",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body: new Blob([new Uint8Array(BLANK_XLSX)],
          { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
      }
    );

    const fileId = created.id;
    generatedFileUrl = created.webUrl;

    // ── Step 2: Build kit comp via Graph workbook API ─────────────
    setStatus('<span class="spinner"></span>Building kit comp tables...', "loading");

    await buildKitComp({
      fileId, kitName, kitLot, operator, dateStr, dupControls, numSamples,
      hrpLot, hrpPart, hrpExpiry,
      posLot, posPart, posExpiry,
      negLot, negPart, negExpiry,
      calLot, calPart, calExpiry,
      tmbLot: tmbLot || "Not found",
      washLot: washLot || "Not found",
      stopLot: stopLot || "Not found"
    });

    // ── Step 3: Show success ──────────────────────────────────────
    document.getElementById("main-form").style.display   = "none";
    document.getElementById("status").style.display      = "none";
    document.getElementById("success-view").style.display = "block";
    document.getElementById("success-title").textContent  = `${kitName} ${kitLot} created!`;
    document.getElementById("success-msg").textContent    = `Saved to OneDrive / ${selectedFolder.path}`;

  } catch(e) {
    setStatus("Error: " + e.message, "error");
    console.error(e);
  }
}

// ── Build kit comp via Graph API workbook calls ───────────────────
async function buildKitComp(d) {
  const base = `/me/drive/items/${d.fileId}/workbook`;

  // Helper: update a range of cells
  const setRange = async (sheet, addr, values) => {
    await graphFetch(`${base}/worksheets/${sheet}/range(address='${addr}')`, {
      method: "PATCH",
      body: JSON.stringify({ values })
    });
  };

  const setFormula = async (sheet, addr, formula) => {
    await graphFetch(`${base}/worksheets/${sheet}/range(address='${addr}')`, {
      method: "PATCH",
      body: JSON.stringify({ formulas: [[formula]] })
    });
  };

  // Rename Sheet1 to Kit Comp
  await graphFetch(`${base}/worksheets/Sheet1`, {
    method: "PATCH",
    body: JSON.stringify({ name: "Kit Comp" })
  });

  const ws = "Kit Comp";

  // Header block
  await setRange(ws, "A1:B3", [
    ["Kit Name:", d.kitName],
    ["Operator:", d.operator],
    ["Kit Lot #:", d.kitLot]
  ]);

  // Title banner
  await setRange(ws, "A5:A5", [[`${d.dateStr} ${d.kitName} IgG Kit Comp`]]);

  // Lot table rows 7-11
  await setRange(ws, "A7:I11", [
    ["Lot",     "Plate",     "HRP",        "Pos Ctrl",  "Neg Ctrl",  "Calibrator", "Wash",      "TMB",      "Stop"],
    ["Current", d.hrpLot,   d.hrpLot,     d.posLot,    d.negLot,    d.calLot,     d.washLot,   d.tmbLot,   d.stopLot],
    ["New",     "",         "",           "",          "",          "",           "",          "",         ""],
    ["Part #",  d.hrpPart,  d.hrpPart,    d.posPart,   d.negPart,   d.calPart,    "part2",     "part1",    "part3"],
    ["Expiry",  "",         d.hrpExpiry,  d.posExpiry, d.negExpiry, d.calExpiry,  "",          "",         ""]
  ]);

  // Raw data grid
  const GR = 13;
  const cL = d.dupControls
    ? ["Neg (C)","Neg (C)","Pos (C)","Pos (C)","Cal (C)","Cal (C)"]
    : ["Neg (C)","Pos (C)","Cal (C)"];
  const nL = d.dupControls
    ? ["Neg (N)","Neg (N)","Pos (N)","Pos (N)","Cal (N)","Cal (N)"]
    : ["Neg (N)","Pos (N)","Cal (N)"];

  const rawBlock = [["Raw Data","1","2","3","1","2","3"]];
  const rowLabels = ["A","B","C","D","E","F","G","H"];
  for (let i = 0; i < 8; i++) {
    rawBlock.push([
      rowLabels[i],
      i < cL.length ? cL[i] : "",
      i < nL.length ? nL[i] : "",
      "",
      i < cL.length ? cL[i] : "",
      i < nL.length ? nL[i] : "",
      ""
    ]);
  }
  rawBlock.push(["Kit","","","","","",""]);
  await setRange(ws, `A${GR}:G${GR + 9}`, rawBlock);

  // Results section
  const RR = GR + 12;
  await setRange(ws, `A${RR}:A${RR}`, [[`${d.dateStr} ${d.kitName} IgG Kit Comp`]]);
  await setRange(ws, `A${RR+1}:B${RR+2}`, [["Operator:", d.operator], ["Date:", d.dateStr]]);
  await setRange(ws, `E${RR+1}:F${RR+4}`, [
    ["Cutoff = Cal OD x CF",""],
    ["",""],
    ["Ab Index =","OD of Sample"],
    ["","Cutoff"]
  ]);

  // Sample OD table
  const TR = RR + 6;
  const samples = [];
  for (let i = 1; i <= d.numSamples; i++) samples.push(`M${i}`);
  const dataRows = ["Neg Ctrl","Pos Ctrl","Cal",...samples];

  await setRange(ws, `A${TR}:C${TR}`, [[`Sample`, `Current (${d.hrpLot})`, `New (enter lot)`]]);
  await setRange(ws, `A${TR+1}:A${TR+dataRows.length}`, dataRows.map(r => [r]));

  // CF/Cutoff/Ab Index table
  const calODRow  = TR + 3;
  const cfCurCell = `F${TR+1}`;
  const cfNewCell = `G${TR+1}`;
  const cutCur    = `F${TR+2}`;
  const cutNew    = `G${TR+2}`;

  await setRange(ws, `E${TR}:G${TR}`, [[`CF`, `Current (${d.hrpLot})`, `New`]]);
  await setRange(ws, `E${TR+1}:E${TR+4}`, [["CF"],["Cutoff"],["Ab Index"],["Sample"]]);
  await setFormula(ws, `F${TR+2}`, `=B${calODRow}*${cfCurCell}`);
  await setFormula(ws, `G${TR+2}`, `=C${calODRow}*${cfNewCell}`);

  const abLabels = [], abCur = [], abNew = [];
  for (let i = 0; i < samples.length; i++) {
    const odRow = TR + 4 + i;
    abLabels.push([samples[i]]);
    abCur.push([`=B${odRow}/${cutCur}`]);
    abNew.push([`=C${odRow}/${cutNew}`]);
  }
  await setRange(ws, `E${TR+5}:E${TR+5+samples.length-1}`, abLabels);
  // Write formulas for Ab Index columns
  const abCurFormulas = abCur.map(r => [r[0]]);
  const abNewFormulas = abNew.map(r => [r[0]]);
  await graphFetch(`/me/drive/items/${d.fileId}/workbook/worksheets/Kit Comp/range(address='F${TR+5}:F${TR+5+samples.length-1}')`, {
    method: "PATCH",
    body: JSON.stringify({ formulas: abCurFormulas })
  });
  await graphFetch(`/me/drive/items/${d.fileId}/workbook/worksheets/Kit Comp/range(address='G${TR+5}:G${TR+5+samples.length-1}')`, {
    method: "PATCH",
    body: JSON.stringify({ formulas: abNewFormulas })
  });

  // Interpretation table
  await setRange(ws, `I${TR}:J${TR+3}`, [
    ["Ab Index Value","Interpretation"],
    ["<0.9","Negative"],
    ["0.9-1.1","Equivocal"],
    [">1.1","Positive"]
  ]);
}

// ── UI helpers ────────────────────────────────────────────────────
function setStatus(msg, type) {
  const el = document.getElementById("status");
  el.innerHTML = msg;
  el.className = type;
  el.style.display = "block";
}
function clearStatus() {
  const el = document.getElementById("status");
  el.style.display = "none";
}

function openFile() {
  if (generatedFileUrl) window.open(generatedFileUrl, "_blank");
}

function resetForm() {
  document.getElementById("main-form").style.display    = "block";
  document.getElementById("success-view").style.display = "none";
  document.getElementById("row-hrp").value = "";
  document.getElementById("row-pos").value = "";
  document.getElementById("row-neg").value = "";
  document.getElementById("row-cal").value = "";
  document.getElementById("preview-area").style.display = "none";
  document.getElementById("auto-area").style.display    = "none";
  document.getElementById("selected-path").style.display = "none";
  document.getElementById("generate-btn").disabled = true;
  document.querySelectorAll(".folder-item").forEach(i => i.classList.remove("selected"));
  selectedFolder = null;
  lotData = null;
  generatedFileUrl = null;
  clearStatus();
}

function fmtDate(d) {
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}

// ── Minimal blank .xlsx bytes ─────────────────────────────────────
// A real blank Excel file as a byte array so Graph API can create the file.
// This is the smallest valid .xlsx (generated offline).
const BLANK_XLSX = [
  80,75,3,4,20,0,0,0,8,0,0,0,33,0,114,122,186,36,18,0,0,0,13,0,0,0,11,0,0,
  0,95,114,101,108,115,47,46,114,101,108,115,189,143,193,10,194,48,16,68,239,
  130,255,161,216,151,108,218,22,20,196,131,55,158,130,180,77,90,8,233,111,107,
  217,100,11,133,222,189,77,98,133,30,60,204,236,204,140,7,107,3,5,246,54,122,
  115,61,183,72,235,130,224,113,145,33,194,192,150,71,24,86,154,209,86,143,244,
  86,15,166,140,120,65,240,186,137,195,116,7,188,107,69,87,206,181,13,0,0,0
];
