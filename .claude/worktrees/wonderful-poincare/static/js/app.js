/* ===== HPM Proposal Generator — Main JS ===== */

"use strict";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _templateDefaults = {};    // loaded template JSON
let _isDirty = false;          // unsaved changes flag
let _currentDraftName = null;  // name of loaded draft (if any)
const _deletedSections = new Set();  // indices of sections deleted by user (Fix 3)

const ROMAN = ["I", "II", "III", "IV"];
const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
document.addEventListener("DOMContentLoaded", async () => {
  // Set today's date
  const dateInput = document.getElementById("date");
  if (dateInput && !dateInput.value) {
    dateInput.value = new Date().toISOString().slice(0, 10);
  }

  // Load template list
  await loadTemplateList();

  // Wire up fee fields
  document.querySelectorAll(".fee-input").forEach(inp => {
    inp.addEventListener("blur", () => fmtFeeBlur(inp));
    inp.addEventListener("input", updateTotal);
  });

  // Mark dirty on any input change
  document.querySelectorAll("input, textarea, select").forEach(el => {
    el.addEventListener("input", markDirty);
    el.addEventListener("change", markDirty);
  });

  // Cc toggle
  const ccChk = document.getElementById("include-cc");
  if (ccChk) ccChk.addEventListener("change", toggleCcField);
  toggleCcField();
});

// ---------------------------------------------------------------------------
// Template loading
// ---------------------------------------------------------------------------
async function loadTemplateList() {
  try {
    const res = await fetch("/api/templates");
    const templates = await res.json();
    const sel = document.getElementById("template-select");
    sel.innerHTML = "";
    templates.forEach(t => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.name;
      sel.appendChild(opt);
    });
    if (templates.length > 0) {
      await loadTemplate(templates[0].id, true);
    }
  } catch (e) {
    showToast("Could not load templates: " + e.message, "error");
  }
}

async function loadTemplate(templateId, initial = false) {
  try {
    const res = await fetch(`/api/template/${templateId}`);
    const tpl = await res.json();
    if (tpl.error) { showToast(tpl.error, "error"); return; }
    _templateDefaults = tpl;

    if (initial) {
      // Populate section headings
      tpl.section_headings.forEach((h, i) => {
        const el = document.getElementById(`section-heading-${i}`);
        if (el) el.value = h;
      });

      // Populate intro
      const introEl = document.getElementById("intro-paragraph");
      if (introEl && !introEl.value) introEl.value = tpl.default_intro || "";

      // Populate assumptions
      const assumList = document.getElementById("assumptions-list");
      if (assumList) {
        assumList.innerHTML = "";
        (tpl.default_assumptions || []).forEach(a => addAssumptionRow(a));
      }

      // Populate closing
      const closingEl = document.getElementById("closing-paragraph");
      if (closingEl && !closingEl.value) closingEl.value = tpl.default_closing || "";
    }
  } catch (e) {
    showToast("Failed to load template: " + e.message, "error");
  }
}

// ---------------------------------------------------------------------------
// Panel collapse/expand
// ---------------------------------------------------------------------------
function togglePanel(headerEl) {
  const body = headerEl.closest(".panel").querySelector(".panel-body");
  const arrow = headerEl.querySelector(".panel-toggle");
  body.classList.toggle("hidden");
  arrow.classList.toggle("collapsed");
}

function toggleSection(idx) {
  const body = document.getElementById(`section-body-${idx}`);
  const btn = document.getElementById(`section-toggle-${idx}`);
  if (!body) return;
  body.classList.toggle("hidden");
  btn.classList.toggle("collapsed");
}

// ---------------------------------------------------------------------------
// Section heading reset
// ---------------------------------------------------------------------------
function resetSectionHeading(idx) {
  const el = document.getElementById(`section-heading-${idx}`);
  if (el && _templateDefaults.section_headings) {
    el.value = _templateDefaults.section_headings[idx] || "";
    markDirty();
  }
}

// ---------------------------------------------------------------------------
// Fix 3: Section delete / restore
// ---------------------------------------------------------------------------
function deleteSection(idx) {
  _deletedSections.add(idx);
  const panel = document.getElementById(`section-panel-${idx}`);
  if (panel) panel.classList.add("section-deleted");
  _updateRestoreArea();
  markDirty();
}

function restoreSection(idx) {
  _deletedSections.delete(idx);
  const panel = document.getElementById(`section-panel-${idx}`);
  if (panel) panel.classList.remove("section-deleted");
  _updateRestoreArea();
  markDirty();
}

function _updateRestoreArea() {
  const area = document.getElementById("deleted-restore-area");
  if (!area) return;
  if (_deletedSections.size === 0) {
    area.classList.remove("visible");
  } else {
    area.classList.add("visible");
  }
  for (let i = 0; i < 4; i++) {
    const btn = document.getElementById(`restore-btn-${i}`);
    if (btn) {
      if (_deletedSections.has(i)) btn.classList.add("visible");
      else btn.classList.remove("visible");
    }
  }
}

// ---------------------------------------------------------------------------
// Sub-items
// ---------------------------------------------------------------------------
function addSubItem(sectionIdx) {
  const list = document.getElementById(`section-${sectionIdx}-items`);
  const idx = list.querySelectorAll(".sub-item-card").length;
  const card = _makeSubItemCard(sectionIdx, idx);
  list.appendChild(card);
  _reindexSubItems(sectionIdx);
  initDragDrop(list);
  markDirty();
}

function _makeSubItemCard(sectionIdx, itemIdx) {
  const card = document.createElement("div");
  card.className = "sub-item-card";
  card.draggable = true;

  card.innerHTML = `
    <div class="sub-item-top">
      <span class="drag-handle" title="Drag to reorder">&#8942;</span>
      <span class="sub-item-label">${LETTERS[itemIdx % 26]})</span>
      <div class="sub-item-content-wrap">
        <div class="rich-toolbar">
          <button class="rich-btn" onclick="richCmd(this,'bold')" title="Bold"><b>B</b></button>
          <button class="rich-btn" onclick="richCmd(this,'italic')" title="Italic"><i>I</i></button>
          <button class="rich-btn" onclick="richCmd(this,'underline')" title="Underline"><u>U</u></button>
          <button class="rich-btn" onclick="richCmd(this,'insertUnorderedList')" title="Bullet list">&#8226; List</button>
        </div>
        <div class="sub-item-editor" contenteditable="true" data-placeholder="Enter scope item..." spellcheck="true"></div>
      </div>
      <button class="btn-sm danger" onclick="deleteSubItem(this)" title="Delete item">&#10005;</button>
    </div>
    <div class="sub-item-actions">
      <button class="btn-sm" onclick="addSubSubItem(this, ${sectionIdx})">+ Sub-item</button>
      <button class="btn-sm btn-up-down" onclick="moveSubItem(this,-1)" title="Move up">&#9650;</button>
      <button class="btn-sm btn-up-down" onclick="moveSubItem(this,1)" title="Move down">&#9660;</button>
    </div>
    <div class="subitems-list"></div>
  `;

  const editor = card.querySelector(".sub-item-editor");
  editor.addEventListener("input", markDirty);

  return card;
}

function deleteSubItem(btn) {
  const card = btn.closest(".sub-item-card");
  const editor = card.querySelector(".sub-item-editor");
  const hasContent = editor && editor.innerText.trim().length > 0;

  if (hasContent) {
    if (!confirm("Delete this scope item and all its content?")) return;
  }
  const list = card.parentElement;
  card.remove();
  // Re-index all sibling cards
  const sectionIdx = _getSectionIdxFromList(list);
  _reindexSubItems(sectionIdx);
  markDirty();
}

function moveSubItem(btn, dir) {
  const card = btn.closest(".sub-item-card");
  const list = card.parentElement;
  if (dir === -1) {
    const prev = card.previousElementSibling;
    if (prev) list.insertBefore(card, prev);
  } else {
    const next = card.nextElementSibling;
    if (next) list.insertBefore(next, card);
  }
  const sectionIdx = _getSectionIdxFromList(list);
  _reindexSubItems(sectionIdx);
  markDirty();
}

function _reindexSubItems(sectionIdx) {
  const list = document.getElementById(`section-${sectionIdx}-items`);
  if (!list) return;
  list.querySelectorAll(".sub-item-card").forEach((card, i) => {
    const lbl = card.querySelector(".sub-item-label");
    if (lbl) lbl.textContent = `${LETTERS[i % 26]})`;
  });
}

function _getSectionIdxFromList(list) {
  const id = list.id; // section-N-items
  const match = id.match(/section-(\d+)-items/);
  return match ? parseInt(match[1]) : 0;
}

// ---------------------------------------------------------------------------
// Sub-sub-items
// ---------------------------------------------------------------------------
function addSubSubItem(btn, sectionIdx) {
  const card = btn.closest(".sub-item-card");
  const list = card.querySelector(".subitems-list");
  const idx = list.querySelectorAll(".subitem-card").length;

  const row = document.createElement("div");
  row.className = "subitem-card";
  row.draggable = true;
  row.innerHTML = `
    <span class="drag-handle" title="Drag to reorder">&#8942;</span>
    <span class="subitem-label">${idx + 1})</span>
    <div class="subitem-editor" contenteditable="true" data-placeholder="Sub-item..." spellcheck="true"></div>
    <button class="btn-sm btn-up-down" onclick="moveSubSubItem(this,-1)" title="Move up">&#9650;</button>
    <button class="btn-sm btn-up-down" onclick="moveSubSubItem(this,1)" title="Move down">&#9660;</button>
    <button class="btn-sm danger" onclick="deleteSubSubItem(this)" title="Delete">&#10005;</button>
  `;
  row.querySelector(".subitem-editor").addEventListener("input", markDirty);
  list.appendChild(row);
  initDragDrop(list);
  _reindexSubSubItems(list);
  markDirty();
}

function deleteSubSubItem(btn) {
  const row = btn.closest(".subitem-card");
  const list = row.parentElement;
  row.remove();
  _reindexSubSubItems(list);
  markDirty();
}

function moveSubSubItem(btn, dir) {
  const row = btn.closest(".subitem-card");
  const list = row.parentElement;
  if (dir === -1) {
    const prev = row.previousElementSibling;
    if (prev) list.insertBefore(row, prev);
  } else {
    const next = row.nextElementSibling;
    if (next) list.insertBefore(next, row);
  }
  _reindexSubSubItems(list);
  markDirty();
}

function _reindexSubSubItems(list) {
  list.querySelectorAll(".subitem-card").forEach((row, i) => {
    const lbl = row.querySelector(".subitem-label");
    if (lbl) lbl.textContent = `${i + 1})`;
  });
}

// ---------------------------------------------------------------------------
// Drag-and-drop (HTML5)
// ---------------------------------------------------------------------------
function initDragDrop(list) {
  if (!list) return;
  // Remove old listeners by cloning children (lightweight approach: use event delegation on list)
  list._dndInit = true;

  list.addEventListener("dragstart", e => {
    const card = e.target.closest(".sub-item-card, .subitem-card");
    if (!card) return;
    e.dataTransfer.effectAllowed = "move";
    list._dragging = card;
    setTimeout(() => card.style.opacity = "0.4", 0);
  }, { capture: false });

  list.addEventListener("dragend", e => {
    const card = e.target.closest(".sub-item-card, .subitem-card");
    if (card) card.style.opacity = "";
    list.querySelectorAll(".drag-over").forEach(el => el.classList.remove("drag-over"));
    list._dragging = null;
  });

  list.addEventListener("dragover", e => {
    e.preventDefault();
    const over = e.target.closest(".sub-item-card, .subitem-card");
    if (over && over !== list._dragging) {
      list.querySelectorAll(".drag-over").forEach(el => el.classList.remove("drag-over"));
      over.classList.add("drag-over");
    }
  });

  list.addEventListener("drop", e => {
    e.preventDefault();
    const over = e.target.closest(".sub-item-card, .subitem-card");
    if (over && list._dragging && over !== list._dragging) {
      over.classList.remove("drag-over");
      const rect = over.getBoundingClientRect();
      const midY = rect.top + rect.height / 2;
      if (e.clientY < midY) {
        list.insertBefore(list._dragging, over);
      } else {
        list.insertBefore(list._dragging, over.nextSibling);
      }
      // Re-index
      const sectionIdx = _getSectionIdxFromList(list);
      if (list.classList.contains("items-list")) {
        _reindexSubItems(sectionIdx);
      } else {
        _reindexSubSubItems(list);
      }
      markDirty();
    }
  });
}

// ---------------------------------------------------------------------------
// Rich text commands
// ---------------------------------------------------------------------------
function richCmd(btn, cmd) {
  const card = btn.closest(".sub-item-card");
  const editor = card.querySelector(".sub-item-editor");
  editor.focus();
  document.execCommand(cmd, false, null);
}

// ---------------------------------------------------------------------------
// Assumptions
// ---------------------------------------------------------------------------
function addAssumptionRow(text = "") {
  const list = document.getElementById("assumptions-list");
  const idx = list.querySelectorAll(".assumption-row").length;

  const row = document.createElement("div");
  row.className = "assumption-row";
  row.draggable = true;
  row.innerHTML = `
    <span class="drag-handle" title="Drag to reorder">&#8942;</span>
    <span class="assumption-label">${LETTERS[idx % 26]})</span>
    <textarea class="assumption-text" rows="2" spellcheck="true">${_escapeHtml(text)}</textarea>
    <button class="btn-sm btn-up-down" onclick="moveAssumption(this,-1)">&#9650;</button>
    <button class="btn-sm btn-up-down" onclick="moveAssumption(this,1)">&#9660;</button>
    <button class="btn-sm danger" onclick="deleteAssumption(this)" title="Delete">&#10005;</button>
  `;
  row.querySelector("textarea").addEventListener("input", markDirty);
  list.appendChild(row);
  initDragDrop(list);
  _reindexAssumptions();
  markDirty();
}

function deleteAssumption(btn) {
  const row = btn.closest(".assumption-row");
  row.remove();
  _reindexAssumptions();
  markDirty();
}

function moveAssumption(btn, dir) {
  const row = btn.closest(".assumption-row");
  const list = row.parentElement;
  if (dir === -1) {
    const prev = row.previousElementSibling;
    if (prev) list.insertBefore(row, prev);
  } else {
    const next = row.nextElementSibling;
    if (next) list.insertBefore(next, row);
  }
  _reindexAssumptions();
  markDirty();
}

function _reindexAssumptions() {
  const list = document.getElementById("assumptions-list");
  if (!list) return;
  list.querySelectorAll(".assumption-row").forEach((row, i) => {
    const lbl = row.querySelector(".assumption-label");
    if (lbl) lbl.textContent = `${LETTERS[i % 26]})`;
  });
}

// ---------------------------------------------------------------------------
// Fee fields
// ---------------------------------------------------------------------------
function fmtFeeBlur(inp) {
  const raw = inp.value.replace(/[^\d]/g, "");
  if (raw === "") { inp.value = ""; return; }
  const n = parseInt(raw, 10);
  inp.value = "$" + n.toLocaleString("en-US");
  updateTotal();
}

function updateTotal() {
  let total = 0;
  for (let i = 1; i <= 4; i++) {
    const inp = document.getElementById(`fee-${i}`);
    if (!inp) continue;
    const raw = inp.value.replace(/[^\d]/g, "");
    if (raw) total += parseInt(raw, 10);
  }
  const display = document.getElementById("fee-total");
  if (display) display.textContent = "$" + total.toLocaleString("en-US");
}

function parseFeeValue(id) {
  const inp = document.getElementById(id);
  if (!inp) return 0;
  const raw = inp.value.replace(/[^\d]/g, "");
  return raw ? parseInt(raw, 10) : 0;
}

// ---------------------------------------------------------------------------
// Cc toggle
// ---------------------------------------------------------------------------
function toggleCcField() {
  const chk = document.getElementById("include-cc");
  const wrap = document.getElementById("cc-text-wrap");
  if (chk && wrap) {
    wrap.style.display = chk.checked ? "" : "none";
  }
}

// ---------------------------------------------------------------------------
// Collect form data
// ---------------------------------------------------------------------------
function getFormData() {
  const sections = [0, 1, 2, 3].map(i => {
    const list = document.getElementById(`section-${i}-items`);
    const items = [];
    if (list) {
      list.querySelectorAll(".sub-item-card").forEach(card => {
        const editor = card.querySelector(".sub-item-editor");
        const html = editor ? editor.innerHTML : "";
        const subitems = [];
        card.querySelectorAll(".subitem-card").forEach(row => {
          const ed = row.querySelector(".subitem-editor");
          subitems.push({ html: ed ? ed.innerHTML : "" });
        });
        items.push({ html, subitems });
      });
    }
    return {
      heading: (document.getElementById(`section-heading-${i}`) || {}).value || "",
      items,
      deleted: _deletedSections.has(i),  // Fix 3
    };
  });

  const assumptions = [];
  document.querySelectorAll("#assumptions-list .assumption-row").forEach(row => {
    const ta = row.querySelector("textarea");
    if (ta) assumptions.push(ta.value.trim());
  });

  return {
    template_id: (document.getElementById("template-select") || {}).value || "",
    date: (document.getElementById("date") || {}).value || "",
    proposal_number: (document.getElementById("proposal-number") || {}).value.trim() || "",
    recipient_first_name: (document.getElementById("recipient-first") || {}).value.trim() || "",
    recipient_last_name: (document.getElementById("recipient-last") || {}).value.trim() || "",
    company_name: (document.getElementById("company-name") || {}).value.trim() || "",
    street_address: (document.getElementById("street-address") || {}).value.trim() || "",
    city_state_zip: (document.getElementById("city-state-zip") || {}).value.trim() || "",
    subject_line: (document.getElementById("subject-line") || {}).value.trim() || "",
    site_location: (document.getElementById("site-location") || {}).value.trim() || "",
    intro_paragraph: (document.getElementById("intro-paragraph") || {}).value.trim() || "",
    sections,
    assumptions,
    fees: {
      section1: parseFeeValue("fee-1"),
      section2: parseFeeValue("fee-2"),
      section3: parseFeeValue("fee-3"),
      section4: parseFeeValue("fee-4"),
    },
    closing_paragraph: (document.getElementById("closing-paragraph") || {}).value.trim() || "",
    include_cc: (document.getElementById("include-cc") || {}).checked !== false,
    cc_line: (document.getElementById("cc-line") || {}).value.trim() || "",
  };
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------
function validateForm() {
  let valid = true;
  const required = [
    { id: "date", label: "Date" },
    { id: "proposal-number", label: "Proposal Number" },
    { id: "recipient-first", label: "Recipient First Name" },
    { id: "recipient-last", label: "Recipient Last Name" },
    { id: "company-name", label: "Company Name" },
    { id: "street-address", label: "Street Address" },
    { id: "city-state-zip", label: "City, State, ZIP" },
    { id: "subject-line", label: "Subject Line" },
    { id: "site-location", label: "Site / Location" },
  ];

  // Clear previous errors
  document.querySelectorAll(".field-error").forEach(el => el.textContent = "");
  document.querySelectorAll("input.error, textarea.error").forEach(el => el.classList.remove("error"));

  required.forEach(({ id, label }) => {
    const el = document.getElementById(id);
    const err = document.getElementById(`${id}-error`);
    if (!el) return;
    if (!el.value.trim()) {
      valid = false;
      el.classList.add("error");
      if (err) err.textContent = `${label} is required.`;
    }
  });

  // Validate proposal number format
  const pnEl = document.getElementById("proposal-number");
  const pnErr = document.getElementById("proposal-number-error");
  if (pnEl && pnEl.value.trim()) {
    const pattern = /^\d{6}-rev\d+$/;
    if (!pattern.test(pnEl.value.trim())) {
      valid = false;
      pnEl.classList.add("error");
      if (pnErr) pnErr.textContent = "Format must be: 226001-rev0 (6 digits, hyphen, rev, number)";
    }
  }

  if (!valid) {
    showToast("Please fix the highlighted fields before generating.", "error");
    // Scroll to first error
    const first = document.querySelector(".error");
    if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  return valid;
}

// ---------------------------------------------------------------------------
// Generate .docx
// ---------------------------------------------------------------------------
async function downloadDocx() {
  if (!validateForm()) return;
  setStatus("Generating document...");
  try {
    const data = getFormData();
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || "Generation failed");
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${data.proposal_number || "proposal"}.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus("");
    showToast("Document downloaded successfully!", "success");
    _isDirty = false;
  } catch (e) {
    setStatus("");
    showToast("Error: " + e.message, "error");
  }
}

// ---------------------------------------------------------------------------
// Export to PDF
// ---------------------------------------------------------------------------
async function exportPdf() {
  if (!validateForm()) return;
  setStatus("Exporting to PDF (this may take a moment)...");
  try {
    const data = getFormData();
    const res = await fetch("/api/export-pdf", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || "PDF export failed");
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${data.proposal_number || "proposal"}.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus("");
    showToast("PDF exported successfully!", "success");
  } catch (e) {
    setStatus("");
    showToast(e.message, "warning");
  }
}

// ---------------------------------------------------------------------------
// Save draft
// ---------------------------------------------------------------------------
function showSaveDraftModal() {
  const modal = document.getElementById("modal-save-draft");
  const nameEl = document.getElementById("draft-save-name");
  // Pre-fill with proposal number or current draft name
  const pn = (document.getElementById("proposal-number") || {}).value.trim();
  nameEl.value = _currentDraftName || pn || "";
  modal.classList.add("active");
  nameEl.focus();
  nameEl.select();
}

function closeSaveDraftModal() {
  document.getElementById("modal-save-draft").classList.remove("active");
}

async function confirmSaveDraft() {
  const nameEl = document.getElementById("draft-save-name");
  const name = nameEl.value.trim();
  if (!name) { showToast("Please enter a draft name.", "warning"); return; }
  closeSaveDraftModal();

  const data = getFormData();
  data.draft_name = name;
  try {
    const res = await fetch("/api/save-draft", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });
    const json = await res.json();
    if (json.ok) {
      _currentDraftName = json.name;
      _isDirty = false;
      showToast(`Draft "${json.name}" saved.`, "success");
    }
  } catch (e) {
    showToast("Save failed: " + e.message, "error");
  }
}

// ---------------------------------------------------------------------------
// Load draft
// ---------------------------------------------------------------------------
async function showLoadDraftModal() {
  const modal = document.getElementById("modal-load-draft");
  const list = document.getElementById("draft-list");
  list.innerHTML = '<p class="text-muted">Loading...</p>';
  modal.classList.add("active");

  try {
    const res = await fetch("/api/list-drafts");
    const drafts = await res.json();
    list.innerHTML = "";
    if (drafts.length === 0) {
      list.innerHTML = '<p class="text-muted">No saved drafts found.</p>';
      return;
    }
    drafts.forEach(d => {
      const item = document.createElement("div");
      item.className = "draft-item";
      item.innerHTML = `
        <span class="draft-item-name">${_escapeHtml(d.name)}</span>
        <span class="draft-item-date">${d.modified}</span>
        <button class="draft-item-del" onclick="deleteDraft('${_escapeHtml(d.name)}', this)" title="Delete draft">&#10005;</button>
      `;
      item.addEventListener("click", e => {
        if (e.target.classList.contains("draft-item-del")) return;
        closeLoadDraftModal();
        loadDraft(d.name);
      });
      list.appendChild(item);
    });
  } catch (err) {
    list.innerHTML = '<p class="text-muted">Could not load drafts.</p>';
  }
}

function closeLoadDraftModal() {
  document.getElementById("modal-load-draft").classList.remove("active");
}

async function deleteDraft(name, btn) {
  if (!confirm(`Delete draft "${name}"?`)) return;
  await fetch(`/api/delete-draft/${encodeURIComponent(name)}`, { method: "DELETE" });
  const item = btn.closest(".draft-item");
  item.remove();
}

async function loadDraft(name) {
  try {
    const res = await fetch(`/api/load-draft/${encodeURIComponent(name)}`);
    const data = await res.json();
    if (data.error) { showToast(data.error, "error"); return; }
    populateForm(data);
    _currentDraftName = name;
    _isDirty = false;
    showToast(`Draft "${name}" loaded.`, "success");
  } catch (e) {
    showToast("Load failed: " + e.message, "error");
  }
}

// ---------------------------------------------------------------------------
// Populate form from saved data
// ---------------------------------------------------------------------------
function populateForm(data) {
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ""; };

  set("date", data.date);
  set("proposal-number", data.proposal_number);
  set("recipient-first", data.recipient_first_name);
  set("recipient-last", data.recipient_last_name);
  set("company-name", data.company_name);
  set("street-address", data.street_address);
  set("city-state-zip", data.city_state_zip);
  set("subject-line", data.subject_line);
  set("site-location", data.site_location);
  set("intro-paragraph", data.intro_paragraph);
  set("closing-paragraph", data.closing_paragraph);

  // Template selector
  const tplSel = document.getElementById("template-select");
  if (tplSel && data.template_id) tplSel.value = data.template_id;

  // Fix 3: reset deleted sections before restoring
  _deletedSections.clear();
  for (let i = 0; i < 4; i++) {
    const panel = document.getElementById(`section-panel-${i}`);
    if (panel) panel.classList.remove("section-deleted");
  }
  _updateRestoreArea();

  // Sections
  (data.sections || []).forEach((sec, i) => {
    const hEl = document.getElementById(`section-heading-${i}`);
    if (hEl) hEl.value = sec.heading || "";
    if (sec.deleted) deleteSection(i);  // Fix 3: re-apply deleted state

    const list = document.getElementById(`section-${i}-items`);
    if (list) {
      list.innerHTML = "";
      (sec.items || []).forEach((item, idx) => {
        const card = _makeSubItemCard(i, idx);
        const editor = card.querySelector(".sub-item-editor");
        if (editor) editor.innerHTML = item.html || "";
        const subList = card.querySelector(".subitems-list");
        (item.subitems || []).forEach((si, si_idx) => {
          const row = _makeSubSubItemRow(si_idx);
          const ed = row.querySelector(".subitem-editor");
          if (ed) ed.innerHTML = si.html || "";
          subList.appendChild(row);
        });
        list.appendChild(card);
      });
      initDragDrop(list);
    }
  });

  // Assumptions
  const assumList = document.getElementById("assumptions-list");
  if (assumList) {
    assumList.innerHTML = "";
    (data.assumptions || []).forEach(a => addAssumptionRow(a));
  }

  // Fees
  for (let i = 1; i <= 4; i++) {
    const inp = document.getElementById(`fee-${i}`);
    const val = data.fees ? data.fees[`section${i}`] : 0;
    if (inp) inp.value = val ? "$" + Number(val).toLocaleString("en-US") : "";
  }
  updateTotal();

  // Cc
  const ccChk = document.getElementById("include-cc");
  if (ccChk) ccChk.checked = data.include_cc !== false;
  set("cc-line", data.cc_line);
  toggleCcField();
}

function _makeSubSubItemRow(idx) {
  const row = document.createElement("div");
  row.className = "subitem-card";
  row.draggable = true;
  row.innerHTML = `
    <span class="drag-handle">&#8942;</span>
    <span class="subitem-label">${idx + 1})</span>
    <div class="subitem-editor" contenteditable="true" data-placeholder="Sub-item..." spellcheck="true"></div>
    <button class="btn-sm btn-up-down" onclick="moveSubSubItem(this,-1)">&#9650;</button>
    <button class="btn-sm btn-up-down" onclick="moveSubSubItem(this,1)">&#9660;</button>
    <button class="btn-sm danger" onclick="deleteSubSubItem(this)">&#10005;</button>
  `;
  row.querySelector(".subitem-editor").addEventListener("input", markDirty);
  return row;
}

// ---------------------------------------------------------------------------
// New proposal
// ---------------------------------------------------------------------------
function newProposal() {
  if (_isDirty) {
    const modal = document.getElementById("modal-new-proposal");
    modal.classList.add("active");
  } else {
    _doNewProposal();
  }
}

function closeNewProposalModal() {
  document.getElementById("modal-new-proposal").classList.remove("active");
}

function confirmNewProposal() {
  closeNewProposalModal();
  _doNewProposal();
}

function _doNewProposal() {
  // Reset all fields
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ""; };
  set("date", new Date().toISOString().slice(0, 10));
  set("proposal-number", "");
  set("recipient-first", "");
  set("recipient-last", "");
  set("company-name", "");
  set("street-address", "");
  set("city-state-zip", "");
  set("subject-line", "");
  set("site-location", "");

  // Clear sections and restore any deleted ones (Fix 3)
  _deletedSections.clear();
  for (let i = 0; i < 4; i++) {
    const list = document.getElementById(`section-${i}-items`);
    if (list) list.innerHTML = "";
    const panel = document.getElementById(`section-panel-${i}`);
    if (panel) panel.classList.remove("section-deleted");
  }
  _updateRestoreArea();

  // Reset headings from template
  if (_templateDefaults.section_headings) {
    _templateDefaults.section_headings.forEach((h, i) => {
      const el = document.getElementById(`section-heading-${i}`);
      if (el) el.value = h;
    });
  }

  // Reset assumptions
  const assumList = document.getElementById("assumptions-list");
  if (assumList) {
    assumList.innerHTML = "";
    (_templateDefaults.default_assumptions || []).forEach(a => addAssumptionRow(a));
  }

  // Reset fees
  for (let i = 1; i <= 4; i++) {
    const inp = document.getElementById(`fee-${i}`);
    if (inp) inp.value = "";
  }
  updateTotal();

  // Reset intro/closing
  set("intro-paragraph", _templateDefaults.default_intro || "");
  set("closing-paragraph", _templateDefaults.default_closing || "");

  // Reset cc
  const ccChk = document.getElementById("include-cc");
  if (ccChk) ccChk.checked = true;
  toggleCcField();

  // Load default cc from settings
  fetch("/api/settings").then(r => r.json()).then(settings => {
    const defCc = (settings.user || {}).default_cc || "";
    set("cc-line", defCc);
  }).catch(() => {});

  _currentDraftName = null;
  _isDirty = false;
  setStatus("");
  window.scrollTo({ top: 0, behavior: "smooth" });
  showToast("New proposal started.", "success");
}

// ---------------------------------------------------------------------------
// About modal
// ---------------------------------------------------------------------------
function showAbout() {
  document.getElementById("modal-about").classList.add("active");
}
function closeAbout() {
  document.getElementById("modal-about").classList.remove("active");
}

// Close modals on overlay click
document.addEventListener("click", e => {
  if (e.target.classList.contains("modal-overlay")) {
    e.target.classList.remove("active");
  }
});

// ---------------------------------------------------------------------------
// Dirty flag
// ---------------------------------------------------------------------------
function markDirty() {
  _isDirty = true;
}

window.addEventListener("beforeunload", e => {
  if (_isDirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});

// ---------------------------------------------------------------------------
// Status bar
// ---------------------------------------------------------------------------
function setStatus(msg) {
  const el = document.getElementById("toolbar-status");
  if (el) el.textContent = msg;
}

// ---------------------------------------------------------------------------
// Toast notifications
// ---------------------------------------------------------------------------
function showToast(msg, type = "info") {
  const container = document.getElementById("toast-container");
  if (!container) return;
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------
function _escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
