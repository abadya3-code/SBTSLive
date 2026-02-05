/****************************************************
 * SBTS - Smart Blind Tracking System (Front-end only)
 * LocalStorage (no backend)
 ****************************************************/

/* ==========================
   CLEANUP3: UTILS & SAFE STORAGE
   - Centralized query helpers
   - Centralized LocalStorage helpers (string + JSON)
========================== */
const SBTS_UTILS = (() => {
  const qs = (sel, root = document) => root.querySelector(sel);
  const qsa = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const byId = (id) => document.getElementById(id);

  const safeJsonParse = (str, fallback = null) => {
    try { return JSON.parse(str); } catch (e) { return fallback; }
  };
  const safeJsonStringify = (obj, fallback = "null") => {
    try { return JSON.stringify(obj); } catch (e) { return fallback; }
  };

  const LS = {
    get: (key, fallback = null) => {
      const v = localStorage.getItem(key);
      return v === null ? fallback : v;
    },
    set: (key, value) => {
      localStorage.setItem(key, String(value));
    },
    remove: (key) => localStorage.removeItem(key),
    getJSON: (key, fallback = null) => safeJsonParse(localStorage.getItem(key), fallback),
    setJSON: (key, obj) => localStorage.setItem(key, safeJsonStringify(obj, "null")),
  };

  return { qs, qsa, byId, safeJsonParse, safeJsonStringify, LS };
})();
/* ==========================
   CLEANUP4: STATE ORGANIZATION
   - Keep current state structure (non-breaking)
   - Mirror theme/UI prefs into state.ui.* for future development
   - Add file-section markers for easier maintenance
========================== */


/* ==========================
   WORKFLOW DEFINITIONS
========================== */

// ==========================
// SUPABASE LIVE (Pilot)
// ==========================
// (Deduped) Supabase LIVE helpers are defined later in this file.


async function sbtsLiveLoadMyProjects(){
  const sb = sbtsSupabaseClient();
  if(!sb) return [];
  const { data, error } = await sb.from("projects").select("id,name,description,created_at").order("created_at", { ascending: false });
  if(error) throw error;
  return data || [];
}

function sbtsApplyLiveUser(profile, user){
  // Create/merge a SBTS local user object mapped to Supabase user id
  const uid = user.id;
  let local = state.users.find(u => u.id === uid);
  if(!local){
    local = {
      id: uid,
      fullName: profile.full_name || user.email || "User",
      username: (user.email || "user").split("@")[0],
      password: "",
      phone: "",
      email: user.email || "",
      role: (profile.role || "user").toLowerCase(),
      status: "active",
      jobTitle: profile.job_title || "",
      profileImage: profile.avatar_url || null
    };
    state.users.push(local);
  } else {
    local.fullName = profile.full_name || local.fullName;
    local.email = user.email || local.email;
    local.role = (profile.role || local.role || "user").toLowerCase();
    local.jobTitle = profile.job_title || local.jobTitle;
    local.profileImage = profile.avatar_url || local.profileImage;
  }

  // Minimal permissions mapping for pilot (can be expanded later)
  state.permissions = state.permissions || {};
  if(!state.permissions[uid]) state.permissions[uid] = {};
  const isAdmin = local.role === "admin";
  // baseline
  const base = {
    manageAreas: isAdmin,
    manageProjects: isAdmin,
    manageBlinds: isAdmin,
    changePhases: true,
    viewReports: isAdmin,
    manageReportsCards: isAdmin,
    manageCertificateSettings: isAdmin,
    manageWorkflowControl: isAdmin,
    editBranding: isAdmin,
    manageTrainingVisibility: isAdmin,
    manageRolesCatalog: isAdmin,
    manageUsers: isAdmin,
    manageRequests: isAdmin,
    manageFinalApprovals: isAdmin,
    managePhaseOwnership: isAdmin
  };
  state.permissions[uid] = Object.assign(base, state.permissions[uid]);

  state.currentUser = local;
}

async function sbtsHydrateLiveDataAfterLogin(user){
  const profile = await sbtsLiveEnsureProfile(user);
  sbtsApplyLiveUser(profile, user);
  try{
    const projects = await sbtsLiveLoadMyProjects();
    // Map remote projects to SBTS local shape (preserve additional local fields if any)
    state.projects = (projects || []).map(p => ({
      id: p.id,
      name: p.name,
      desc: p.description || "",
      description: p.description || "",
      createdAt: p.created_at || new Date().toISOString()
    }));
  }catch(e){
    console.warn("Live projects load failed", e);
    toast("Live projects load failed - using local");
  }
  saveState();
}

const PHASES = ["broken", "assembly", "tightTorque", "finalTight", "inspectionReady"];
const phaseLabels = {
  broken: "Broken / Preparation",
  assembly: "Assembly",
  tightTorque: "Tight & Torque",
  finalTight: "Final Tight",
  inspectionReady: "Inspection & Ready",
};
const WORKFLOW_REQUIRED_ROLE = {
  broken: "coordinator",
  assembly: "technician",
  tightTorque: "safety",
  finalTight: "qc",
  inspectionReady: "inspection",
};
const FINAL_APPROVALS = [
  { key: "inspection", label: "Inspection approval", role: "inspection" },
  { key: "ti_engineer", label: "T&I Engineer approval", role: "ti_engineer" },
  { key: "operation_foreman", label: "Operation Foreman approval", role: "operation_foreman" },
];
const EXTRA_SLIP_APPROVAL = { key: "metal_foreman_demolish", label: "Metal Foreman approval (Demolish)", role: "metal_foreman" };
const ROLE_LABELS = {
  admin: "Admin",
  coordinator: "Coordinator",
  technician: "Technician",
  safety: "Safety",
  qc: "QC",
  inspection: "Inspection",
  ti_engineer: "T&I Engineer",
  operation_foreman: "Operation Foreman",
  metal_foreman: "Metal Foreman",
};


// Roles & Specialties Catalog (Admin managed)
const ROLES_CATALOG_KEY = "sbts_roles_catalog_v1";

function defaultRolesCatalog() {
  return [
    { id: "admin", label: "Admin", type: "role", active: true, locked: true },
    { id: "supervisor", label: "Supervisor", type: "role", active: true, locked: false },
    { id: "technician", label: "Technician", type: "role", active: true, locked: false },
    { id: "qaqc", label: "QA/QC", type: "role", active: true, locked: false },
    { id: "safety", label: "Safety", type: "role", active: true, locked: false },
    // Specialties (optional – can be used for filtering/labels later)
    { id: "inspection", label: "Inspection", type: "specialty", active: true, locked: false },
  ];
}

function slugifyId(label) {
  return String(label || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^a-z0-9_\/\-]+/g, "")
    .replace(/[\/\-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeRoleId(input) {
  const raw = String(input ?? "").trim();
  if (!raw) return "";
  const lc = raw.toLowerCase().trim();
  // common aliases
  const alias = {
    "qc": "qaqc",
    "qaqc": "qaqc",
    "qa/qc": "qaqc",
    "qa-qc": "qaqc",
    "qa qc": "qaqc",
    "quality": "qaqc",
    "qualitycontrol": "qaqc",
    "quality_control": "qaqc",
    "inspector": "inspection",
  };
  if (alias[lc]) return alias[lc];
  return slugifyId(lc);
}

function loadRolesCatalog() {
  try {
    const raw = SBTS_UTILS.LS.get(ROLES_CATALOG_KEY);
    const base = defaultRolesCatalog();
    if (!raw) return base;

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return base;

    // merge base with parsed, keep base locked ones
    const map = new Map();
    base.forEach((r) => map.set(r.id, r));
    parsed.forEach((r) => {
      if (!r || !r.id) return;
      const id = normalizeRoleId(r.id);
      const existing = map.get(id);
      const merged = {
        id,
        label: String(r.label || existing?.label || id),
        type: r.type === "specialty" ? "specialty" : "role",
        active: typeof r.active === "boolean" ? r.active : (existing?.active ?? true),
        locked: existing?.locked ?? false,
      };
      map.set(id, merged);
    });
    return Array.from(map.values());
  } catch (e) {
    console.warn("loadRolesCatalog failed:", e);
    return defaultRolesCatalog();
  }
}

function saveRolesCatalog() {
  SBTS_UTILS.LS.set(ROLES_CATALOG_KEY, JSON.stringify(state.rolesCatalog || []));
}

function getRolesCatalog(type) {
  const list = Array.isArray(state.rolesCatalog) ? state.rolesCatalog : [];
  if (!type) return list;
  return list.filter((r) => r.type === type);
}

function getActiveRoleOptions(type = "role") {
  return getRolesCatalog(type).filter((r) => r.active);
}

function getRoleLabelById(idOrLabel) {
  const id = normalizeRoleId(idOrLabel);
  const hit = (state.rolesCatalog || []).find((r) => r.id === id);
  return hit?.label || ROLE_LABELS[id] || idOrLabel || "-";
}


function populateRoleSelects() {
  // Register role dropdown
  const regRole = document.getElementById("reg_role");
  if (regRole) {
    regRole.innerHTML = "";
    getActiveRoleOptions("role").forEach((r) => {
      const opt = document.createElement("option");
      opt.value = r.id;
      opt.textContent = r.label;
      regRole.appendChild(opt);
    });
    // Keep admin first if exists
    const adminOpt = Array.from(regRole.options).find(o => o.value === "admin");
    if (adminOpt) regRole.insertBefore(adminOpt, regRole.firstChild);
  }
}

function renderRolesCatalogManager() {
  const tbody = document.getElementById("rolesCatalogTableBody");
  if (!tbody) return;

  const q = (document.getElementById("roleCatalogSearch")?.value || "").trim().toLowerCase();
  const filter = (document.getElementById("roleCatalogFilter")?.value || "all");
  const sort = (document.getElementById("roleCatalogSort")?.value || "type");

  let rows = (state.rolesCatalog || []).slice();

  if (filter === "active") rows = rows.filter(r => !!r.active);
  if (filter === "hidden") rows = rows.filter(r => !r.active);

  if (q) rows = rows.filter(r => (r.label || "").toLowerCase().includes(q) || (r.id || "").toLowerCase().includes(q));

  rows.sort((a, b) => {
    if (sort === "az") return (a.label || "").localeCompare(b.label || "");
    if (sort === "za") return (b.label || "").localeCompare(a.label || "");
    if ((a.type || "") !== (b.type || "")) return (a.type || "").localeCompare(b.type || "");
    return (a.label || "").localeCompare(b.label || "");
  });

  tbody.innerHTML = "";

  if (rows.length === 0) {
    const tr = document.createElement("tr");
    try{ tr.dataset.blindId = b.id; tr.id = "slip_row_" + b.id; }catch(e){}
    tr.innerHTML = `<td colspan="5" class="muted">No items found.</td>`;
    tbody.appendChild(tr);
    return;
  }

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");

    const typeBadge = r.type === "role" ? "badge-blue" : "badge-gray";
    const statusHtml = r.active
      ? '<span class="badge badge-green">active</span>'
      : '<span class="badge badge-red">hidden</span>';

    const lockedHtml = r.locked ? ' <span class="badge badge-gold">locked</span>' : "";

    const deleteBtn = r.locked
      ? ""
      : `<button class="secondary-btn tiny danger" onclick="deprecateCatalogRole('${escapeJs(r.id)}')">Delete</button>`;

    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>
        <div class="role-name">
          <div><b>${escapeHtml(r.label)}</b>${lockedHtml}</div>
          <div class="role-meta">
            <span class="badge ${typeBadge}">${escapeHtml(r.type)}</span>
            <span class="role-id">${escapeHtml(r.id)}</span>
          </div>
        </div>
      </td>
      <td>
        <span class="badge ${typeBadge}">${r.type === "role" ? "Role" : "Specialty"}</span>
      </td>
      <td>${statusHtml}</td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn tiny" onclick="renameCatalogRole('${escapeJs(r.id)}')">Rename</button>
          <button class="secondary-btn tiny" onclick="toggleCatalogRole('${escapeJs(r.id)}')">${r.active ? "Hide" : "Unhide"}</button>
          ${deleteBtn}
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}


function bindRolesCatalogControls() {
  const btnToggle = document.getElementById("btnToggleRoleAdd");
  const form = document.getElementById("roleAddForm");
  const search = document.getElementById("roleCatalogSearch");
  const filter = document.getElementById("roleCatalogFilter");
  const sort = document.getElementById("roleCatalogSort");

  if (btnToggle && form && !btnToggle._bound) {
    btnToggle._bound = true;
    btnToggle.onclick = () => {
      const open = form.style.display !== "none";
      form.style.display = open ? "none" : "block";
      btnToggle.textContent = open ? "+ Add new" : "− Close";
    };
  }

  const reRender = () => renderRolesCatalogManager();
  [search, filter, sort].forEach(el => {
    if (el && !el._bound) {
      el._bound = true;
      el.addEventListener("input", reRender);
      el.addEventListener("change", reRender);
    }
  });
}

function addCatalogRole() {
  const labelEl = document.getElementById("roleAddLabel");
  const typeEl  = document.getElementById("roleAddType");
  if (!labelEl || !typeEl) return;

  const label = labelEl.value.trim();
  const type = typeEl.value === "specialty" ? "specialty" : "role";
  if (!label) return alert("Enter a name.");

  let id = slugifyId(label);
  if (!id) return alert("Invalid name.");
  // ensure uniqueness
  const exists = (x) => (state.rolesCatalog || []).some(r => r.id === x);
  if (exists(id)) {
    let n = 2;
    while (exists(`${id}_${n}`)) n++;
    id = `${id}_${n}`;
  }

  state.rolesCatalog = state.rolesCatalog || [];
  state.rolesCatalog.push({ id, label, type, active: true, locked: false });
  saveRolesCatalog();
  populateRoleSelects();
  renderRolesCatalogManager();
  toast("Role added ✅");
  labelEl.value = "";
}

function renameCatalogRole(id) {
  const r = (state.rolesCatalog || []).find(x => x.id === id);
  if (!r) return;
  const next = prompt("New name:", r.label);
  if (next === null) return;
  const label = next.trim();
  if (!label) return alert("Invalid name.");
  r.label = label;
  saveRolesCatalog();
  populateRoleSelects();
  renderRolesCatalogManager();
  // rerender current visible pages if needed
  if (state.currentPage === "permissionsPage") renderUsers();
  toast("Renamed ✅");
}

function toggleCatalogRole(id) {
  const r = (state.rolesCatalog || []).find(x => x.id === id);
  if (!r) return;
  if (r.locked) return alert("This item is locked.");
  r.active = !r.active;
  saveRolesCatalog();
  populateRoleSelects();
  renderRolesCatalogManager();
  if (state.currentPage === "permissionsPage") renderUsers();
  toast(r.active ? "Unhidden ✅" : "Hidden ✅");
}

function deprecateCatalogRole(id) {
  const r = (state.rolesCatalog || []).find(x => x.id === id);
  if (!r) return;
  if (r.locked) return alert("This item is locked.");
  if (!confirm("Delete (hide) this role/specialty? Old records will keep it in history.")) return;
  r.active = false; // soft delete
  saveRolesCatalog();
  populateRoleSelects();
  renderRolesCatalogManager();
  if (state.currentPage === "permissionsPage") renderUsers();
  toast("Deleted (hidden) ✅");
}



/* ==========================
   THEMES PRESETS
========================== */
const THEME_PRESETS = {
  classic: {
    primary: "#174c7e",
    primaryStrong: "#0f3154",
    bodyBg: "linear-gradient(to bottom, #f0f4ff, #ffffff)",
    tableHead: "#f3f4ff",
  },
  midnight: {
    primary: "#111827",
    primaryStrong: "#0b1220",
    bodyBg: "linear-gradient(to bottom, #eef2ff, #ffffff)",
    tableHead: "#eef2ff",
  },
  slate: {
    primary: "#334155",
    primaryStrong: "#1f2937",
    bodyBg: "linear-gradient(to bottom, #f1f5f9, #ffffff)",
    tableHead: "#f1f5f9",
  },
  emerald: {
    primary: "#065f46",
    primaryStrong: "#064e3b",
    bodyBg: "linear-gradient(to bottom, #ecfeff, #ffffff)",
    tableHead: "#e6fffb",
  },
};

/* ==========================
   STATE
========================== */
const state = {
  notifSystemVersion: "47.15",
  users: [],
  currentUser: null,
  areas: [],
  projects: [],
  blinds: [],
  permissions: {},
  rolesCatalog: [],
  notifications: { byUser: {} },

  // Request inbox (Phase owner change / Assistance)
  requests: [],

  themePreset: "classic",
  themeColor: "#174c7e",
  fontSize: 14,

  currentProjectId: null,
  currentBlindId: null,
  currentProjectAreaFilter: null, // when navigating from Area -> Projects

  branding: {
    programTitle: "Smart Blind Tag System",
    programSubtitle: "Smart Blind Tag System",
    siteTitle: "Shedgum Gas Plant",
    siteSubtitle: "Smart Blind Tag System",
    companyName: "Aramco",
    companySub: "Saudi Aramco",
    companyLogo: null, // dataURL
  },

  certificate: {
    templates: [
      { id: "default", name: "Default" },
      { id: "clean", name: "Clean" },
    ],
    activeTemplate: "default",
    title: "Smart Blind Tag System Certificate",
    headerBg: "#ffffff",
    statusStyle: "big", // "pill" or "big"
    showWorkflow: true,
    showApprovals: true,
    footerText: "This is a digital certificate (no handwritten signature required).",
  },

  // Tag printing theme (single global color)
  tagTheme: {
    color: "#0F6D8C",
    audit: []
  },

  // Slip blind UI state
  slip: { areaId: "", projectId: "", selectedIds: [] },

  ui: {
    showTrainingPage: true,
    workflowSandboxEnforcement: true,
    enforcePhaseOrder: true,
    notifDrawer: {
      tab: "action",        // all | action | unread | done
      projectFilter: "auto", // auto | current | all
      groupBy: "project",     // project | type (future)
      hideInfo: false,       // if true: hide non-action notifications in drawer
    },
  },

  reports: {
    // order of cards by key
    cardOrder: [],
    // visibility map by key
    cardVisibility: {},
    // active card filter key (null means no card filter)
    activeCardKey: null,
    // preset name
    preset: "management",
  },
};

// CLEANUP4: Mirror theme prefs into state.ui (backward compatible)
state.ui = state.ui || {};
if (state.ui.themePreset == null) state.ui.themePreset = state.themePreset;
if (state.ui.themeColor == null) state.ui.themeColor = state.themeColor;
if (state.ui.fontSize == null) state.ui.fontSize = state.fontSize;


// Runtime-only (not saved)
const runtime = { afterLoginBlindId: null };


// Hybrid permissions: keep roles minimal (User/Supervisor/Admin), and
// control access via per-user permission toggles.
const PERMISSION_DEFS = [
  { key: "manageAreas", label: "Manage areas" },
  { key: "manageProjects", label: "Manage projects" },
  { key: "managePhaseOwnership", label: "Manage phase ownership" },
  { key: "manageRequests", label: "Approve & manage requests" },
  { key: "manageBlinds", label: "Manage blinds" },
  { key: "changePhases", label: "Change blind phases" },
  { key: "viewReports", label: "View reports" },
  { key: "manageReportsCards", label: "Manage reports cards" },
  { key: "manageCertificateSettings", label: "Manage certificate settings" },
  { key: "manageTagSettings", label: "Manage tag settings" },
  { key: "manageWorkflowControl", label: "Workflow control" },
  { key: "editBranding", label: "Edit app/project branding" },
  { key: "manageTrainingVisibility", label: "Show/Hide training page" },
  { key: "manageRolesCatalog", label: "Roles & specialties manager" },
  { key: "manageFinalApprovals", label: "Final approvals manager" },
  { key: "manageUsers", label: "Manage users & permissions" },
];

let currentPermissionsUserId = null;

// Phase confirm modal payload (not persisted)
let phaseConfirmPayload = null;


/* ==========================
   STORAGE
========================== */
/* ==========================
   PERSISTENCE (Local)
========================== */
function saveState() {
  SBTS_UTILS.LS.set("sbts_state", JSON.stringify(state));
  updateNotificationsBadge();
}

/* ==========================
   BACKUP & RESTORE (Admin only)
   Local now + Hooks ready for Supabase later (disabled)
========================== */
const SBTS_BACKUP_KEYS_V1 = [
  "sbts_state",
  "sbts_roles_catalog_v1",
  "sbts_workflow_config_v1",
  "sbts_reports",
  "sbts_last_page",
  "sbts_nav_stack_v1",
  "sbts_finalApprovalsDirectory_v1",
  "sbts_finalApprovalsSafeMode_v1",
  "sbts_finalApprovalsRedirects_v1",
  "sbts_live_mode",
];

const SBTS_BACKUP_CFG = {
  autoEnabledKey: "sbts_backup_auto_enabled_v1",
  autoLastTsKey: "sbts_backup_auto_last_ts_v1",
  autoStoreKey: "sbts_backup_auto_store_v1",
  keepLast: 10,
};

function sbtsIsAdmin() {
  return state.currentUser?.role === "admin";
}

const SBTS_REAL_USER_KEY = "sbts_real_user_id_v1";
function sbtsEnsureRealUser(){
  try{
    const cur = state.currentUser?.id;
    const saved = localStorage.getItem(SBTS_REAL_USER_KEY);
    if((!saved || saved === "") && cur){
      localStorage.setItem(SBTS_REAL_USER_KEY, cur);
    }
  }catch(e){}
}

function sbtsIsAdminLike(){
  // Demo/admin-only controls should remain available even when "acting as" another role.
  const realId = (()=>{ try{ return localStorage.getItem(SBTS_REAL_USER_KEY) || ""; }catch(e){ return ""; } })();
  const realUser = (Array.isArray(state.users)?state.users:[]).find(u=>u && u.id===realId);
  const r = normalizeRoleId((realUser?.role) || (state.currentUser?.role) || "");
  return r === "admin" || r === "system_admin" || r === "systemadmin";
}

function sbtsBackupSafeParse(raw) {
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function sbtsCollectBackupData() {
  const data = {};
  SBTS_BACKUP_KEYS_V1.forEach((k) => {
    const raw = SBTS_UTILS.LS.get(k);
    data[k] = raw ? sbtsBackupSafeParse(raw) : null;
  });
  return data;
}

function sbtsApplyBackupData(dataObj) {
  SBTS_BACKUP_KEYS_V1.forEach((k) => {
    if (!Object.prototype.hasOwnProperty.call(dataObj, k)) return;
    const val = dataObj[k];
    if (val === null || typeof val === "undefined") localStorage.removeItem(k);
    else SBTS_UTILS.LS.set(k, JSON.stringify(val));
  });
}

function sbtsCreateBackupPayload(createdBy, extraMeta) {
  const baseMeta = {
    app: "SBTS",
    version: window.SBTS_VERSION || "v1.0_patch34",
    createdAt: new Date().toISOString(),
    createdBy: createdBy || (state.currentUser?.fullName || "Admin"),
    mode: (window.SBTS_LIVE_MODE ? "live" : "local"),
  };
  const meta = Object.assign({}, baseMeta, (extraMeta || {}));
  return { meta, data: sbtsCollectBackupData() };
}


function sbtsDownloadJson(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}


function sbtsEscapeHtml(str) {
  return String(str || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function sbtsPromptBackupMeta() {
  let label = prompt("Backup name/label (optional). مثال: قبل تحديث الرولات", "");
  if (label === null) return null;
  label = String(label || "").trim();

  let note = prompt("Notes (optional). مثال: نسخة مستقرة قبل تعديل صفحة Dashboard", "");
  if (note === null) return null;
  note = String(note || "").trim();

  const sanitize = (s) => s.replace(/[^a-zA-Z0-9_\-\u0600-\u06FF ]/g, "").trim();
  return { label, note, labelSafe: sanitize(label).replace(/\s+/g, "_") };
}

function sbtsLocalBackupsWrite(arr) {
  sbtsLocalBackupsWrite(arr);
}


function sbtsCreateBackupNow() {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const meta = sbtsPromptBackupMeta();
  if (meta === null) return;
  const payload = sbtsCreateBackupPayload(null, { label: meta.label, note: meta.note });
  try { sbtsSaveLocalBackup(payload); } catch (e) {}
  const date = payload.meta.createdAt.slice(0, 10);
  const labelPart = meta.labelSafe ? `_${meta.labelSafe}` : "";
  sbtsDownloadJson(`SBTS_backup_${date}${labelPart}.json`, payload);
  toast("Backup created (downloaded + saved locally).");
}


function sbtsSaveBackupLocalNow() {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const meta = sbtsPromptBackupMeta();
  if (meta === null) return;
  const payload = sbtsCreateBackupPayload(null, { label: meta.label, note: meta.note });
  sbtsSaveLocalBackup(payload);
  toast("Backup saved locally.");
  try { sbtsOpenLocalBackups(); } catch (e) {}
}



async function sbtsRestoreBackupFromFile(event) {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const file = event?.target?.files?.[0];
  if (!file) return;

  const ok1 = confirm("⚠️ Restore will overwrite SBTS data in this browser. Continue?");
  if (!ok1) {
    event.target.value = "";
    return;
  }
  const ok2 = confirm("Last confirmation: restore now?");
  if (!ok2) {
    event.target.value = "";
    return;
  }

  sbtsAutoBackupBefore("Restore (File Backup)");

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const payload = JSON.parse(String(reader.result));
      if (!payload?.data) throw new Error("Invalid backup file");
      sbtsApplyBackupData(payload.data);
      alert("✅ Restore completed. The app will reload now.");
      location.reload();
    } catch (e) {
      alert("❌ Restore failed: " + (e.message || e));
    } finally {
      event.target.value = "";
    }
  };
  reader.readAsText(file);
}

function sbtsSetWeeklyBackupEnabled(enabled) {
  SBTS_UTILS.LS.set(SBTS_BACKUP_CFG.autoEnabledKey, enabled ? "1" : "0");
}

function sbtsIsWeeklyBackupEnabled() {
  return SBTS_UTILS.LS.get(SBTS_BACKUP_CFG.autoEnabledKey) === "1";
}

function sbtsGetLocalBackups() {
  const raw = SBTS_UTILS.LS.get(SBTS_BACKUP_CFG.autoStoreKey);
  const arr = raw ? sbtsBackupSafeParse(raw) : [];
  return Array.isArray(arr) ? arr : [];
}

function sbtsSaveLocalBackup(payload) {
  const arr = sbtsGetLocalBackups();
  arr.unshift(payload);
  SBTS_UTILS.LS.set(
    SBTS_BACKUP_CFG.autoStoreKey,
    JSON.stringify(arr.slice(0, SBTS_BACKUP_CFG.keepLast))
  );
}




function sbtsAutoBackupBefore(actionTitle) {
  // Auto safety snapshot (local only) before high-impact changes
  if (!sbtsIsAdmin()) return;
  try {
    const ts = new Date();
    const stamp = ts.toISOString().replace("T", " ").slice(0, 19);
    const payload = sbtsCreateBackupPayload(null, {
      label: `AUTO • قبل ${actionTitle} • ${stamp}`,
      note: `Auto backup before: ${actionTitle}`,
      auto: true,
      autoType: "pre_action",
    });
    sbtsSaveLocalBackup(payload);
  } catch (e) {}
}
function sbtsRunWeeklyAutoBackup() {
  if (!sbtsIsWeeklyBackupEnabled()) return;
  const lastTs = parseInt(SBTS_UTILS.LS.get(SBTS_BACKUP_CFG.autoLastTsKey) || "0", 10);
  const now = Date.now();
  const every7d = 7 * 24 * 60 * 60 * 1000;
  if (!lastTs || now - lastTs >= every7d) {
    const payload = sbtsCreateBackupPayload();
    sbtsSaveLocalBackup(payload);
    SBTS_UTILS.LS.set(SBTS_BACKUP_CFG.autoLastTsKey, String(now));
  }
}

function sbtsToggleWeeklyBackup(enabled) {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  sbtsSetWeeklyBackupEnabled(enabled);
  toast(enabled ? "Weekly backup enabled." : "Weekly backup disabled.");
}

function sbtsOpenLocalBackups() {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const box = document.getElementById("sbtsLocalBackupsBox");
  if (!box) return;
  const arr = sbtsGetLocalBackups();
  if (!arr.length) {
    box.innerHTML = "<div style='opacity:.85;'>No local backups saved yet.</div>";
    return;
  }
  const fmt = (iso) => {
    try {
      return new Date(iso).toLocaleString();
    } catch (e) {
      return iso;
    }
  };
  box.innerHTML = arr
    .map((b, i) => {
      const when = fmt(b?.meta?.createdAt || "");
      const ver = b?.meta?.version || "";
      const label = (b?.meta?.label || "").trim();
      const note = (b?.meta?.note || "").trim();
      const labelLine = label ? `<div style="margin-top:2px;"><b>${sbtsEscapeHtml(label)}</b></div>` : "";
      const noteLine = note ? `<div style="opacity:.75; font-size:12px; margin-top:2px;">${sbtsEscapeHtml(note)}</div>` : "";
      return `
        <div style="display:flex; justify-content:space-between; gap:10px; padding:10px; border:1px solid rgba(255,255,255,.08); border-radius:12px; margin-top:8px; background:rgba(0,0,0,.06);">
          <div style="min-width:0;">
            <div><b>Local Backup #${i + 1}</b></div>
            ${labelLine}
            ${noteLine}
            <div style="opacity:.75; font-size:12px; margin-top:4px;">${when} • ${ver}</div>
          </div>
          <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; justify-content:flex-end;">
            <button class="btn-neutral" onclick="sbtsRestoreLocalBackup(${i})">Restore</button>
            <button class="btn-neutral" onclick="sbtsDownloadLocalBackup(${i})">Download</button>
            <button class="btn-danger" onclick="sbtsDeleteLocalBackup(${i})">Delete</button>
          </div>
        </div>
      `;
    })
    .join("");
}

async function sbtsRestoreLocalBackup(index) {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const arr = sbtsGetLocalBackups();
  const payload = arr[index];
  if (!payload) return;
  const ok1 = confirm("⚠️ Restore will overwrite SBTS data in this browser. Continue?");
  if (!ok1) return;
  const ok2 = confirm("Last confirmation: restore now?");
  if (!ok2) return;
  sbtsAutoBackupBefore("Restore (Local Backup)");
  try {
    sbtsApplyBackupData(payload.data || {});
    alert("✅ Restored. The app will reload now.");
    location.reload();
  } catch (e) {
    alert("❌ Restore failed: " + (e.message || e));
  }
}


function sbtsDownloadLocalBackup(index) {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const arr = sbtsGetLocalBackups();
  const payload = arr[index];
  if (!payload) return;
  const date = (payload?.meta?.createdAt || new Date().toISOString()).slice(0, 10);
  const label = String(payload?.meta?.label || "").trim();
  const labelSafe = label ? label.replace(/[^a-zA-Z0-9_\-\u0600-\u06FF ]/g, "").trim().replace(/\s+/g, "_") : "";
  const labelPart = labelSafe ? `_${labelSafe}` : "";
  sbtsDownloadJson(`SBTS_backup_${date}${labelPart}.json`, payload);
  toast("Backup downloaded.");
}

function sbtsDeleteLocalBackup(index) {
  if (!sbtsIsAdmin()) return toast("Admin only.");
  const arr = sbtsGetLocalBackups();
  if (!arr[index]) return;
  const ok = confirm("Delete this local backup? (This cannot be undone)");
  if (!ok) return;
  arr.splice(index, 1);
  sbtsLocalBackupsWrite(arr);
  sbtsOpenLocalBackups();
  toast("Local backup deleted.");
}


// Placeholder hooks for future Supabase backups (disabled in this build)
async function sbtsSupabaseBackupHook(_payload) {
  // TODO: upload payload to Supabase storage/table (disabled)
  return false;
}

function loadState() {
  const raw = SBTS_UTILS.LS.get("sbts_state");

  if (!raw) {
    const admin = {
      id: "admin-1",
      fullName: "System Admin",
      username: "admin",
      password: "admin123",
      role: "admin",
      phone: "",
      email: "",
      status: "active",
      jobTitle: "Admin",
      profileImage: null,
    };
    state.users.push(admin);
    state.permissions[admin.id] = {
      manageAreas: true,
      manageProjects: true,
      managePhaseOwnership: true,
      manageRequests: true,
      manageBlinds: true,
      changePhases: true,
      viewReports: true,
      manageReportsCards: true,
      manageCertificateSettings: true,
      manageTagSettings: true,
      manageWorkflowControl: true,
      editBranding: true,
      manageTrainingVisibility: true,
      manageRolesCatalog: true,
      manageFinalApprovals: true,
      manageUsers: true,
      // backward compat
      adminApproveUsers: true,
    };

    applyThemePreset("classic", true);
    applyFontSize(14);

    saveState();
    return;
  }

  const data = JSON.parse(raw);
  Object.assign(state, data);
  // defaults for new versions
  if (!state.ui) state.ui = { showTrainingPage: true };
  if (state.ui.showTrainingPage === undefined) state.ui.showTrainingPage = true;

  state.users = data.users || [];
  state.areas = data.areas || [];
  state.projects = data.projects || [];
  state.blinds = data.blinds || [];
  state.permissions = data.permissions || {};
  // Migration: older versions used `adminApproveUsers` and lacked new keys
  Object.keys(state.permissions).forEach((uid) => {
    const p = state.permissions[uid] || {};
    if (p.adminApproveUsers !== undefined && p.manageUsers === undefined) {
      p.manageUsers = !!p.adminApproveUsers;
    }
    if (p.manageFinalApprovals === undefined) {
      p.manageFinalApprovals = false;
    }
    if (p.managePhaseOwnership === undefined) {
      p.managePhaseOwnership = false;
    }
    if (p.manageRequests === undefined) {
      p.manageRequests = false;
    }
    if (p.manageTagSettings === undefined) {
      p.manageTagSettings = false;
    }
    state.permissions[uid] = p;
  });

  // Ensure admins always have Final Approvals permission
  try {
    (state.users || []).forEach((u) => {
      if (normalizeRole(u?.role) === "admin") {
        state.permissions[u.id] = state.permissions[u.id] || {};
        state.permissions[u.id].manageFinalApprovals = true;

        state.permissions[u.id].managePhaseOwnership = true;

        state.permissions[u.id].manageRequests = true;

        state.permissions[u.id].manageTagSettings = true;
      }
    });
  } catch (e) {}

  // notifications (per user)
  state.notifications = data.notifications || { byUser: {} };
  if (!state.notifications || typeof state.notifications !== 'object') state.notifications = { byUser: {} };
  if (!state.notifications.byUser || typeof state.notifications.byUser !== 'object') state.notifications.byUser = {};

  // requests
  state.requests = Array.isArray(data.requests) ? data.requests : [];

  state.themePreset = data.themePreset || "classic";
  state.ui = state.ui || {};
  state.ui.themePreset = state.themePreset;
  state.themeColor = data.themeColor || "#174c7e";
  state.ui.themeColor = state.themeColor;
  state.fontSize = data.fontSize || 14;
  state.ui.fontSize = state.fontSize;

  state.branding = { ...state.branding, ...(data.branding || {}) };
  state.certificate = { ...state.certificate, ...(data.certificate || {}) };

  // Tag theme (single global color)
  state.tagTheme = { ...state.tagTheme, ...(data.tagTheme || {}) };
  if (!state.tagTheme || typeof state.tagTheme !== 'object') state.tagTheme = { color: '#0F6D8C', audit: [] };
  if (!state.tagTheme.color) state.tagTheme.color = '#0F6D8C';
  if (!Array.isArray(state.tagTheme.audit)) state.tagTheme.audit = [];

  // Slip state
  state.slip = { ...(state.slip || { areaId: '', projectId: '', selectedIds: [] }), ...(data.slip || {}) };
  if (!Array.isArray(state.slip.selectedIds)) state.slip.selectedIds = [];

  applyThemePreset(state.themePreset, true);
  if (state.themePreset === "custom") applyThemeColor(state.themeColor);
  applyFontSize(state.fontSize);

  // Fix: ensure defaults exist
  if (!state.certificate.templates || state.certificate.templates.length === 0) {
    state.certificate.templates = [
      { id: "default", name: "Default" },
      { id: "clean", name: "Clean" },
    ];
  }
  if (!state.certificate.activeTemplate) state.certificate.activeTemplate = "default";

  // Fix: areaId for projects
  if (state.areas.length > 0) {
    const defaultAreaId = state.areas[0].id;
    state.projects.forEach((p) => { if (!p.areaId) p.areaId = defaultAreaId; });
  }

  // Fix: blinds fields
  state.blinds.forEach((b) => {
    // Migration: older versions may store history under different keys.
    // Prefer existing b.history; if missing/empty and b.historyLog exists, adopt it.
    if ((!b.history || !Array.isArray(b.history) || b.history.length === 0) && Array.isArray(b.historyLog) && b.historyLog.length) {
      b.history = b.historyLog;
    }
    if (!Array.isArray(b.history)) b.history = [];
    if (!b.finalApprovals) b.finalApprovals = {};
    if (!b.phase) b.phase = "broken";
    if (!b.size) b.size = "";
    if (!b.type) b.type = "Isolation Blind";
  });
}

/* ==========================
   HELPERS
========================== */
function phaseLabel(ph) {
  if (ph === "__final__") return "Final approvals";
  const cfg = getWorkflowConfig();
  const found = cfg?.phases?.find(p=>p.id===ph);
  if (found && found.label) return found.label;
  return phaseLabels[ph] || ph || "-";
}
function roleLabel(role) { return getRoleLabelById(role); }

function canUser(permKey) {
  if (!state.currentUser) return false;
  if (state.currentUser.role === "admin") return true;
  const perms = state.permissions[state.currentUser.id] || {};
  // Backward compatibility: older builds used `adminApproveUsers`.
  if (permKey === "manageUsers") return !!(perms.manageUsers || perms.adminApproveUsers);
  if (permKey === "adminApproveUsers") return !!(perms.manageUsers || perms.adminApproveUsers);
  return !!perms[permKey];
}

function requirePerm(permKey, message = 'No permission.') {
  if (canUser(permKey)) return true;
  toast(message);
  return false;
}

// Check permissions for an arbitrary user id (not the current user)
function userHasPerm(userId, permKey) {
  const u = (state.users || []).find((x) => x && x.id === userId);
  if (!u) return false;
  if (normalizeRole(u.role) === 'admin') return true;
  const perms = state.permissions[userId] || {};
  if (permKey === 'manageUsers') return !!(perms.manageUsers || perms.adminApproveUsers);
  if (permKey === 'adminApproveUsers') return !!(perms.manageUsers || perms.adminApproveUsers);
  return !!perms[permKey];
}

function nextPhaseOf(currentPhase) {
  const ids = wfProjectPhaseIds(state.currentProjectId, { includeInactive: false });
  const cur = currentPhase || "broken";
  if (cur === "__final__") return "__final__";
  const idx = ids.indexOf(cur);
  if (idx < 0) return ids[0] || "broken";
  if (idx >= ids.length - 1) return "__final__";
  return ids[idx + 1] || "__final__";
}

function isAllPhasesComplete(blind) {
  const lastId = wfLastActivePhaseId();
  // "__final__" means all phases finished and system is waiting for final approvals.
  if (blind.phase === "__final__") return true;
  return blind.phase === lastId;
}

function wfLastActivePhaseId() {
  const ids = wfProjectPhaseIds(state.currentProjectId, { includeInactive: false });
  return ids.length ? ids[ids.length - 1] : "inspection_ready";
}

function getFinalApprovalsModel(blind) {
  // If Safe Mode is ON: use Final Approvals Policy (directory) instead of Workflow Control.
  if (faIsSafeModeOn && faIsSafeModeOn()) {
    return faPolicyModelForBlind(blind);
  }

  // Final approvals are derived from Workflow Control config of the LAST ACTIVE phase
  const cfg = getWorkflowConfig();
  const lastId = wfLastActivePhaseId();
  const p = (cfg.phases || []).find((x) => x.id === lastId) || null;

  const required = [];
  const optional = [];

  // Base approvals from phase.approval
  if (p?.approval?.enabled && Array.isArray(p.approval.roles)) {
    p.approval.roles.forEach((roleId) => {
      const key = `final:${lastId}:role:${roleId}`;
      const label = `${ROLE_LABELS[roleId] || roleId} approval`;
      required.push({ key, label, role: roleId });
    });
  }

  // Extra approvals from phase.extra.items
  if (p?.extra?.enabled && Array.isArray(p.extra.items)) {
    p.extra.items.forEach((it, idx) => {
      const name = (it?.name || `Extra approval ${idx + 1}`).trim();
      const roles = Array.isArray(it?.roles) ? it.roles : [];
      const requiredFlag = it?.required !== false; // default required
      const key = `final:${lastId}:extra:${idx}:${name.replace(/\s+/g, "_").toLowerCase()}`;

      // If roles are specified, we treat each role as a signer option (any one). For v1, first role is used for permission check.
      const role = roles[0] || "admin";

      const entry = { key, label: name, role, roles, required: requiredFlag };
      (requiredFlag ? required : optional).push(entry);
    });
  }


  // D4: Slip Blind rule (optional): add Metal Foreman approval when enabled
  const isSlip = (blind?.type || "") === "Slip Blind";
  const slipRule = cfg?.rules?.slipBlindMetalForeman;
  if (isSlip && slipRule?.enabled) {
    const roleId = slipRule.role || "metal_foreman";
    const name = (slipRule.name || "Metal Foreman – Demolish").trim();
    const key = `rule:slip:${roleId}`;
    const exists = required.some(x => x.role === roleId) || optional.some(x => x.role === roleId) || required.some(x => x.key === key);
    if (!exists) {
      required.push({ key, label: name, role: roleId, required: true });
    }
  }

  // Backward compatibility: if nothing configured, fallback to legacy FINAL_APPROVALS (+ slip blind extra)
  if (!required.length && !optional.length) {
    const legacy = getFinalApprovalsForBlind_legacy(blind);
    return { required: legacy, optional: [] };
  }

  return { required, optional };
}

// Legacy behavior preserved for older configs
function getFinalApprovalsForBlind_legacy(blind) {
  const isSlip = (blind?.type || "") === "Slip Blind";
  const base = FINAL_APPROVALS.slice();
  return isSlip ? base.concat([EXTRA_SLIP_APPROVAL]) : base;
}

function getFinalApprovalsForBlind(blind) {
  const model = getFinalApprovalsModel(blind);
  // For completion checks we only require "required"
  return model.required;
}
function isAllFinalApprovalsDone(blind) {
  const fa = blind.finalApprovals || {};
  return getFinalApprovalsForBlind(blind).every((a) => faIsIdApprovedForKey(fa, a.key));
}

function isCertificateApproved(blind) {
  return isAllPhasesComplete(blind) && isAllFinalApprovalsDone(blind);
}

function formatDateDDMMYYYY(d = new Date()) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Backward-compatible helper used across Inbox / Notification Center rendering.
// Accepts Date | number(ms) | ISO string and returns a stable UI string.
function formatDate(input, opts = { withTime: true }) {
  try {
    if (!input) return "";
    const d = input instanceof Date ? input : new Date(input);
    if (Number.isNaN(d.getTime())) return "";
    const base = formatDateDDMMYYYY(d);
    if (!opts || opts.withTime === false) return base;
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${base} ${hh}:${mi}`;
  } catch (e) {
    return "";
  }
}

/* ==========================
   WORKFLOW RULES
========================== */
function normalizeRole(role) {
  const r = String(role || "").trim().toLowerCase();
  if (!r) return "";
  // Normalize common QA/QC variants
  if (r in {"qc":1, "qaqc":1, "qa/qc":1, "qa-qc":1, "qa qc":1, "quality":1, "quality control":1}) return "qaqc";
  if (r in {"supervisor":1, "sup":1}) return "supervisor";
  if (r in {"technician":1, "tech":1}) return "technician";
  if (r in {"safety":1, "safety officer":1}) return "safety";
  if (r in {"admin":1, "system admin":1}) return "admin";
  return r;
}

function normalizePhaseId(phaseId) {
  return String(phaseId || "").trim().toLowerCase();
}

function getProjectByIdSafe(projectId) {
  const pid = projectId || state.currentProjectId;
  return state.projects.find((p) => p.id === pid) || null;
}

function getPhaseOwners(projectId, phaseId) {
  const project = getProjectByIdSafe(projectId);
  const phasesObj = project?.phaseOwners?.phases || {};
  let arr = phasesObj?.[phaseId];
  if (Array.isArray(arr)) return arr;

  // Case-insensitive / normalized fallback
  const want = normalizePhaseId(phaseId);
  for (const k of Object.keys(phasesObj)) {
    if (normalizePhaseId(k) === want) {
      arr = phasesObj[k];
      if (Array.isArray(arr)) return arr;
    }
  }
  return [];
}

function isPhaseOwner(user, blind, toPhase) {
  if (!user) return false;
  const owners = getPhaseOwners(blind?.projectId, toPhase);
  return owners.includes(user.id);
}

/* ==========================
   WORKFLOW RULES
========================== */
function canTransition(fromPhase, toPhase, user, blind = null) {
  if (!user) return false;
  if (normalizeRole(user.role) === "admin") return true;
  if (!canUser("changePhases")) return false;

  const ids = wfPhaseIds({ includeInactive: true });
  const fromIndex = ids.indexOf(fromPhase || ids[0] || "broken");
  const toIndex = ids.indexOf(toPhase);

  if (toIndex < 0) return false;

  // no backward for non-admin
  if (toIndex < fromIndex) return false;

  // phase order enforcement (Admin can always skip; controlled by Settings)
  const enforceOrder = state.ui?.enforcePhaseOrder !== false;
  if (enforceOrder) {
    // only next step or same
    if (toIndex != fromIndex + 1 && toIndex != fromIndex) return false;
  }

  // Phase-based update policy (per project)
  const policy = getProjectUpdatePolicy(blind?.projectId);

  if (policy === "owners") {
    // Owners-only: must be assigned as an owner of the target phase
    return !!(blind && isPhaseOwner(user, blind, toPhase));
  }

  // Hybrid: owner OR role
  if (policy === "hybrid") {
    if (blind && isPhaseOwner(user, blind, toPhase)) return true;
  }

  // Role-based permission from workflow config (canUpdate)
  const cfg = getWorkflowConfig();
  const target = cfg.phases.find((p) => normalizePhaseId(p.id) === normalizePhaseId(toPhase));
  const allowedRoles = Array.isArray(target?.canUpdate) ? target.canUpdate : [];
  if (allowedRoles.length > 0) {
    const allowedN = allowedRoles.map(normalizeRole);
    return allowedN.includes(normalizeRole(user.role));
  }

  // Legacy required-role fallback (compat). If missing, deny.
  const requiredRole = WORKFLOW_REQUIRED_ROLE[toPhase];
  if (!requiredRole) return false;
  return normalizeRole(user.role) === normalizeRole(requiredRole);
}


/* ==========================
   HEADER + SIDEBAR USER BOX
========================== */
function applyBrandingToHeader() {
  const b = state.branding;

  const programTitleText = document.getElementById("programTitleText");
  const programSubtitleText = document.getElementById("programSubtitleText");
  const siteTitleText = document.getElementById("siteTitleText");
  const siteSubtitleText = document.getElementById("siteSubtitleText");
  const companyNameText = document.getElementById("companyNameText");
  const companySubText = document.getElementById("companySubText");
  const companyLogoBox = document.getElementById("companyLogoBox");

  // Optional texts: if empty => hide
  const setOptionalText = (el, value) => {
    if (!el) return;
    const v = (value || "").trim();
    el.textContent = v;
    el.style.display = v ? "block" : "none";
  };

  setOptionalText(programTitleText, b.programTitle);
  setOptionalText(programSubtitleText, b.programSubtitle);

  if (siteTitleText) siteTitleText.textContent = (b.siteTitle || "Shedgum Gas Plant");
  if (siteSubtitleText) siteSubtitleText.textContent = (b.siteSubtitle || "Smart Blind Tag System");

  setOptionalText(companyNameText, b.companyName);
  setOptionalText(companySubText, b.companySub);

  if (companyLogoBox) {
    if (b.companyLogo) {
      companyLogoBox.style.backgroundImage = `url(${b.companyLogo})`;
    } else {
      companyLogoBox.style.backgroundImage = "";
    }
  }

  // dates
  const today = formatDateDDMMYYYY(new Date());
  const topbarDate = document.getElementById("topbarDate");
  const sidebarDate = document.getElementById("sidebarDate");
  if (topbarDate) topbarDate.textContent = today;
  if (sidebarDate) sidebarDate.textContent = today;
}

function applySidebarUserBox() {
  const u = state.currentUser;
  const avatar = document.getElementById("sidebarAvatar");
  const nameEl = document.getElementById("sidebarUserName");
  const roleEl = document.getElementById("sidebarUserRole");

  if (!u) return;
  if (nameEl) nameEl.textContent = u.fullName || u.username || "User";
  if (roleEl) roleEl.textContent = (u.jobTitle && u.jobTitle.trim()) ? u.jobTitle : roleLabel(u.role);

  if (avatar) {
    if (u.profileImage) {
      avatar.style.backgroundImage = `url(${u.profileImage})`;
      avatar.style.backgroundSize = "cover";
      avatar.style.backgroundPosition = "center";
      avatar.textContent = "";
    } else {
      avatar.style.backgroundImage = "";
      avatar.textContent = (u.fullName || "U").trim().slice(0, 1).toUpperCase();
    }
  }
  const box = document.getElementById("sidebarUserBox");
  if (box) box.style.cursor = "pointer";
}


/* ==========================
   NAV
========================== */

// ==========================
// SUPABASE LIVE (Pilot)
// ==========================
function sbtsLiveModeEnabled() {
  try {
    const v = SBTS_UTILS.LS.get("sbts_live_mode");
    if (v === "0") return false;
    if (v === "1") return true;
  } catch (e) {}
  return !!window.SBTS_LIVE_MODE_DEFAULT;
}

function sbtsHasSupabaseConfig() {
  return !!(window.SBTS_SUPABASE_URL && window.SBTS_SUPABASE_ANON && window.supabase && typeof window.supabase.createClient === "function");
}

function sbtsSupabaseClient() {
  try {
    if (!sbtsHasSupabaseConfig()) return null;
    if (!window.__SBTS_SUPABASE__) {
      window.__SBTS_SUPABASE__ = window.supabase.createClient(window.SBTS_SUPABASE_URL, window.SBTS_SUPABASE_ANON);
    }
    return window.__SBTS_SUPABASE__;
  } catch (e) {
    console.warn("Supabase init failed", e);
    return null;
  }
}

async function sbtsLiveSignIn(email, password) {
  const client = sbtsSupabaseClient();
  if (!client) throw new Error("Supabase not configured");
  const { data, error } = await client.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data.user;
}

async function sbtsLiveSignOut() {
  const client = sbtsSupabaseClient();
  if (!client) return;
  await client.auth.signOut();
}

async function sbtsLiveEnsureProfile(user) {
  const client = sbtsSupabaseClient();
  if (!client || !user) return null;
  const { data: prof, error } = await client.from("profiles").select("id, full_name, role, job_title, avatar_url, theme").eq("id", user.id).maybeSingle();
  if (error) throw error;
  if (prof) return prof;
  const insert = { id: user.id, full_name: user.email || "User", role: "user" };
  const { data: created, error: err2 } = await client.from("profiles").insert(insert).select().maybeSingle();
  if (err2) throw err2;
  return created || insert;
}

async function sbtsLiveFetchProjects() {
  const client = sbtsSupabaseClient();
  if (!client) return [];
  const { data, error } = await client.from("projects").select("id, name, description, created_at").order("created_at", { ascending: false });
  if (error) throw error;
  return data || [];
}

function applyNavVisibility() {
  const trainingItem = document.getElementById("menuTrainingItem");
  if (trainingItem) {
    const visible = state.ui?.showTrainingPage !== false;
    trainingItem.style.display = visible ? "" : "none";
  }

  // Role/permission based menu rendering
  const setVis = (id, ok) => {
    const el = document.getElementById(id);
    if (el) el.style.display = ok ? "" : "none";
  };
  setVis("menuReports", canUser("viewReports"));
  setVis("menuReportsCards", canUser("manageReportsCards"));
  setVis("menuWorkflowControl", canUser("manageWorkflowControl"));
  setVis("menuUsers", canUser("manageUsers"));
  setVis("menuCertificateSettings", canUser("manageCertificateSettings"));
  setVis("menuTagSettings", canUser("manageTagSettings"));
  setVis("menuTagDesigner", canUser("manageTagSettings"));
}

// ==========================
// Mobile sidebar overlay
// ==========================
function openMobileSidebar() {
  if (!isMobileView()) return;
  const sb = document.getElementById("sidebar");
  const ov = document.getElementById("mobileSidebarOverlay");
  if (sb) sb.classList.add("open");
  if (ov) ov.classList.remove("hidden");
  document.body.classList.add("sbts-lock-scroll");
}

function closeMobileSidebar(silent) {
  const sb = document.getElementById("sidebar");
  const ov = document.getElementById("mobileSidebarOverlay");
  if (sb) sb.classList.remove("open");
  if (ov) ov.classList.add("hidden");
  document.body.classList.remove("sbts-lock-scroll");
  if (!silent && isMobileView()) {
    // no-op
  }
}

function toggleMobileSidebar(evt) {
  if (evt) evt.stopPropagation();
  const sb = document.getElementById("sidebar");
  if (!sb) return;
  if (sb.classList.contains("open")) closeMobileSidebar(true);
  else openMobileSidebar();
}

function openPage(pageId) {
  // Permission guards (hard lock)
  const deny = () => {
    toast("No permission");
    pageId = "dashboardPage";
  };

  // Visitor role: read-only limited navigation
  if (state.currentUser && normalizeRole(state.currentUser.role) === "visitor") {
    const allowed = ["dashboardPage", "reportsPage"];
    if (!allowed.includes(pageId)) {
      pageId = "dashboardPage";
    }
  }
  if (pageId === "workflowControlPage" && !canUser("manageWorkflowControl")) deny();
  if (pageId === "reportsCardsPage" && !canUser("manageReportsCards")) deny();
  if (pageId === "permissionsPage" && !canUser("manageUsers")) deny();
  if (pageId === "certificateSettingsPage" && !canUser("manageCertificateSettings")) deny();
  // Tag Settings merged into Tag Designer (Patch39.2)
  if (pageId === "tagDesignerPage" && !canUser("manageTagSettings")) deny();
  if (pageId === "reportsPage" && !canUser("viewReports")) deny();
  // Training visibility guard (for all users)
  if (pageId === "trainingPage" && state.ui?.showTrainingPage === false) {
    alert("Training page is hidden by admin.");
    return;
  }

  // Mobile UX: close sidebar overlay after navigation
  try { closeMobileSidebar(true); } catch (e) {}

  // Remember last visited page so browser refresh (or any UI re-render) returns user
  // to the same place instead of forcing Dashboard.
  try {
    state.ui = state.ui || {};
    state.ui.lastPage = pageId;
    SBTS_UTILS.LS.set("sbts_last_page", pageId);
    // Persist inside state as well (for single-storage compatibility)
    saveState();
  } catch (e) {
    // no-op
  }

  // Hide all pages then show selected page
  document.querySelectorAll(".page").forEach((p) => { p.classList.add("hidden"); p.classList.remove("active"); });
  const page = document.getElementById(pageId);
  if (page) { page.classList.remove("hidden"); page.classList.add("active"); }

  if (pageId === "workflowControlPage") {
    renderWorkflowControlPage();
  }

  if (pageId === "notificationsPage") {
    try {
      ensureNotificationsState();
      ensureRequestsArray();
      renderNotificationsInbox();
      updateNotificationsBadge();
    } catch (e) {
      console.error("[notificationsPage] render failed", e);
      showToast(
        "Notifications failed to render: " + (e && e.message ? e.message : "Unknown error"),
        "error"
      );
    }
  }

  // Reset scroll (fix: sometimes new page appears at bottom because window kept old scroll position)
  try {
    window.scrollTo({ top: 0, left: 0, behavior: "instant" });
  } catch (e) {
    window.scrollTo(0, 0);
  }
  const main = document.querySelector(".main-content");
  if (main) main.scrollTop = 0;

  // Active menu highlight
  document.querySelectorAll(".menu-item").forEach((m) => m.classList.remove("active"));
  [...document.querySelectorAll(".menu-item")].forEach((item) => {
    const txt = item.textContent.trim().toLowerCase();
    if (pageId === "dashboardPage" && txt.includes("dashboard")) item.classList.add("active");
    if (pageId === "areasPage" && txt.includes("areas")) item.classList.add("active");
    if (pageId === "projectsPage" && txt.includes("projects")) item.classList.add("active");
    if (pageId === "slipBlindPage" && txt.includes("slip blind")) item.classList.add("active");
    if (pageId === "reportsPage" && txt.includes("reports")) item.classList.add("active");
    if (pageId === "reportsCardsPage" && txt.includes("reports cards")) item.classList.add("active");
    if (pageId === "trainingPage" && txt.includes("training")) item.classList.add("active");
    if (pageId === "permissionsPage" && txt.includes("users")) item.classList.add("active");
    if (pageId === "settingsPage" && txt.includes("settings")) item.classList.add("active");
    if (pageId === "certificateSettingsPage" && txt.includes("certificate settings")) item.classList.add("active");
    // Tag Settings merged into Tag Designer
    if (pageId === "tagDesignerPage" && txt.includes("tag designer")) item.classList.add("active");
  });

  applyBrandingToHeader();
  applySidebarUserBox();
  applyNavVisibility();

  // Page hooks
  if (pageId === "dashboardPage") renderDashboard();
  if (pageId === "areasPage") renderAreas();
  if (pageId === "projectsPage") renderProjects();
  if (pageId === "slipBlindPage") renderSlipBlindPage();
  if (pageId === "reportsPage") initReportsUI();
  if (pageId === "permissionsPage") renderUsers();
  if (pageId === "settingsPage") { hydrateSettingsInputs(); }
  if (pageId === "reportsCardsPage") hydrateReportsCardsSettingsUI();
  if (pageId === "certificateSettingsPage") hydrateCertificateSettingsUI();
  // Tag Settings merged into Tag Designer
  if (pageId === "tagDesignerPage") hydrateTagDesignerUI();
  updateNotificationsBadge();

  // Record navigation (smart back + breadcrumbs)
  try {
    if (!NAV_SUPPRESS_PUSH) navPush(navEntryForCurrent(pageId));
  } catch (e) {}

  if(pageId==='settingsPage'){ setTimeout(faEnsureInit,0); }
}

/* ==========================
   AUTH
========================== */
function showAuthScreen() {
  document.getElementById("authContainer").classList.remove("hidden");
  document.getElementById("appContainer").classList.add("hidden");
  document.getElementById("loginForm").classList.remove("hidden");
  document.getElementById("registerForm").classList.add("hidden");
  // NOTE: Do not reference page-specific variables here.
  // Auth screen must be safe to open from anywhere (e.g., Logout).
}

function showAppScreen() {
  document.getElementById("authContainer").classList.add("hidden");
  document.getElementById("appContainer").classList.remove("hidden");
  applyBrandingToHeader();
  applySidebarUserBox();
  // Restore last page (Smart Resume). Fallback to Dashboard.
  let last = null;
  try {
    last = (state.ui && state.ui.lastPage) || SBTS_UTILS.LS.get("sbts_last_page");
  } catch (e) {
    last = (state.ui && state.ui.lastPage) || null;
  }
  const allowed = [
    "dashboardPage",
    "areasPage",
    "projectsPage",
    "slipBlindPage",
    "reportsPage",
    "reportsCardsPage",
    "trainingPage",
    "notificationsPage",
    "permissionsPage",
    "settingsPage",
    "certificateSettingsPage",
    // "tagSettingsPage" removed (merged)
    "tagDesignerPage",
    "workflowControlPage"
  ];
  if (!allowed.includes(last)) last = "dashboardPage";
  openPage(last);
}

function openRegister() {
  document.getElementById("loginForm").classList.add("hidden");
  document.getElementById("registerForm").classList.remove("hidden");
}

function openLogin() {
  document.getElementById("registerForm").classList.add("hidden");
  document.getElementById("loginForm").classList.remove("hidden");
}

async function handleLogin() {
  // Live Mode (Supabase): treat "Username" field as Email
  try {
    if (sbtsLiveModeEnabled() && sbtsHasSupabaseConfig()) {
      const email = document.getElementById("login_username").value.trim();
      const password = document.getElementById("login_password").value.trim();
      if (!email || !password) return alert("Please enter email & password.");

      const user = await sbtsLiveSignIn(email, password);
      const prof = await sbtsLiveEnsureProfile(user);

      // Build/merge current user into local state shape
      const liveUser = {
        id: user.id,
        fullName: (prof && prof.full_name) || user.email || "User",
        username: user.email || email,
        password: "",
        phone: "",
        email: user.email || email,
        role: (prof && prof.role) || "user",
        status: "active",
        jobTitle: (prof && prof.job_title) || "",
        profileImage: (prof && prof.avatar_url) || null,
        theme: (prof && prof.theme) || null,
      };

      // Keep a local entry so existing UI helpers work without refactor
      const existing = state.users.find((u) => u.id === liveUser.id);
      if (!existing) state.users.push(liveUser);

      // Minimal permissions defaults (RLS is the real security in Supabase).
      state.permissions[liveUser.id] = state.permissions[liveUser.id] || {
        manageAreas: false,
        manageProjects: false,
        manageBlinds: false,
        changePhases: true,
        viewReports: true,
        manageReportsCards: false,
        manageCertificateSettings: false,
        manageWorkflowControl: false,
        editBranding: false,
        manageTrainingVisibility: false,
        manageRolesCatalog: false,
        manageUsers: false,
        manageRequests: false,
        manageFinalApprovals: false,
      };

      // Admin shortcut if profile says admin
      if (String(liveUser.role).toLowerCase() === "admin") {
        state.permissions[liveUser.id].manageUsers = true;
        state.permissions[liveUser.id].manageProjects = true;
        state.permissions[liveUser.id].manageAreas = true;
        state.permissions[liveUser.id].editBranding = true;
        state.permissions[liveUser.id].manageRequests = true;
      }

      state.currentUser = liveUser;

      // Hydrate Projects from Supabase
      try {
        const remoteProjects = await sbtsLiveFetchProjects();
        state.projects = (remoteProjects || []).map((p) => ({
          id: p.id,
          name: p.name,
          description: p.description || "",
          status: "active",
          createdAt: p.created_at || new Date().toISOString(),
        }));
      } catch (e) {
        console.warn("Live projects load failed, falling back to local projects.", e);
      }

      saveState();
      showAppScreen();
      return;
    }
  } catch (e) {
    console.warn(e);
    alert("Live login failed. Check Supabase URL/Key, and that the account exists.");
    return;
  }
  const username = document.getElementById("login_username").value.trim();
  const password = document.getElementById("login_password").value.trim();
  const user = state.users.find((u) => u.username === username && u.password === password);
  if (!user) return alert("Invalid username or password");
  if (user.status === "pending") return alert("Your account is waiting for admin approval.");
  state.currentUser = user;
  saveState();
  showAppScreen();
  // If user came from QR public view, open blind directly
  if (runtime.afterLoginBlindId) {
    const bid = runtime.afterLoginBlindId;
    runtime.afterLoginBlindId = null;
    openBlindDetails(bid);
  }
}

function handleRegister() {
  const fullName = document.getElementById("reg_fullName").value.trim();
  const username = document.getElementById("reg_username").value.trim();
  const password = document.getElementById("reg_password").value.trim();
  const phone = document.getElementById("reg_phone").value.trim();
  const email = document.getElementById("reg_email").value.trim();
  const role = document.getElementById("reg_role").value;

  if (!fullName || !username || !password || !role) return alert("Please fill all required fields.");
  if (state.users.find((u) => u.username === username)) return alert("Username already exists.");

  const newUser = {
    id: crypto.randomUUID(),
    fullName,
    username,
    password,
    phone,
    email,
    role,
    status: "pending",
    jobTitle: "",
    profileImage: null,
    projectAssignments: [],
  };

  state.users.push(newUser);
  state.permissions[newUser.id] = {
    manageAreas: false,
    manageProjects: false,
    manageBlinds: false,
    changePhases: true,
    viewReports: true,
    manageReportsCards: false,
    manageCertificateSettings: false,
    manageWorkflowControl: false,
    editBranding: false,
    manageTrainingVisibility: false,
    manageRolesCatalog: false,
    manageUsers: false,
    // backward compat
    adminApproveUsers: false,
  };

    // Notify admins / user-approvers that a new account is pending approval (Inbox + Notifications)
  try {
    const approvers = getUserApproverIds().filter((id) => id && id !== newUser.id);
    if (approvers.length) {
      SBTS_ACTIVITY.pushRequest({
        requestType: "NEW_USER",
        title: "User registration approval",
        message: `${newUser.fullName} (@${newUser.username}) requested access. Review and approve/reject.`,
        scope: "user",
        recipients: approvers,
        priority: "high",
        actionKey: `user_approval:${newUser.id}`,
        meta: {
          userId: newUser.id,
          username: newUser.username,
          fullName: newUser.fullName,
          email: newUser.email || null,
          phone: newUser.phone || null,
          requestedRole: newUser.role || null,
        }
      });
    }

    // Confirmation to the user (no action)
    SBTS_ACTIVITY.pushNotification({
      category: "system",
      scope: "user",
      recipients: [newUser.id],
      title: "Registration submitted",
      message: "Your account request was submitted and is pending approval.",
      requiresAction: false,
      resolved: true,
      actorId: newUser.id
    });
  } catch(e) {}

  saveState();
  alert("Registered successfully. Your account is pending approval.");
  openLogin();
}

function logout() {
  // Live Mode: also sign out from Supabase (best-effort)
  try {
    if (sbtsLiveModeEnabled() && sbtsHasSupabaseConfig()) {
      sbtsLiveSignOut();
    }
  } catch (e) {}
  state.currentUser = null;
  saveState();
  showAuthScreen();
}

/* ==========================
   MODALS
========================== */
function openModal(id) {
  document.getElementById("modalOverlay").classList.remove("hidden");
  document.getElementById(id).classList.remove("hidden");
}
function closeModal(id) {
  document.getElementById(id).classList.add("hidden");
  const anyOpen = [...document.querySelectorAll(".modal")].some((m) => !m.classList.contains("hidden"));
  if (!anyOpen) document.getElementById("modalOverlay").classList.add("hidden");
}

/* ==========================
   DASHBOARD
========================== */
function renderDashboard() {
  document.getElementById("stat_totalAreas").textContent = state.areas.length;
  document.getElementById("stat_totalProjects").textContent = state.projects.length;
  document.getElementById("stat_totalBlinds").textContent = state.blinds.length;

  // Slip Blind KPI summary (All Projects) + Overview card (PATCH 45.1)
  try {
    const kpi = document.getElementById("dashboardSlipKpis");
    const ov = document.getElementById("dashboardSlipOverview");
    if (kpi) {
      const slip = (state.blinds||[]).filter(b => String(b.type||"").toLowerCase().includes("slip"));
      const total = slip.length;
      const done = slip.filter(b => String(b.phase||"").toLowerCase().includes("final")).length;
      const active = Math.max(0, total - done);

      // Completed but NOT demolished (Metal Foreman demolish approval not approved)
      const doneList = slip.filter(b => String(b.phase||"").toLowerCase().includes("final"));
      const demolished = doneList.filter(b => ((b.finalApprovals||{}).metal_foreman_demolish||{}).status === "approved").length;
      const notDemolished = Math.max(0, doneList.length - demolished);

      kpi.innerHTML = `
        <div class="kpi-pill">
          <div class="phase-left"><span class="phase-dot" style="background:#2D7FF9;"></span><div class="phase-name">Slip Blinds (All Projects)</div></div>
          <div class="phase-count">${total}</div>
        </div>
        <div class="kpi-pill">
          <div class="phase-left"><span class="phase-dot" style="background:#18B26A;"></span><div class="phase-name">Completed</div></div>
          <div class="phase-count">${done}</div>
        </div>
        <div class="kpi-pill">
          <div class="phase-left"><span class="phase-dot" style="background:#F4B400;"></span><div class="phase-name">In Progress</div></div>
          <div class="phase-count">${active}</div>
        </div>
      `;

      // Overview card
      if (ov) {
        const pct = total ? Math.round((done / total) * 100) : 0;
        ov.innerHTML = `
          <div class="slip-overview-compact-head">
            <div class="slip-overview-compact-title">Slip Blinds Overview</div>
            <div class="slip-overview-compact-right">Overall Progress: <b>${pct}%</b></div>
          </div>

          <div class="sbts-progress"><div class="sbts-progress-bar" style="width:${pct}%; background:#18B26A;"></div></div>

          <div class="slip-overview-compact-row">
            <div class="slip-mini-pill">
              <div class="slip-mini-label"><span class="dot" style="background:#18B26A;"></span>Completed</div>
              <div class="slip-mini-count">${done}</div>
            </div>
            <div class="slip-mini-pill">
              <div class="slip-mini-label"><span class="dot" style="background:#F4B400;"></span>Still Active</div>
              <div class="slip-mini-count">${active}</div>
            </div>
          </div>
        `;
        ov.classList.add("clickable");
        ov.onclick = () => slipOpenDashboard("");
}
    }
  } catch(e){}

  let myChanges = 0;
  if (state.currentUser) {
    state.blinds.forEach((b) => (b.history || []).forEach((h) => {
      if (h.userId === state.currentUser.id) myChanges++;
    }));
  }
  document.getElementById("stat_phaseChanges").textContent = myChanges;

  const byPhase = {};
  PHASES.forEach((p) => (byPhase[p] = 0));
  state.blinds.forEach((b) => { byPhase[b.phase || "broken"]++; });

  const container = document.getElementById("dashboardPhaseCounters");
  container.innerHTML = "";
  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
    const phColor = wfPhaseColor(ph);
    const card = document.createElement("div");
    card.className = "phase-card";
    card.setAttribute("data-phase", ph);
    card.innerHTML = `
      <div class="phase-left">
        <span class="phase-dot" style="background:${phColor}"></span>
        <div class="phase-name">${phaseLabel(ph)}</div>
      </div>
      <div class="phase-count">${byPhase[ph] || 0}</div>
    `;
    container.appendChild(card);
  });

  // Recent Activity (PATCH 45.2)
  try { renderDashboardRecentActivity(); } catch(e) {}
  try { initDashboardSlipSummaryToggle(); } catch(e) {}
  try { initDashboardRecentActivityToggle(); } catch(e) {}
}




function initDashboardSlipSummaryToggle(){
  const wrap = document.getElementById("dashboardSlipSummaryWrap");
  const btn = document.getElementById("toggleSlipSummaryBtn");
  if(!wrap || !btn) return;

  const KEY = "sbts_dash_slip_summary_collapsed";
  const apply = (collapsed) => {
    wrap.classList.toggle("collapsed", !!collapsed);
    btn.setAttribute("aria-expanded", String(!collapsed));
  };

  const initial = localStorage.getItem(KEY) === "1";
  apply(initial);

  btn.onclick = () => {
    const collapsed = !wrap.classList.contains("collapsed");
    localStorage.setItem(KEY, collapsed ? "1" : "0");
    apply(collapsed);
  };
}



function initDashboardRecentActivityToggle(){
  const wrap = document.getElementById("dashboardRecentActivityWrap");
  const btn = document.getElementById("toggleRecentActivityBtn");
  if(!wrap || !btn) return;

  const KEY = "sbts_dash_recent_activity_collapsed";
  const apply = (collapsed) => {
    wrap.classList.toggle("collapsed", !!collapsed);
    btn.setAttribute("aria-expanded", String(!collapsed));
  };

  const initial = localStorage.getItem(KEY) === "1";
  apply(initial);

  btn.onclick = () => {
    const collapsed = !wrap.classList.contains("collapsed");
    localStorage.setItem(KEY, collapsed ? "1" : "0");
    apply(collapsed);
  };
}




/* ==========================
   PATCH 45.2 - Recent Activity (Dashboard)
   - Renders last 5 events using existing blind.history (phase + finalApproval)
   - "View all activity" is a placeholder for now
========================== */

function sbtsBlindRef(blind){
  if(!blind) return "Blind";
  // Prefer human-friendly references (Tag No / Blind No / Line No). Avoid showing full UUID to the user.
  const tag = blind.tag_no || blind.tagNo || blind.tag || blind.tagNoText || blind.tag_no_text || blind.tagNumber;
  const blindNo = blind.blind_no || blind.blindNo || blind.blindNoText || blind.blindNumber;
  const line = blind.line_no || blind.lineNo || blind.line || blind.lineNumber;
  const area = blind.area || blind.areaName;
  const type = blind.type || blind.blindType;
  if(tag) return line ? `${tag} (${line})` : `${tag}`;
  if(blindNo) return line ? `Blind ${blindNo} (${line})` : `Blind ${blindNo}`;
  if(line) return type ? `Line ${line} (${type})` : `Line ${line}`;
  const id = blind.id || blind.uuid || blind.uid;
  if(id){
    const s = String(id);
    const short = s.includes("-") ? s.split("-")[0] : s.slice(0,8);
    return `Blind ${short}`;
  }
  return "Blind";
}


function sbtsBlindOfficialParts(blind){
  if(!blind) return {id:"Blind", restParts:[]};
  const tag = blind.tag_no || blind.tagNo || blind.tag || blind.tagNoText || blind.tag_no_text || blind.tagNumber || blind.name;
  const idText = tag ? String(tag) : sbtsBlindRef(blind);

  const areaObj = (blind.areaId ? state.areas.find(a => a.id === blind.areaId) : null);
  const areaName = blind.areaName || blind.area || (areaObj ? areaObj.name : "");

  const projectObj = (blind.projectId ? state.projects.find(p => p.id === blind.projectId) : null);
  const projectName = blind.projectName || (projectObj ? projectObj.name : "");

  const line = blind.line_no || blind.lineNo || blind.line || blind.lineNumber || blind.equipment || blind.equipmentNo || "";
  const size = blind.size || blind.diameter || blind.nps || "";

  const rest = [];
  if(areaName) rest.push(`Area ${areaName}`);
  if(projectName) rest.push(projectName);
  if(line) rest.push(line);
  if(size) rest.push(`${size}`);
  return { id: idText, restParts: rest };
}

function sbtsBlindOfficialText(blind){
  const p = sbtsBlindOfficialParts(blind);
  const id = p.id || "Blind";
  const rest = (p.restParts || []).filter(Boolean);
  if (rest.length) return `${id} | ${rest.join(" | ")}`;
  return `${id}`;
}


// Returns HTML like: <span class="ra-id">SB-025</span><span class="ra-rest"> | Area 2 | ...</span>
function sbtsBlindOfficialHTML(blind){
  const p = sbtsBlindOfficialParts(blind);
  const id = escapeHtml(p.id);
  const rest = p.restParts.map(escapeHtml);
  const restHtml = rest.length ? `<span class="ra-rest"> | ${rest.join(" | ")}</span>` : "";
  return `<span class="ra-id">${id}</span>${restHtml}`;
}

function sbtsFormatTimeAgo(dateObj){
  try{
    const d = (dateObj instanceof Date) ? dateObj : new Date(dateObj);
    if(isNaN(d)) return "";
    const diffMs = Date.now() - d.getTime();
    const mins = Math.floor(diffMs/60000);
    if(mins < 1) return "just now";
    if(mins < 60) return mins + " min ago";
    const hrs = Math.floor(mins/60);
    if(hrs < 24) return hrs + " hr ago";
    const days = Math.floor(hrs/24);
    return days + " day" + (days>1?"s":"") + " ago";
  }catch(e){ return ""; }
}

function sbtsGetRecentActivity(limit=5){
  const out = [];
  (state.blinds || []).forEach((b) => {
    (b.history || []).forEach((h) => {
      const d = new Date(h.date || h.createdAt || 0);
      if(!isNaN(d)) out.push({date:d, blind:b, h});
    });
  });
  out.sort((a,b)=> b.date - a.date);
  return out.slice(0, limit);
}

function renderDashboardRecentActivity(){
  const host = document.getElementById("dashboardRecentActivity");
  if(!host) return;
  const items = sbtsGetRecentActivity(5);
  if(!items.length){
    host.innerHTML = `<div class="recent-activity-empty">No recent activity yet.</div>`;
  } else {
    host.innerHTML = items.map(({date, blind, h}) => {
      const who = h.workerName || h.workerId || "User";
      const when = sbtsFormatTimeAgo(date);

      let iconClass="info", iconChar="•", actionText="Updated";
      if(h.type === "phase"){
        iconClass="info"; iconChar="↻";
        actionText = `${phaseLabel(h.toPhase)} updated`;
      } else if(h.type === "finalApproval"){
        iconClass="ok"; iconChar="✓";
        actionText = `${h.approvalName || "Final approval"} approved`;
      } else if(h.type === "certificateIssued"){
        iconClass="ok"; iconChar="🏷️";
        actionText = `Certificate issued`;
      } else {
        iconClass="info"; iconChar="•";
        actionText = `Updated`;
      }

      const refHtml = sbtsBlindOfficialHTML(blind);
      const whoHtml = `<span class="ra-who">${escapeHtml(who)}</span>`;
      const actionHtml = `<span class="ra-action"> — ${escapeHtml(actionText)}</span>`;
      const mainHtml = `${whoHtml} <span class="ra-on">on</span> ${refHtml}${actionHtml}`;

      const meta = `${when}${when?" • ":""}${new Date(date).toLocaleString()}`;

      return `
        <div class="recent-activity-item">
          <div class="recent-activity-ico ${iconClass}">${iconChar}</div>
          <div class="recent-activity-text">
            <div class="recent-activity-main">${mainHtml}</div>
            <div class="recent-activity-meta">${escapeHtml(meta)}</div>
          </div>
        </div>
      `;
    }).join("");
  }

  // View all placeholder
  const viewAll = document.getElementById("dashboardViewAllActivity");
  if(viewAll){
    viewAll.onclick = (e) => {
      e.preventDefault();
      toast("View All Activity (coming soon)");
    };
  }
}



/* ==========================
   AREAS
========================== */
function openAddAreaModal() {
  if (!canUser("manageAreas")) return alert("No permission.");
  state._editingAreaId = null;
  document.getElementById("areaModalTitle").textContent = "Add area";
  document.getElementById("addAreaName").value = "";
  openModal("addAreaModal");
}

function openEditAreaModal(areaId) {
  if (!canUser("manageAreas")) return alert("No permission.");
  const a = state.areas.find((x) => x.id === areaId);
  if (!a) return;
  state._editingAreaId = areaId;
  document.getElementById("areaModalTitle").textContent = "Edit area";
  document.getElementById("addAreaName").value = a.name;
  openModal("addAreaModal");
}

function confirmAddArea() {
  const name = document.getElementById("addAreaName").value.trim();
  if (!name) return alert("Enter area name.");

  const normalized = name.toLowerCase();
  const exists = state.areas.some((a) => (a.name || "").trim().toLowerCase() === normalized && a.id !== state._editingAreaId);
  if (exists) return alert("Area name already exists.");

  if (state._editingAreaId) {
    const a = state.areas.find((x) => x.id === state._editingAreaId);
    if (a) a.name = name;
    state._editingAreaId = null;
  } else {
    state.areas.push({ id: crypto.randomUUID(), name });

    // Notify admins (info)
    try {
      const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
      if (admins.length) {
        addNotification(admins, {
          type: "area",
          title: "New area created",
          message: `${name}`,
          requiresAction: false,
          actorId: state.currentUser?.id || null,
        });
      }
    } catch(e) {}

  }

  saveState();
  closeModal("addAreaModal");
  renderAreas();
  renderDashboard();
}

function deleteArea(id) {
  if (!canUser("manageAreas")) return alert("No permission.");
  if (!confirm("Delete area and related projects / blinds?")) return;

  sbtsAutoBackupBefore("Delete Area (and related data)")

  // Notify admins (info)
  try {
    const a = state.areas.find(x => x.id === id);
    const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
    if (admins.length) {
      addNotification(admins, {
        type: "area",
        title: "Area deleted",
        message: `${a?.name || "Area"} — removed (with its related projects/blinds).`,
        requiresAction: false,
        actorId: state.currentUser?.id || null,
      });
    }
  } catch(e) {}
;

  state.areas = state.areas.filter((a) => a.id !== id);
  state.projects = state.projects.filter((p) => p.areaId !== id);
  state.blinds = state.blinds.filter((b) => b.areaId !== id);

  saveState();
  renderAreas();
  renderDashboard();
}

function renderAreas() {
  const body = document.getElementById("areasTableBody");
  body.innerHTML = "";

  if (!state.areas || state.areas.length === 0) {
    body.innerHTML = `
      <tr>
        <td colspan="4">
          <div class="empty-state">
            <div class="empty-title">No areas created yet</div>
            <div class="empty-sub">Create your first area to organize projects and blinds.</div>
          </div>
        </td>
      </tr>`;
    return;
  }

  state.areas.forEach((area, index) => {
    const projectsInArea = state.projects.filter((p) => p.areaId === area.id);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${area.name}</td>
      <td>${projectsInArea.length}</td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn tiny" onclick="openAreaProjects('${area.id}')">View</button>
          <button class="secondary-btn tiny" onclick="openEditAreaModal('${area.id}')">Edit</button>
          <button class="secondary-btn tiny" onclick="deleteArea('${area.id}')">Delete</button>
        </div>
      </td>
    `;
    body.appendChild(tr);
  });
}

function openAreaProjects(areaId) {
  // This is the important part you asked:
  // When entering projects from an area, we lock filter + default area for add project
  state.currentProjectAreaFilter = areaId;
  openPage("projectsPage");
}

/* ==========================
   PROJECTS
========================== */
function fillAreasSelect(selectId, onlyAreaId = null) {
  const sel = document.getElementById(selectId);
  sel.innerHTML = "";

  const list = onlyAreaId ? state.areas.filter(a => a.id === onlyAreaId) : state.areas;

  list.forEach((a) => {
    const opt = document.createElement("option");
    opt.value = a.id;
    opt.textContent = a.name;
    sel.appendChild(opt);
  });

  // If only one area -> disable select (as requested)
  sel.disabled = !!onlyAreaId;
}

function updateProjectPreview() {
  const areaId = document.getElementById("addProjectAreaSelect")?.value;
  const areaName = state.areas.find(a => a.id === areaId)?.name || "-";
  const name = document.getElementById("addProjectName")?.value?.trim() || "-";
  const code = document.getElementById("addProjectCode")?.value?.trim() || "-";
  const desc = document.getElementById("addProjectDesc")?.value?.trim() || "-";

  const pvArea = document.getElementById("pvArea");
  const pvName = document.getElementById("pvName");
  const pvCode = document.getElementById("pvCode");
  const pvDesc = document.getElementById("pvDesc");

  if (pvArea) pvArea.textContent = areaName;
  if (pvName) pvName.textContent = name;
  if (pvCode) pvCode.textContent = code;
  if (pvDesc) pvDesc.textContent = desc;
}

function openAddProjectModal() {
  if (!canUser("manageProjects")) return alert("No permission.");
  if (state.areas.length === 0) return alert("Add at least one Area first.");

  state._editingProjectId = null;
  document.getElementById("projectModalTitle").textContent = "Add project";

  // If we came from Area -> Projects, lock to that area
  const lockedAreaId = state.currentProjectAreaFilter || null;
  fillAreasSelect("addProjectAreaSelect", lockedAreaId);

  if (lockedAreaId) {
    document.getElementById("addProjectAreaSelect").value = lockedAreaId;
  }

  document.getElementById("addProjectName").value = "";
  document.getElementById("addProjectCode").value = "";
  document.getElementById("addProjectDesc").value = "";
  updateProjectPreview();
  openModal("addProjectModal");
}

function openEditProjectModal(projectId) {
  if (!canUser("manageProjects")) return alert("No permission.");
  const p = state.projects.find((x) => x.id === projectId);
  if (!p) return;

  state._editingProjectId = projectId;
  document.getElementById("projectModalTitle").textContent = "Edit project";

  // Editing must show all areas
  fillAreasSelect("addProjectAreaSelect", null);

  document.getElementById("addProjectAreaSelect").value = p.areaId;
  document.getElementById("addProjectName").value = p.name || "";
  document.getElementById("addProjectCode").value = p.code || "";
  document.getElementById("addProjectDesc").value = p.desc || "";
  updateProjectPreview();
  openModal("addProjectModal");
}

function confirmAddProject() {
  const name = document.getElementById("addProjectName").value.trim();
  const areaId = document.getElementById("addProjectAreaSelect").value;
  const code = document.getElementById("addProjectCode").value.trim();
  const desc = document.getElementById("addProjectDesc").value.trim();

  if (!name || !areaId) return alert("Fill required fields.");

  // prevent duplicate project name (case-insensitive)
  const normalizedName = name.toLowerCase();
  const dup = state.projects.some((p) => (p.name || "").trim().toLowerCase() === normalizedName && p.id !== state._editingProjectId);
  if (dup) return alert("Project name already exists.");

  if (state._editingProjectId) {
    const p = state.projects.find((x) => x.id === state._editingProjectId);
    if (p) {
      p.name = name;
      p.areaId = areaId;
      p.code = code;
      p.desc = desc;
    }
    state.blinds.forEach((b) => {
      if (b.projectId === state._editingProjectId) b.areaId = areaId;
    });
    state._editingProjectId = null;
  } else {
    state.projects.push({ id: crypto.randomUUID(), name, areaId, code, desc });

    // Notify admins (info)
    try {
      const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
      if (admins.length) {
        const areaObj = state.areas.find(a => a.id === areaId);
        addNotification(admins, {
          type: "project",
          title: "New project created",
          message: `${name}${areaObj?.name ? " — Area " + areaObj.name : ""}`,
          projectId: null,
          blindId: null,
          requiresAction: false,
          actorId: state.currentUser?.id || null,
        });
      }
    } catch(e) {}

  }

  saveState();
  closeModal("addProjectModal");
  renderProjects();
  renderDashboard();
}

function deleteProject(id) {
  if (!canUser("manageProjects")) return alert("No permission.");
  if (!confirm("Delete this project and its blinds?")) return;

  sbtsAutoBackupBefore("Delete Project (and its blinds)")

  // Notify admins (info)
  try {
    const p = state.projects.find(x => x.id === id);
    const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
    if (admins.length) {
      addNotification(admins, {
        type: "project",
        title: "Project deleted",
        message: `${p?.name || "Project"} — removed (with its blinds).`,
        projectId: id,
        requiresAction: false,
        actorId: state.currentUser?.id || null,
      });
    }
  } catch(e) {}
;

  state.projects = state.projects.filter((p) => p.id !== id);
  state.blinds = state.blinds.filter((b) => b.projectId !== id);

  saveState();
  renderProjects();
  renderDashboard();
}

function renderProjects() {
  const body = document.getElementById("projectsTableBody");
  body.innerHTML = "";

  const list = state.currentProjectAreaFilter
    ? state.projects.filter((p) => p.areaId === state.currentProjectAreaFilter)
    : state.projects;

  if (!list || list.length === 0) {
    const label = state.currentProjectAreaFilter ? "in this area" : "";
    body.innerHTML = `
      <tr>
        <td colspan="5">
          <div class="empty-state">
            <div class="empty-title">No projects found ${label}</div>
            <div class="empty-sub">Add a project to start tracking blinds and phases.</div>
          </div>
        </td>
      </tr>`;
    return;
  }

  list.forEach((p, index) => {
    const area = state.areas.find((a) => a.id === p.areaId);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${p.name}</td>
      <td>${area ? area.name : "-"}</td>
      <td>${state.blinds.filter((b) => b.projectId === p.id).length}</td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn tiny" onclick="openProjectDetails('${p.id}')">Open</button>
          <button class="secondary-btn tiny" onclick="openEditProjectModal('${p.id}')">Edit</button>
          <button class="secondary-btn tiny" onclick="deleteProject('${p.id}')">Delete</button>
        </div>
      </td>
    `;
    body.appendChild(tr);
  });
}

/* ==========================
   PROJECT DETAILS
========================== */
function openProjectDetails(projectId) {
  // Robust open: never block navigation even if some widgets fail to render.
  state.currentProjectId = projectId;

  const project = (state.projects||[]).find((p) => p.id === projectId);
  if (!project) {
    try{ showToast("Project not found.", "warn"); }catch(e){}
    try{ openPage("projectsPage"); }catch(e){}
    return;
  }

  try {
    // Navigate first (inside try so any DOM issues are caught)
    openPage("projectDetailsPage");
    openProjectSubTab("blinds");

    const area = state.areas.find((a) => a.id === project.areaId);

    const titleEl = document.getElementById("projectDetailsTitle");
    if (titleEl) titleEl.textContent = "Project: " + project.name;

    const areaEl = document.getElementById("projectDetailArea");
    if (areaEl) areaEl.textContent = area ? area.name : "-";

    const blinds = state.blinds.filter((b) => b.projectId === projectId);

    const blindsEl = document.getElementById("projectDetailBlinds");
    if (blindsEl) blindsEl.textContent = String(blinds.length);

    const completedEl = document.getElementById("projectDetailCompleted");
    if (completedEl) completedEl.textContent = String(blinds.filter((b) => b.phase === "inspectionReady").length);

    // Render blocks (each is optional)
    try { if (document.getElementById("projectPhaseCards")) renderProjectPhaseCards(); } catch (e) { console.error("renderProjectPhaseCards failed:", e); }
    try { if (document.getElementById("projectBlindsTableBody")) renderProjectBlindsTable(); } catch (e) { console.error("renderProjectBlindsTable failed:", e); }
    try { if (document.getElementById("psPhaseSummary")) renderProjectSettingsSummary(); } catch (e) { console.error("renderProjectSettingsSummary failed:", e); }
  } catch (err) {
    // As a last resort, still attempt to show the page and avoid blocking the user.
    console.error("openProjectDetails failed:", err);
    try { openPage("projectDetailsPage"); } catch (_) {}
    // No alert (alerts make it feel like the button is broken). Console will have details.
  }
}




/* ==========================
   Project Settings (Per Project)
========================== */

// Phase-based ownership (per project)
// - Mode is stored at project.phaseOwners.mode (Basic/Advanced/Hybrid)
// - Owners are stored by phase id: project.phaseOwners.phases[phaseId] = [userId, ...]

function defaultProjectPhaseOwners() {
  return {
    mode: "advanced",
    policy: "hybrid", // owners | roles | hybrid
    phases: {},
    support: {},
  };
}

function ensureCurrentProjectPhaseOwners() {
  const project = state.projects.find((p) => p.id === state.currentProjectId);
  if (!project) return null;
  if (!project.phaseOwners) project.phaseOwners = defaultProjectPhaseOwners();

  project.phaseOwners.mode ??= "advanced";
  project.phaseOwners.policy ??= "hybrid"; // owners | roles | hybrid
  if (!project.phaseOwners.phases || typeof project.phaseOwners.phases !== "object") {
    project.phaseOwners.phases = {};
  }

  if (!project.phaseOwners.support || typeof project.phaseOwners.support !== "object") {
    project.phaseOwners.support = {};
  }

  // Heal to include current phases keys (keep existing selections)
  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
    if (!Array.isArray(project.phaseOwners.phases[ph])) project.phaseOwners.phases[ph] = [];
    if (!Array.isArray(project.phaseOwners.support[ph])) project.phaseOwners.support[ph] = [];
  });

  return project;
}


/* ==========================
   PHASE-BASED NOTIFICATIONS
   (per project, per phase owners)
========================== */

function ensureNotificationsState() {
  if (!state.notifications || typeof state.notifications !== "object") {
    // byUser: per-user notifications
    // global: notifications visible to all users (merged at read time)
    state.notifications = { byUser: {}, global: [] };
  }
  if (!state.notifications.byUser || typeof state.notifications.byUser !== "object") {
    state.notifications.byUser = {};
  }
  if (!Array.isArray(state.notifications.global)) {
    state.notifications.global = [];
  }
  
  // PATCH47.18 - Notification Rules (defaults)
  if (!state.notifications.rules || typeof state.notifications.rules !== "object") {
    state.notifications.rules = {
      defaults: { system: false, admin: false, warning: false, action: true },
      actionKeys: {
        NEW_USER_REQUEST: true,
        CERT_FINAL_APPROVAL: true,
        CERTIFICATE_APPROVAL: true,
        PHASE_APPROVAL: true
      }
    };
  } else {
    state.notifications.rules.defaults = state.notifications.rules.defaults || { system: false, admin: false, warning: false, action: true };
    state.notifications.rules.actionKeys = state.notifications.rules.actionKeys || {};
  }
return state.notifications;
}

/* ==========================
   PATCH47.18 - NOTIFICATION RULES HELPERS
========================== */
function getNotifRules(){
  ensureNotificationsState();
  return state.notifications.rules || { defaults: { system:false, admin:false, warning:false, action:true }, actionKeys:{} };
}
function notifRulesRequiresAction(category, actionKey){
  const rules = getNotifRules();
  const cat = (category||"system").toLowerCase();
  if (actionKey && typeof rules.actionKeys?.[actionKey] === "boolean") return !!rules.actionKeys[actionKey];
  if (typeof rules.defaults?.[cat] === "boolean") return !!rules.defaults[cat];
  return (cat === "action");
}


function applyNotifRulesToAllNotifications(){
  ensureNotificationsState();
  const applyArr = (arr)=>{
    if (!Array.isArray(arr)) return;
    arr.forEach(n=>{
      const cat = (n.category || "system").toLowerCase();
      // Apply rules to ALL notifications (requested)
      n.requiresAction = notifRulesRequiresAction(cat, n.actionKey);
    });
  };
  try{
    Object.keys(state.notifications.byUser || {}).forEach(uid=>{
      applyArr(state.notifications.byUser[uid]);
    });
    applyArr(state.notifications.global);
  }catch(e){}
}

function hydrateNotifRulesSettings(){
  const sec = document.getElementById("notifRulesSection");
  if (!sec) return;
  // admin only
  const canEdit = !!(state.currentUser && normalizeRole(state.currentUser.role) === "admin");
  sec.style.opacity = canEdit ? "" : "0.55";
  sec.querySelectorAll("input,button").forEach(el=>{
    el.disabled = !canEdit;
  });

  const rules = getNotifRules();
  const setChk = (id, val) => { const el = document.getElementById(id); if (el) el.checked = !!val; };

  setChk("nr_def_system", !!rules.defaults?.system);
  setChk("nr_def_admin", !!rules.defaults?.admin);
  setChk("nr_def_warning", !!rules.defaults?.warning);

  setChk("nr_key_new_user", rules.actionKeys?.NEW_USER_REQUEST !== undefined ? !!rules.actionKeys.NEW_USER_REQUEST : true);
  setChk("nr_key_cert_final", rules.actionKeys?.CERT_FINAL_APPROVAL !== undefined ? !!rules.actionKeys.CERT_FINAL_APPROVAL : true);
  setChk("nr_key_phase_approval", rules.actionKeys?.PHASE_APPROVAL !== undefined ? !!rules.actionKeys.PHASE_APPROVAL : true);
}

function saveNotifRulesFromSettings(){
  if (!(state.currentUser && normalizeRole(state.currentUser.role) === "admin")) return toast("Admins only");
  ensureNotificationsState();
  const rules = getNotifRules();

  const getChk = (id) => !!document.getElementById(id)?.checked;

  rules.defaults = rules.defaults || { system:false, admin:false, warning:false, action:true };
  rules.defaults.system = getChk("nr_def_system");
  rules.defaults.admin = getChk("nr_def_admin");
  rules.defaults.warning = getChk("nr_def_warning");
  rules.defaults.action = true;

  rules.actionKeys = rules.actionKeys || {};
  rules.actionKeys.NEW_USER_REQUEST = getChk("nr_key_new_user");
  rules.actionKeys.CERT_FINAL_APPROVAL = getChk("nr_key_cert_final");
  rules.actionKeys.PHASE_APPROVAL = getChk("nr_key_phase_approval");

  state.notifications.rules = rules;
  saveState();
  toast("Notification rules saved");
  // Apply to all existing notifications immediately
  applyNotifRulesToAllNotifications();
  renderNotificationsInbox();
  renderNotificationsDrawer();
  try { renderNotificationsInbox(); } catch(e){}
  try { updateNotificationsBadge(); } catch(e){}
}

function resetNotifRulesToDefault(){
  if (!(state.currentUser && normalizeRole(state.currentUser.role) === "admin")) return toast("Admins only");
  ensureNotificationsState();
  state.notifications.rules = {
    defaults: { system:false, admin:false, warning:false, action:true },
    actionKeys: { NEW_USER_REQUEST:true, CERT_FINAL_APPROVAL:true, CERTIFICATE_APPROVAL:true, PHASE_APPROVAL:true }
  };
  saveState();
  hydrateNotifRulesSettings();
  toast("Rules reset");
  try { renderNotificationsInbox(); } catch(e){}
  try { updateNotificationsBadge(); } catch(e){}
}



/* ==========================
   PATCH 47.5 - SBTS ACTIVITY ENGINE (Clean Core)
   - Single source for: Notifications + Inbox Requests + Recent Activity hooks
   - Supports scopes: user / project / global
   - Defensive guards to avoid "badge without items" & script errors
========================== */

function getProjectMemberIds(projectId){
  try {
    const pid = projectId || null;
    if(!pid) return [];
    const users = Array.isArray(state.users) ? state.users : [];
    const ids = users
      .filter(u => u && u.status === "active" && Array.isArray(u.projectAssignments) && u.projectAssignments.includes(pid))
      .map(u => u.id)
      .filter(Boolean);
    return Array.from(new Set(ids));
  } catch(e){
    return [];
  }
}

const SBTS_ACTIVITY = (() => {
  const CAT = {
    action:  { color: "red",    icon: "icon-action",  label: "Action required" },
    system:  { color: "blue",   icon: "icon-system",  label: "System" },
    warning: { color: "yellow", icon: "icon-warning", label: "Warning" },
    admin:   { color: "gray",   icon: "icon-admin",   label: "Admin note" },
  };

  function normalizeScope(scope){
    const s = (scope || "user").toLowerCase();
    if(["user","project","global"].includes(s)) return s;
    return "user";
  }

  function resolveRecipients({ scope, projectId, recipients }){
    const sc = normalizeScope(scope);
    let rec = [];
    if (Array.isArray(recipients)) rec = recipients.filter(Boolean);
    else if (recipients) rec = [recipients];

    if (sc === "global"){
      // global uses "*" broadcast bucket
      return ["*"];
    }
    if (sc === "project"){
      // project = all project members + any explicitly provided
      const members = getProjectMemberIds(projectId);
      return Array.from(new Set([ ...members, ...rec ]));
    }
    return rec;
  }

  function pushNotification(opts){
    const o = (opts && typeof opts === "object") ? opts : {};
    const category = (o.category || "system").toLowerCase();
    const cat = CAT[category] || CAT.system;

    const rec = resolveRecipients({
      scope: o.scope,
      projectId: o.projectId,
      recipients: o.recipients
    });

    // Map to existing notification payload shape
    addNotification(rec, {
      type: category.toUpperCase(),
      title: o.title || cat.label,
      message: o.message || "",
      projectId: o.projectId || null,
      blindId: o.blindId || null,
      phaseId: o.phaseId || null,
      fromPhase: o.fromPhase || null,
      toPhase: o.toPhase || null,
      actorId: o.actorId || null,
      requiresAction: (typeof o.requiresAction === "boolean") ? o.requiresAction : notifRulesRequiresAction(category, o.actionKey),
      resolved: !!o.resolved,
      actionKey: o.actionKey || null,
      archived: !!o.archived,
    });
  }

  function pushRequest(opts){
    const o = (opts && typeof opts === "object") ? opts : {};
    ensureRequestsArray();

    const category = "action";
    const cat = CAT.action;
    const now = new Date().toISOString();
    const req = {
      id: uid("req"),
      type: o.requestType || "GENERIC",
      kind: requestTypeLabel(o.requestType || "GENERIC"),
      status: "pending",
      priority: o.priority || "normal",
      projectId: o.projectId || null,
      projectName: o.projectName || (o.projectId ? (state.projects||[]).find(p=>p.id===o.projectId)?.name : null),
      blindId: o.blindId || null,
      phaseId: o.phaseId || null,
      targetUser: o.targetUser || null,
      reason: o.reason || null,
      requestedBy: o.requestedBy || getCurrentUserIdStable() || null,
      requestedByName: o.requestedByName || (state.currentUser?.fullName || state.currentUser?.username || "User"),
      createdAt: now,
      updatedAt: now,
      comments: Array.isArray(o.comments) ? o.comments : [],
      meta: o.meta || {},
    };

    // Store assignees/recipients on the request for threads & routing
    try {
      const raw = Array.isArray(o.recipients) ? o.recipients : (Array.isArray(o.assigneeIds) ? o.assigneeIds : []);
      req.assigneeIds = Array.from(new Set(raw.filter(Boolean)));
    } catch(e) { req.assigneeIds = req.assigneeIds || []; }

    state.requests.unshift(req);
    saveState();

    // notify assignees / approvers
    const actionKey = o.actionKey || `request:${req.id}`;
    pushNotification({
      category,
      scope: o.scope || "user",
      projectId: req.projectId,
      recipients: o.recipients || o.assigneeIds || [],
      title: o.notifTitle || (o.title || cat.label),
      message: o.notifMessage || o.message || "",
      requiresAction: true,
      resolved: false,
      actionKey,
      actorId: o.actorId || req.requestedBy,
      blindId: req.blindId,
      phaseId: req.phaseId
    });

    return req;
  }

  function seedLiveScenarios(){
    // Seed realistic, near-real workflows into Inbox/Notifications.
    // This is intentionally repeatable (adds more samples each time) for practical testing.
    ensureNotificationsState();
    ensureRequestsArray();

    const marker = "sbts_seed_live_scenarios_v2_count";
    const prev = parseInt(localStorage.getItem(marker) || "0", 10) || 0;
    const run = prev + 1;
    localStorage.setItem(marker, String(run));

    const usersArr = Array.isArray(state.users) ? state.users : [];
    const adminId = usersArr.find(u => u && String(u.role||"").toLowerCase()==="admin" && u.status==="active")?.id
      || usersArr.find(u=>u && String(u.role||"").toLowerCase()==="admin")?.id
      || usersArr[0]?.id || null;

    // Ensure at least one project exists
    state.projects = Array.isArray(state.projects) ? state.projects : [];
    state.blinds = Array.isArray(state.blinds) ? state.blinds : [];

    let project = state.projects.find(p => p && p.status !== "archived");
    if (!project) {
      project = { id: uid("p"), name: "SGP – SBTS Demo Project", status: "active", createdAt: new Date().toISOString() };
      state.projects.unshift(project);
    }

    // Ensure demo blinds exist under that project
    const pid = project.id;
    let blindA = state.blinds.find(b => (b.projectId||b.project_id) === pid && (b.tagNo||b.tag_no||b.name) === "BL-101");
    let blindB = state.blinds.find(b => (b.projectId||b.project_id) === pid && (b.tagNo||b.tag_no||b.name) === "BL-102");

    if (!blindA) {
      blindA = { id: uid("b"), projectId: pid, tagNo: "BL-101", area: "Area 2", lineNo: "D-111", size: '10"', rating: "300#", type: "slip", createdAt: new Date().toISOString() };
      state.blinds.unshift(blindA);
    }
    if (!blindB) {
      blindB = { id: uid("b"), projectId: pid, tagNo: "BL-102", area: "Area 3", lineNo: "C-204", size: '12"', rating: "600#", type: "slip", createdAt: new Date().toISOString() };
      state.blinds.unshift(blindB);
    }

    // --- Scenario 1: New User Approval (Admin Action)
    pushRequest({
      requestType: "NEW_USER",
      title: "User registration approval",
      message: "New user requested access. Assign role and approve/reject.",
      recipients: adminId ? [adminId] : [],
      scope: "user",
      priority: "high",
      actionKey: `user_approval:${uid("u")}`,
      meta: { username: `fahad${run}`, fullName: "Fahad Al-Otaibi", requestedRoleHint: "fitter" }
    });

    // --- Scenario 2: Blind workflow updates (Info) + Final approval (Action)
    pushNotification({
      category: "system",
      scope: "user",
      projectId: pid,
      blindId: blindA.id,
      recipients: adminId ? [adminId] : [],
      title: `BL-101 – PT Completed`,
      message: `PT completed for BL-101 (Area 2 • D-111 • 10" • 300#). Ready for QA/QC review.`,
      requiresAction: false,
      resolved: true,
      actionKey: `blind_update:pt:${blindA.id}`
    });

    pushNotification({
      category: "action",
      scope: "user",
      projectId: pid,
      blindId: blindA.id,
      recipients: adminId ? [adminId] : [],
      title: `Final Approval Required – BL-101`,
      message: `Certificate / Final Approval pending for BL-101. Review and approve.`,
      requiresAction: true,
      resolved: false,
      actionKey: `final_approval:${pid}:${blindA.id}`
    });

    // Add another related update under same Project>Blind thread
    pushNotification({
      category: "admin",
      scope: "user",
      projectId: pid,
      blindId: blindA.id,
      recipients: adminId ? [adminId] : [],
      title: `Supervisor Note – BL-101`,
      message: `Ensure isolation boundaries verified before closing BL-101 approval.`,
      requiresAction: false,
      resolved: true,
      actionKey: `note:${pid}:${blindA.id}:${run}`
    });

    // --- Scenario 3: Another blind (separate thread under same project)
    pushNotification({
      category: "warning",
      scope: "user",
      projectId: pid,
      blindId: blindB.id,
      recipients: adminId ? [adminId] : [],
      title: `BL-102 – Mismatch detected`,
      message: `Size/Rating mismatch flagged for BL-102. Verify line data before proceeding.`,
      requiresAction: true,
      resolved: false,
      actionKey: `blind_warning:mismatch:${blindB.id}`
    });

    // --- Scenario 4: Project-level update (no blind) (Project bucket)
    pushNotification({
      category: "admin",
      scope: "user",
      projectId: pid,
      recipients: adminId ? [adminId] : [],
      title: `Project Update – Shift Handover`,
      message: `Reminder: Update all in-progress blinds before shift handover.`,
      requiresAction: false,
      resolved: true,
      actionKey: `project_note:${pid}:${run}`
    });

    saveState();
    return true;
  }


  function seedLiveScenariosForActing(actingUserId){
    // Seed scenarios tailored to the currently "acting as" user (role-aware).
    ensureNotificationsState();
    ensureRequestsArray();

    const usersArr = Array.isArray(state.users) ? state.users : [];
    const acting = usersArr.find(u=>u && u.id===actingUserId) || usersArr[0] || null;
    const roleId = normalizeRoleId(acting?.role || "");

    // ensure a demo project + two blinds exist (reuse same logic as seedLiveScenarios)
    state.projects = Array.isArray(state.projects) ? state.projects : [];
    state.blinds = Array.isArray(state.blinds) ? state.blinds : [];

    let project = state.projects.find(p => p && p.status !== "archived");
    if (!project) {
      project = { id: uid("p"), name: "SGP – SBTS Demo Project", status: "active", createdAt: new Date().toISOString() };
      state.projects.unshift(project);
    }
    const pid = project.id;

    let blindA = state.blinds.find(b => (b.projectId||b.project_id) === pid && (b.tagNo||b.tag_no||b.name) === "BL-101");
    let blindB = state.blinds.find(b => (b.projectId||b.project_id) === pid && (b.tagNo||b.tag_no||b.name) === "BL-102");

    if (!blindA) {
      blindA = { id: uid("b"), projectId: pid, tagNo: "BL-101", area: "NGL", lineNo: "10-NGL-001", size: "8", rating: "300", type: "slip", status: "active", createdAt: new Date().toISOString() };
      state.blinds.unshift(blindA);
    }
    if (!blindB) {
      blindB = { id: uid("b"), projectId: pid, tagNo: "BL-102", area: "NGL", lineNo: "10-NGL-002", size: "6", rating: "300", type: "slip", status: "active", createdAt: new Date().toISOString() };
      state.blinds.unshift(blindB);
    }

    // Ensure at least one pending user exists (for New User Approval demo)
    const pendingUser = usersArr.find(u => u && (u.status === "pending" || u.pending === true));
    let newUser = pendingUser;
    if(!newUser){
      newUser = { id: uid("u"), fullName: "New User (Demo)", username: "new.user", role: "", status: "pending", createdAt: new Date().toISOString(), __demo: true };
      state.users = Array.isArray(state.users) ? state.users : [];
      state.users.unshift(newUser);
    }

    const now = new Date();
    const isoNow = now.toISOString();
    const who = acting?.fullName || acting?.username || roleId || "User";

    // marker per role
    const marker = "sbts_seed_role_scenarios_v1_" + (roleId || "unknown");
    const prev = parseInt(localStorage.getItem(marker) || "0", 10) || 0;
    const run = prev + 1;
    localStorage.setItem(marker, String(run));

    // Helper to target "my actions"
    const me = acting?.id || actingUserId;
    const adminId = usersArr.find(u => normalizeRoleId(u.role||"")==="admin" && u.status==="active")?.id
      || usersArr.find(u=>normalizeRoleId(u.role||"")==="admin")?.id
      || usersArr[0]?.id || null;

    // Always add a general update to the acting user (so Today/This Week light up)
    addNotification(me, {
      type: "update",
      title: `Shift update • ${project.name}`,
      message: `${who} — Demo run ${run}: inbox update created for testing filters.`,
      projectId: pid,
      actorId: adminId,
      createdBy: "System",
      requiresAction: false,
      meta: { projectId: pid }
    });

    // Role-aware scenarios
    if(roleId === "qc" || roleId === "inspection"){
      // QC: PT completed + Final approval required
      addNotification(me, {
        type: "update",
        title: `PT Completed • ${blindA.tagNo}`,
        message: `PT test completed for ${blindA.tagNo}. Ready for QC review.`,
        projectId: pid,
        blindId: blindA.id,
        actorId: adminId,
        requiresAction: false,
        meta: { projectId: pid, blindId: blindA.id }
      });
      addNotification(me, {
        type: "request",
        requiresAction: true,
        actionKey: "final_approval_required",
        title: `Final Approval required • ${blindA.tagNo}`,
        message: `Please review documents and approve final certificate for ${blindA.tagNo}.`,
        projectId: pid,
        blindId: blindA.id,
        actorId: adminId,
        meta: { projectId: pid, blindId: blindA.id, scope: "certificate" }
      });
    } else if(roleId === "safety"){
      addNotification(me, {
        type: "warning",
        title: `JSA review required • ${project.name}`,
        message: `New job started in ${blindB.area}. Please verify JSA / dynamic risk assessment.`,
        projectId: pid,
        actorId: adminId,
        requiresAction: true,
        actionKey: "jsa_review",
        meta: { projectId: pid, area: blindB.area }
      });
    } else if(roleId === "operation_foreman" || roleId === "metal_foreman" || roleId === "supervisor"){
      addNotification(me, {
        type: "request",
        requiresAction: true,
        actionKey: "owner_change_request",
        title: `Owner change request • ${project.name}`,
        message: `A request to change owner/assignee for phase "Preparation" has been submitted.`,
        projectId: pid,
        actorId: adminId,
        meta: { projectId: pid, phase: "Preparation" }
      });
      addNotification(me, {
        type: "update",
        title: `Assignment note • ${blindB.tagNo}`,
        message: `Work assigned: prepare access & barricade, then coordinate with QC.`,
        projectId: pid,
        blindId: blindB.id,
        actorId: adminId,
        requiresAction: false,
        meta: { projectId: pid, blindId: blindB.id }
      });
    } else if(roleId === "admin"){
      // Admin sees new user approval as action
      addNotification(me, {
        type: "request",
        requiresAction: true,
        actionKey: "new_user_approval",
        title: `New user approval • ${newUser.fullName || newUser.username}`,
        message: `A new user requested access. Assign a role and approve.`,
        actorId: me,
        meta: { userId: newUser.id }
      });
      // and a blind action
      addNotification(me, {
        type: "request",
        requiresAction: true,
        actionKey: "final_approval_required",
        title: `Final Approval required • ${blindA.tagNo}`,
        message: `Admin review: approve final certificate for ${blindA.tagNo}.`,
        projectId: pid,
        blindId: blindA.id,
        actorId: me,
        meta: { projectId: pid, blindId: blindA.id, scope: "certificate" }
      });
    } else {
      // default: provide a generic action assigned to user
      addNotification(me, {
        type: "request",
        requiresAction: true,
        actionKey: "action_required",
        title: `Action required • ${project.name}`,
        message: `Please review the latest updates and take action as needed.`,
        projectId: pid,
        actorId: adminId,
        meta: { projectId: pid }
      });
    }

    saveState();
    return { ok: true, roleId, run };
  }

  return {
    CAT,
    pushNotification,
    pushRequest,
    seedLiveScenarios,
    seedLiveScenariosForActing
  };
})();

// SLA timers for action notifications (v1: global 24h)
const SLA_HOURS = 24;
function smartLinkHtml(label, kind, id){
  if(!label) return "";
  const safeLabel = escapeHtml(String(label));
  const safeKind = escapeHtml(String(kind||""));
  const safeId = escapeHtml(String(id||""));
  return `<a href="#" class="sbts-smart-link" data-kind="${safeKind}" data-id="${safeId}">${safeLabel}</a>`;
}

function getNotifContextLinks(n){
  try{
    const parts=[];
    const projects = Array.isArray(state.projects) ? state.projects : [];
    const blinds = Array.isArray(state.blinds) ? state.blinds : [];
    const users = Array.isArray(state.users) ? state.users : [];

    if(n.projectId){
      const p = projects.find(x=>x && x.id===n.projectId);
      parts.push(`<span><i class="ph ph-briefcase"></i> ${smartLinkHtml(p?.name||n.projectId,"project",n.projectId)}</span>`);
    } else {
      parts.push(`<span><i class="ph ph-briefcase"></i> <span class="tiny muted">System</span></span>`);
    }

    if(n.blindId){
      const b = blinds.find(x=>x && x.id===n.blindId);
      const lbl = b?.name || n.blindRef || n.blindId;
      parts.push(`<span><i class="ph ph-tag"></i> ${smartLinkHtml(lbl,"blind",n.blindId)}</span>`);
    }

    const ak = n.actionKey ? String(n.actionKey) : "";
    if(ak.startsWith("user_approval:")){
      const uid = ak.slice("user_approval:".length);
      const u = users.find(x=>x && x.id===uid);
      parts.push(`<span><i class="ph ph-user"></i> ${smartLinkHtml(u?.name||uid,"user",uid)}</span>`);
    } else if(ak.startsWith("assign_projects:")){
      const uid = ak.slice("assign_projects:".length);
      const u = users.find(x=>x && x.id===uid);
      parts.push(`<span><i class="ph ph-user"></i> ${smartLinkHtml(u?.name||uid,"user",uid)}</span>`);
    } else if(ak.startsWith("req:")){
      const rid = ak.slice(4);
      parts.push(`<span><i class="ph ph-chat-circle-text"></i> ${smartLinkHtml("Open thread","request",rid)}</span>`);
    }
    return parts.join("");
  }catch(e){
    return "";
  }
}

function initSmartLinkDelegation(){
  if(window.__sbtsSmartLinkBound) return;
  window.__sbtsSmartLinkBound = true;
  document.addEventListener("click", (ev)=>{
    const a = ev.target && ev.target.closest ? ev.target.closest("a.sbts-smart-link") : null;
    if(!a) return;
    ev.preventDefault();
    const kind = a.getAttribute("data-kind");
    const id = a.getAttribute("data-id");
    try{
      if(kind==="project"){
        try{ closeNotificationsDrawer(); }catch(_){}
        openProjectDetails(id);
        return;
      }
      if(kind==="blind"){
        try{ closeNotificationsDrawer(); }catch(_){}
        openBlindDetails(id);
        return;
      }
      if(kind==="user"){
        state.ui = state.ui || {};
        state.ui.usersFocusId = id;
        saveState();
        try{ closeNotificationsDrawer(); }catch(_){}
        openPage("usersPage");
        return;
      }
      if(kind==="request"){
        state.ui = state.ui || {};
        state.ui.reqInboxFocusId = id;
        saveState();
        try{ closeNotificationsDrawer(); }catch(_){}
        openRequestsInbox();
        return;
      }
    }catch(e){
      try{ toast("Could not open item."); }catch(_){}
    }
  }, true);
}

function notifSlaPill(n) {
  try {
    if (!n || !n.requiresAction || n.resolved) return "";
    const t = (typeof n.ts === "number") ? n.ts : Date.parse(n.ts || "");
    if (!t || Number.isNaN(t)) return "";
    const now = Date.now();
    const slaMs = SLA_HOURS * 60 * 60 * 1000;
    const remain = (t + slaMs) - now;
    const mins = Math.round(Math.abs(remain) / 60000);
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    const fmt = (h > 0) ? `${h}h ${m}m` : `${m}m`;
    if (remain <= 0) {
      return `<span class="notif-badge-pill notif-pill-sla overdue" title="SLA ${SLA_HOURS}h">Overdue: ${fmt}</span>`;
    }
    // soon threshold: last 2 hours
    const soon = remain <= (2 * 60 * 60 * 1000);
    const cls = soon ? "notif-badge-pill notif-pill-sla soon" : "notif-badge-pill notif-pill-sla";
    return `<span class="${cls}" title="SLA ${SLA_HOURS}h">Due in: ${fmt}</span>`;
  } catch (e) { return ""; }
}

function uid(prefix = "n") {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getUserDirectorySmart() {
  // Reuse the smart directory approach (works even if refactors happen)
  const sources = [];

  try { if (Array.isArray(state?.users)) sources.push(state.users); } catch (_) {}
  try {
    const raw = SBTS_UTILS.LS.get("sbts_state");
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed?.users)) sources.push(parsed.users);
    }
  } catch (_) {}

  try {
    if (Array.isArray(window?.appState?.users)) sources.push(window.appState.users);
    if (Array.isArray(window?.SBTS?.users)) sources.push(window.SBTS.users);
    if (Array.isArray(window?.users)) sources.push(window.users);
  } catch (_) {}

  const seen = new Set();
  const dir = [];
  sources.flat().forEach((u) => {
    if (!u) return;
    const id = u.id || u.userId || u.uid;
    if (!id || seen.has(id)) return;
    seen.add(id);

    const fullName = u.fullName || u.name || u.username || u.displayName || "";
    const username = u.username || u.userName || "";
    const role = u.role || "";
    const profileImage = u.profileImage || u.avatar || u.photo || null;

    dir.push({ id, fullName: fullName || username || String(id), username, role, profileImage });
  });

  return dir;
}

function userDisplayNameById(userId) {
  const dir = getUserDirectorySmart();
  const u = dir.find(x => x.id === userId);
  return u?.fullName || u?.username || String(userId);
}

function addNotification(recipients, payload) {
  // Backward compatible signature:
  // 1) addNotification([userId], { ...payload })
  // 2) addNotification({ recipients:[...], ...payload })
  // 3) addNotification(userId, { ...payload })
  let rec = recipients;
  let p = payload;
  if (rec && typeof rec === "object" && !Array.isArray(rec) && p === undefined) {
    p = { ...rec };
    rec = p.recipients;
  }
  if (!Array.isArray(rec)) rec = rec ? [rec] : [];
  p = p && typeof p === "object" ? p : {};

  const n = ensureNotificationsState();
  const now = new Date().toISOString();

  const base = {
    id: uid("notif"),
    ts: now,
    read: false,
    requiresAction: !!p.requiresAction,
    resolved: !!p.resolved,
    actionKey: p.actionKey || null,
    type: p.type || "info",
    title: p.title || "Notification",
    message: p.message || "",
    projectId: p.projectId || null,
    blindId: p.blindId || null,
    phaseId: p.phaseId || null,
    fromPhase: p.fromPhase || null,
    toPhase: p.toPhase || null,
    actorId: p.actorId || null,
    createdBy: p.createdBy || (state.currentUser?.fullName || state.currentUser?.username || "System"),
    snoozedUntil: p.snoozedUntil || null,
    archived: !!p.archived,
    meta: (p && typeof p === "object") ? (p.meta || null) : null,
  };

  // recipients: list of userIds. Use "*" to broadcast to all users.
  rec.forEach((userId) => {
    if (!userId) return;
    if (userId === "*") {
      if (!Array.isArray(n.global)) n.global = [];
      n.global.unshift({ ...base, id: uid("notif") });
      n.global = n.global.slice(0, 200);
      return;
    }
    if (!n.byUser[userId]) n.byUser[userId] = [];
    n.byUser[userId].unshift({ ...base, id: uid("notif") });
    n.byUser[userId] = n.byUser[userId].slice(0, 200);
  });

  saveState();
}

function getUserApproverIds() {
  const out = [];
  (state.users || []).forEach((u) => {
    if (!u || !u.id) return;
    if (normalizeRole(u.role) === "admin") {
      out.push(u.id);
      return;
    }
    const p = state.permissions?.[u.id] || {};
    if (p.manageUsers || p.adminApproveUsers) out.push(u.id);
  });
  return Array.from(new Set(out));
}


function getAdminUserIds(){
  const users = Array.isArray(state.users) ? state.users : [];
  return users.filter(u => u && String(u.role||"").toLowerCase() === "admin" && u.id).map(u=>u.id);
}


// ------------------------------
// Notifications: user id resolver
// ------------------------------
// Some accounts may persist the signed-in user as {username/...} without a stable `id`.
// Notifications are keyed by a stable user id in state.notifications.byUser.
// This helper resolves the current user to a stable id by checking common fields and
// mapping against state.users.
function getCurrentUserIdStable() {
  const cu = state.currentUser || {};
  // Prefer explicit ids
  const direct = cu.id || cu.userId || cu.uid;
  if (direct) return direct;

  // Fallback keys
  const key = cu.username || cu.email || cu.name;
  if (!key) return null;

  const users = state.users || [];
  const found = users.find(u => u && (
    u.id === key ||
    u.userId === key ||
    u.uid === key ||
    u.username === key ||
    u.email === key ||
    u.name === key
  ));
  return found?.id || found?.userId || found?.uid || key;
}


function getNotificationsForCurrentUser() {
  ensureNotificationsState();
  applyNotifRulesToAllNotifications();
  const uid = getCurrentUserIdStable();
  if (!uid) return [];
  const personal = state.notifications.byUser[uid] || [];
  const globalList = Array.isArray(state.notifications.global) ? state.notifications.global : [];
  // Merge and sort newest first
  return personal
    .concat(globalList)
    .slice()
    .sort((a,b)=> ((b.ts ?? b.createdAt) || 0) - ((a.ts ?? a.createdAt) || 0));
}

function countPendingActionsForCurrentUser() {
  const list = getNotificationsForCurrentUser();
  return list.filter(n => n.requiresAction && !n.resolved).length;
}

function markSeenNonActionNotifications() {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  let changed = false;
  list.forEach(n => {
    if (!n.requiresAction && !n.read) {
      n.read = true;
      changed = true;
    }
  });
  if (changed) saveState();
}

function resolveNotificationsByActionKey(actionKey) {
  if (!actionKey) return;
  ensureNotificationsState();
  Object.keys(state.notifications.byUser || {}).forEach(uid => {
    const list = state.notifications.byUser[uid] || [];
    list.forEach(n => {
      if (n.actionKey === actionKey && n.requiresAction && !n.resolved) {
        n.resolved = true;
      }
    });
  });
  saveState();
}



function openNotificationsDrawer() {
  if (!state.currentUser) return;
  // smart default: if a project is open, filter to it
  const hasProject = !!state.currentProjectId;
  const nd = ensureNotifDrawerUI();
  if (nd.projectFilter === "auto") {
    nd.projectFilter = hasProject ? "current" : "all";
  }
  // keep tab sensible
  if (!nd.tab) nd.tab = "inbox";

  document.getElementById("notifDrawerOverlay")?.classList.remove("hidden");
  document.getElementById("notifDrawer")?.classList.remove("hidden");
  renderNotificationsDrawer();
  updateNotificationsBadge();
}

function closeNotificationsDrawer() {
  document.getElementById("notifDrawerOverlay")?.classList.add("hidden");
  document.getElementById("notifDrawer")?.classList.add("hidden");
}


function ensureNotifDrawerUI() {
  if (!state.ui) state.ui = {};
  if (!state.ui.notifDrawer) state.ui.notifDrawer = {};
  const nd = state.ui.notifDrawer;
  if (!nd.tab) nd.tab = "inbox";
  // Legacy tab mapping (pre-46.2)
  if (nd.tab === "action") nd.tab = "inbox";
  if (nd.tab === "all" || nd.tab === "unread") nd.tab = "updates";

  if (!nd.projectFilter) nd.projectFilter = "auto";
  if (!nd.groupBy) nd.groupBy = "project";
  if (typeof nd.hideInfo !== "boolean") nd.hideInfo = false;
  return nd;
}

function toggleNotifHideInfo() {
  const nd = ensureNotifDrawerUI();
  nd.hideInfo = !nd.hideInfo;
  saveState();
  renderNotificationsDrawer();
}

function setNotifDrawerTab(tab) {
  const nd = ensureNotifDrawerUI();
  nd.tab = tab;
  renderNotificationsDrawer();
}

function setNotifProjectFilter(mode) {
  const nd = ensureNotifDrawerUI();
  nd.projectFilter = mode;
  renderNotificationsDrawer();
}

function sbtsToIso(v){
  try{
    if(v===null || v===undefined) return new Date().toISOString();
    if(typeof v === "number") return new Date(v).toISOString();
    if(v instanceof Date) return v.toISOString();
    const d = new Date(v);
    if(!isNaN(d.getTime())) return d.toISOString();
  }catch(e){}
  return new Date().toISOString();
}

function notifRelativeTime(iso) {
  try {
    const d = new Date(iso);
    const ms = Date.now() - d.getTime();
    if (!isFinite(ms)) return "";
    const s = Math.floor(ms / 1000);
    if (s < 60) return `${s}s ago`;
    const m = Math.floor(s / 60);
    if (m < 60) return `${m}m ago`;
    const h = Math.floor(m / 60);
    if (h < 24) return `${h}h ago`;
    const days = Math.floor(h / 24);
    return `${days}d ago`;
  } catch {
    return "";
  }
}



function renderNotifDrawerInbox(box){
  ensureRequestsArray();
  const nd = ensureNotifDrawerUI();
  const mode = state.ui.notifDrawer?.projectFilter || "all";
  const curPid = state.currentProjectId;
  const q = (document.getElementById("notifDrawerSearch")?.value || "").trim().toLowerCase();

  // Selected request detail view
  const selId = nd.selectedReqId || (state.ui?.reqInboxFocusId || null);
  if (selId) {
    const req = (state.requests || []).find(r => r && r.id === selId);
    // clear one-time focus
    if (state.ui?.reqInboxFocusId) { state.ui.reqInboxFocusId = null; saveState(); }
    if (!req) { nd.selectedReqId = null; saveState(); }
  }

  let reqs = (state.requests || []).slice();

  // project filter
  if (mode === "current" && curPid) reqs = reqs.filter(r => (r.projectId || null) === curPid);

  // open only in inbox tab
  reqs = reqs.filter(r => isRequestOpen(r.status));

  // search
  if (q) {
    reqs = reqs.filter(r => {
      const hay = `${r.kind||""} ${r.phaseId||""} ${r.reason||""} ${r.projectName||""} ${formatReqContextLine(r)}`.toLowerCase();
      return hay.includes(q);
    });
  }

  // Sort newest first
  reqs.sort((a,b) => String(b.createdAt||"").localeCompare(String(a.createdAt||"")));

  // If a request is selected, show detail view
  if (nd.selectedReqId) {
    const req = (state.requests || []).find(r => r && r.id === nd.selectedReqId);
    if (!req) { nd.selectedReqId = null; saveState(); return renderNotifDrawerInbox(box); }

    const ctx = formatReqContextLine(req);
    const status = requestStatusLabel(req.status);
    const canManage = canUser("manageRequests");

    box.innerHTML = `
      <div class="notif-detail">
        <div class="notif-detail-top">
          <button class="btn btn-sm" id="notifReqBackBtn">← Back</button>
          <div class="notif-detail-title">
            <div class="title">${escapeHtml(req.kind || "Request")}</div>
            <div class="tiny">${ctx}</div>
          </div>
          <span class="badge">${escapeHtml(status)}</span>
        </div>

        <div class="notif-detail-body">
          <div class="tiny" style="margin-bottom:8px;">
            Requested by <b>${escapeHtml(req.requestedByName || "User")}</b>
            • ${escapeHtml(formatDate(req.createdAt))}
          </div>

          ${req.reason ? `<div class="card" style="padding:10px;margin-bottom:10px;">
            <div class="tiny" style="opacity:.8;margin-bottom:6px;">Reason</div>
            <div>${escapeHtml(req.reason)}</div>
          </div>` : ""}

          <div class="card" style="padding:10px;">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
              <div class="tiny" style="opacity:.85;">Comments</div>
              <div class="tiny" style="opacity:.65;">${(req.comments||[]).length}</div>
            </div>
            <div id="notifReqComments" style="margin-top:8px;"></div>
            <div style="display:flex;gap:8px;margin-top:10px;">
              <input class="input" id="notifReqCommentInput" placeholder="Write a comment…" />
              <button class="btn btn-sm" id="notifReqCommentBtn">Post</button>
            </div>
          </div>

          <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px;">
            ${canManage ? `<button class="btn btn-sm" id="notifReqInReview">Mark In Review</button>` : ""}
            ${canManage ? `<button class="btn btn-sm btn-primary" id="notifReqApprove">Approve</button>` : ""}
            ${canManage ? `<button class="btn btn-sm btn-danger" id="notifReqReject">Reject</button>` : ""}
            <button class="btn btn-sm" id="notifReqOpenTarget">Open related item</button>
            <button class="btn btn-sm" id="notifReqArchive">Archive</button>
            <button class="btn btn-sm btn-danger" id="notifReqDelete">Delete</button>
          </div>
        </div>
      </div>
    `;

    // bind
    document.getElementById("notifReqBackBtn")?.addEventListener("click", () => {
      nd.selectedReqId = null;
      saveState();
      renderNotificationsDrawer();
    });

    const cbox = document.getElementById("notifReqComments");
    if (cbox) {
      const cmts = Array.isArray(req.comments) ? req.comments : [];
      if (!cmts.length) {
        cbox.innerHTML = `<div class="tiny" style="opacity:.7;">No comments yet.</div>`;
      } else {
        cbox.innerHTML = cmts.map(c => `
          <div class="notif-comment">
            <div class="tiny"><b>${escapeHtml(c.byName||"User")}</b> • ${escapeHtml(formatDate(c.at))}</div>
            <div>${escapeHtml(c.text||"")}</div>
          </div>
        `).join("");
      }
    }

    document.getElementById("notifReqCommentBtn")?.addEventListener("click", () => {
      const inp = document.getElementById("notifReqCommentInput");
      const val = inp?.value || "";
      addRequestComment(req.id, val);
      if (inp) inp.value = "";
      renderNotificationsDrawer();
    });

    if (canManage) {
      document.getElementById("notifReqInReview")?.addEventListener("click", () => {
        setRequestStatus(req.id, "in_review");
        addRequestComment(req.id, "Status set to In Review.");
        renderNotificationsDrawer();
        updateNotificationsBadge();
      });
      document.getElementById("notifReqApprove")?.addEventListener("click", () => {
        approveRequest(req.id);
        // keep drawer view consistent
        renderNotificationsDrawer();
        updateNotificationsBadge();
      });
      document.getElementById("notifReqReject")?.addEventListener("click", () => {
        rejectRequest(req.id);
        renderNotificationsDrawer();
        updateNotificationsBadge();
      });
    }

    document.getElementById("notifReqOpenTarget")?.addEventListener("click", () => {
      closeNotificationsDrawer();
      openNotificationTarget({ blindId: req.blindId, projectId: req.projectId, requestId: req.id, actionKey: req.actionKey, meta: req.meta });
    });
document.getElementById("notifReqArchive")?.addEventListener("click", () => {
      archiveRequest(req.id);
      nd.selectedReqId = null;
      saveState();
      renderNotificationsDrawer();
      updateNotificationsBadge();
    });
    document.getElementById("notifReqDelete")?.addEventListener("click", () => {
      deleteRequest(req.id);
      nd.selectedReqId = null;
      saveState();
      renderNotificationsDrawer();
      updateNotificationsBadge();
    });



    return;
  }

  // List view
  if (!reqs.length) {
    box.innerHTML = `<div class="empty-state" style="margin:8px 0;">
      <div class="empty-icon">📥</div>
      <div class="empty-title">No inbox requests</div>
      <div class="empty-sub">Requests that need action will appear here.</div>
    </div>`;
    return;
  }

  // Group by project
  const grouped = new Map();
  reqs.forEach(r => {
    const pid = r.projectId || "__none__";
    if (!grouped.has(pid)) grouped.set(pid, []);
    grouped.get(pid).push(r);
  });

  const projects = Array.isArray(state.projects) ? state.projects : [];
  const blinds = Array.isArray(state.blinds) ? state.blinds : [];
  const projectLabel = (pid) => {
    if (pid === "__none__") return "No project";
    return projects.find(p => p.id === pid)?.name || pid;
  };

  const keys = Array.from(grouped.keys()).sort((a,b)=>projectLabel(a).localeCompare(projectLabel(b)));
  if (state.currentProjectId && mode === "all") {
    const i = keys.indexOf(state.currentProjectId);
    if (i > 0) { keys.splice(i,1); keys.unshift(state.currentProjectId); }
  }

  box.innerHTML = "";
  keys.forEach(pid => {
    const items = grouped.get(pid) || [];
    const header = document.createElement("div");
    header.className = "notif-group-header";
    header.innerHTML = `<span>${escapeHtml(projectLabel(pid))}</span><span class="notif-group-sub">${items.length}</span>`;
    box.appendChild(header);

    items.forEach(r => {
      try {
        const el = document.createElement("div");
      el.className = "notif-item";
      const ctx = formatReqContextLine(r);
      const tag = escapeHtml(requestStatusLabel(r.status));
      const reason = r.reason ? escapeHtml(r.reason) : "";
      const ccount = (r.comments||[]).length;
      el.innerHTML = `
        <div class="notif-item-main">
          <div class="notif-item-title">${escapeHtml(r.kind || "Request")} <span class="badge">${tag}</span></div>
          <div class="tiny">${ctx}${reason ? ` • ${reason}` : ""}</div>
          <div class="tiny" style="opacity:.65;margin-top:4px;">${escapeHtml(formatDate(r.createdAt))} • ${ccount} comment${ccount===1?"":"s"}</div>
        </div>
        <div class="notif-item-actions">
          <button class="btn btn-sm">Open</button>
        </div>
      `;
      el.querySelector("button")?.addEventListener("click", () => {
        nd.selectedReqId = r.id;
        saveState();
        renderNotificationsDrawer();
      });
        box.appendChild(el);
      } catch (e) {
        console.warn('NotifDrawer item render failed', e);
      }
    });
  });
}

function renderNotifDrawerInboxCombined(box) {
  // Combines Requests + Action Alerts (notifications that require action) in the Inbox tab.
  if (!box) return;
  ensureNotificationsState();

  const nd = ensureNotifDrawerUI();
  const modePF = state.ui.notifDrawer?.projectFilter || "all";
  const curPid = state.currentProjectId;

  // Build layout: requests first, then action alerts.
  box.innerHTML = `
    <div class="nd-section">
      <div class="nd-section-title">Requests</div>
      <div id="ndReqWrap"></div>
    </div>
    <div class="nd-section" style="margin-top:12px;">
      <div class="nd-section-title">Action alerts</div>
      <div class="tiny muted" style="margin-top:4px;">System items that require your action (approvals, assignments, urgent notes).</div>
      <div id="ndActionWrap" style="margin-top:8px;"></div>
    </div>
    <div class="nd-section" style="margin-top:12px;">
      <div class="nd-section-title">Updates</div>
      <div class="tiny muted" style="margin-top:4px;">Recent updates and info notifications (non-action).</div>
      <div id="ndUpdatesWrap" style="margin-top:8px;"></div>
    </div>
  `;

  // Render Requests (existing engine) into its own container.
  const reqWrap = document.getElementById("ndReqWrap");
  if (reqWrap) {
    renderNotifDrawerInbox(reqWrap);
  }

  // Render Action Alerts (notifications with requiresAction)
  const aWrap = document.getElementById("ndActionWrap");
  if (!aWrap) return;

  let list = getNotificationsForCurrentUser().filter(n => n && n.requiresAction && !n.resolved);

  if (modePF === "current" && curPid) {
    list = list.filter(n => (n.projectId || null) === curPid);
  }

  // optional: hideInfo only affects non-action items, but keep safe
  list = list.filter(n => n && typeof n === "object");

  // search (reuse drawer search box)
  const q = (document.getElementById("notifDrawerSearch")?.value || "").trim().toLowerCase();
  if (q) {
    const projects = Array.isArray(state.projects) ? state.projects : [];
    const blinds = Array.isArray(state.blinds) ? state.blinds : [];
    list = list.filter(n => {
      const project = n.projectId ? (projects.find(p => p.id === n.projectId)?.name || "") : "";
      const blind = n.blindId ? (blinds.find(b => b.id === n.blindId)?.name || "") : "";
      const hay = `${n.title} ${n.message} ${project} ${blind}`.toLowerCase();
      return hay.includes(q);
    });
  }

  if (!list.length) {
    aWrap.innerHTML = `<div class="empty-state" style="margin:8px 0;">
      <div class="empty-icon">✅</div>
      <div class="empty-title">No action alerts</div>
      <div class="empty-sub">You're all caught up.</div>
    </div>`;
    return;
  }

  // Group by project, but label null as "System"
  const grouped = new Map();
  list.forEach(n => {
    const pid = n.projectId || "__system__";
    if (!grouped.has(pid)) grouped.set(pid, []);
    grouped.get(pid).push(n);
  });

  const projects = Array.isArray(state.projects) ? state.projects : [];
  const projectLabel = (pid) => {
    if (pid === "__system__") return "System";
    return projects.find(p => p.id === pid)?.name || pid;
  };

  const keys = Array.from(grouped.keys()).sort((a,b)=>projectLabel(a).localeCompare(projectLabel(b)));
  if (state.currentProjectId && modePF === "all") {
    const i = keys.indexOf(state.currentProjectId);
    if (i > 0) { keys.splice(i,1); keys.unshift(state.currentProjectId); }
  }

  aWrap.innerHTML = "";
  keys.forEach(pid => {
    const items = grouped.get(pid) || [];
    const header = document.createElement("div");
    header.className = "notif-group-header";
    header.innerHTML = `<span>${escapeHtml(projectLabel(pid))}</span><span class="notif-group-sub">${items.length}</span>`;
    aWrap.appendChild(header);

    items.forEach(n => {
      const el = document.createElement("div");
      el.className = "notif-item";
      const when = escapeHtml(notifRelativeTime(sbtsToIso(n.createdAt || n.ts || Date.now())));
      const title = escapeHtml(n.title || "Action required");
      const msg = escapeHtml(n.message || "");
      el.innerHTML = `
        <div class="notif-item-main">
          <div class="notif-item-title">${title} <span class="badge badge-warn">Action</span></div>
          <div class="tiny">${msg}</div>
          <div class="tiny" style="opacity:.65;margin-top:4px;">${when}</div>
        </div>
        <div class="notif-item-actions">
          <button class="btn btn-sm">Open</button>
          <button class="btn btn-sm btn-ghost">Done</button>
        </div>
      `;
      const btns = el.querySelectorAll("button");
      btns[0]?.addEventListener("click", () => openNotificationTarget(n.id));
      btns[1]?.addEventListener("click", () => resolveNotification(n.id));
      aWrap.appendChild(el);
    });
  });

// Updates (non-action, not archived)
const uWrap = document.getElementById("ndUpdatesWrap");
if (uWrap){
  let up = getNotificationsForCurrentUser().filter(n => n && !n.archived && !n.requiresAction && !n.resolved);
  if (modePF === "current" && curPid) up = up.filter(n => (n.projectId || null) === curPid);
  up.sort((a,b)=> new Date(b.ts||b.createdAt||0).getTime() - new Date(a.ts||a.createdAt||0).getTime());
  up = up.slice(0,20);

  if (!up.length){
    uWrap.innerHTML = `<div class="empty-state" style="margin:8px 0;">
      <div class="empty-icon">📬</div>
      <div class="empty-title">No updates</div>
      <div class="empty-sub">No new updates right now.</div>
    </div>`;
  } else {
    uWrap.innerHTML = "";
    up.forEach(n => {
      const el = document.createElement("div");
      el.className = "notif-item";
      const when = escapeHtml(notifRelativeTime(sbtsToIso(n.createdAt || n.ts || Date.now())));
      const title = escapeHtml(n.title || "Update");
      const msg = escapeHtml((n.message || "").toString().slice(0,140));
      el.innerHTML = `
        <div class="notif-item-main">
          <div class="notif-item-title">${title} <span class="badge badge-info">Update</span></div>
          <div class="tiny">${msg}</div>
          <div class="tiny" style="opacity:.65;margin-top:4px;">${when}</div>
        </div>
        <div class="notif-item-actions">
          <button class="btn btn-sm">Open</button>
        </div>
      `;
      el.querySelector("button")?.addEventListener("click", () => openNotificationTarget(n.id));
      el.addEventListener("click", (ev) => {
        if (ev.target && ev.target.tagName === "BUTTON") return;
        try{ openInboxReader(n.id); }catch(e){}
      });
      uWrap.appendChild(el);
    });
  }
}


}


function getNotifDrawerFilteredList() {
  const all = getNotificationsForCurrentUser();
  const tab = state.ui.notifDrawer?.tab || "inbox";
  const mode = state.ui.notifDrawer?.projectFilter || "all";
  const currentPid = state.currentProjectId;

  let list = all;

  // project filter
  if (mode === "current" && currentPid) {
    list = list.filter(n => (n.projectId || null) === currentPid);
  }

  // tab filter
  if (tab === "inbox") {
    list = list.filter(n => n.requiresAction && !n.resolved);
  } else if (tab === "updates") {
    list = list.filter(n => !n.requiresAction && !n.resolved);
  } else if (tab === "done") {
    list = list.filter(n => !!n.resolved);
  }

// search// search
  const q = (document.getElementById("notifDrawerSearch")?.value || "").trim().toLowerCase();
  if (q) {
    list = list.filter(n => {
      const projects = Array.isArray(state.projects) ? state.projects : [];
      const blinds = Array.isArray(state.blinds) ? state.blinds : [];
      const project = n.projectId ? (projects.find(p => p.id === n.projectId)?.name || "") : "";
      const blind = n.blindId ? (blinds.find(b => b.id === n.blindId)?.name || "") : "";
      const hay = `${n.title} ${n.message} ${project} ${blind}`.toLowerCase();
      return hay.includes(q);
    });
  }

  return list;
}


/* ==========================
   Inbox/Notification navigation helpers (Patch 47.14)
   - Focus a specific slip blind row from Inbox/Notifications
========================== */
function focusBlindRow(blindId){
  try{
    if(!blindId) return;
    const b = (state.blinds||[]).find(x=>x.id===blindId);
    if(!b){
      showToast("Blind not found: " + blindId, "warn");
      return;
    }
    // Ensure slip selection aligns with the blind's area & project
    state.slip = state.slip || {};
    state.slip.areaId = b.areaId || state.slip.areaId || "";
    state.slip.projectId = b.projectId || state.slip.projectId || "";
    // Render page (safe)
    try{ renderSlipBlindPage(); }catch(e){}
    // Focus row after DOM paint
    setTimeout(()=>{
      try{
        const row = document.getElementById("slip_row_"+blindId) 
          || document.querySelector(`#slipBlindsTableBody tr[data-blind-id="${blindId}"]`);
        if(!row) return;
        // clear previous focus
        document.querySelectorAll("#slipBlindsTableBody tr.row-focus").forEach(r=>r.classList.remove("row-focus"));
        row.classList.add("row-focus");
        row.scrollIntoView({behavior:"smooth", block:"center"});
      }catch(e){}
    }, 80);
  }catch(e){
    console.warn("[focusBlindRow] failed", e);
  }
}

function openNotificationTarget(notifOrId){
  try{
    // Accept either a notification id (string/number) or a notification object.
    const n = (typeof notifOrId === "object" && notifOrId) ? notifOrId : getNotificationById(notifOrId);
    if(!n) return;

    // Mark as read when user opens the target.
    markNotificationRead(n.id);

    // Prefer explicit actionKey routing.
    if(n.actionKey){
      if(routeByActionKey(n.actionKey, n)) return;
      // if actionKey exists but not routable, continue to id-based routing
    }

    // ID-based routing (preferred)
    const userId = (n.userId) || (n.meta && n.meta.userId) || (String(n.actionKey||"").startsWith("user:") ? String(n.actionKey).split(":")[1] : null);
    if(userId){
      openUserManagement(userId);
      return;
    }

    const blindId = n.blindId || (n.meta && n.meta.blindId);
    if(blindId){
      // Go to Slip Blind list and focus the blind (best UX)
      openPage("slipBlindPage");
      try{ renderSlipBlindPage(); }catch(e){}
      try{ focusBlindRow(blindId); }catch(e){}
      return;
    }

    const projectId = n.projectId || (n.meta && n.meta.projectId);
    if(projectId){
      // Go to Projects page and open the project
      try{
        if(typeof openProjectDetails === "function") openProjectDetails(projectId);
        else openProjectFromNotif(projectId);
      }catch(e){
        openProjectFromNotif(projectId);
      }
      return;
    }

    // Fallback based on meta (legacy)
    const meta = n.meta || {};
    if(meta.projectId){
      openProjectDetails(meta.projectId);
      return;
    }
    if(meta.userId){
      openPage("usersPage");
      if(typeof focusUserRow === "function") focusUserRow(meta.userId);
      return;
    }
    if(meta.blindId){
      openPage("slipBlindPage");
      if(typeof focusBlindRow === "function") focusBlindRow(meta.blindId);
      return;
    }

    // Last resort: open inbox
    openPage("notificationsPage");
  }catch(e){
    console.warn("[openNotificationTarget] failed", e);
    openPage("notificationsPage");
  }
}

function openProjectSettingsForProject(projectId) {
  if (!projectId) return;
  try {
    openProjectDetails(projectId);
    // Ensure settings tab is visible
    setTimeout(() => {
      try { openProjectSubTab("settings"); } catch (_) {}
    }, 0);
  } catch (e) {
    // fallback
    openProjectDetails(projectId);
  }
}

function resolveNotification(notifId) {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;
  n.resolved = true;
  n.read = true;
  n.resolvedTs = Date.now();
  const u = (state.users || []).find(x => x.id === uid);
  const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";
  n.activity = Array.isArray(n.activity) ? n.activity : [];
  n.activity.push({ ts: Date.now(), action: "Action marked done", by: byName });
saveState();
  renderNotificationsDrawer();
  renderNotificationsInbox();
  updateNotificationsBadge();
}

function clearResolvedNotifications() {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  state.notifications.byUser[uid] = list.filter(n => !n.resolved);
  saveState();
  renderNotificationsDrawer();
  renderNotificationsInbox();
  updateNotificationsBadge();
}

function renderNotificationsDrawer() {
  try {

  const nd = ensureNotifDrawerUI();
  const hasProject = !!state.currentProjectId;
  const projectName = hasProject ? (state.projects.find(p => p.id === state.currentProjectId)?.name || "") : "";

  // Auto mode fallback (if user never touched filter)
  if (state.ui.notifDrawer?.projectFilter === "auto") {
    nd.projectFilter = hasProject ? "current" : "all";
  }

  const mode = state.ui.notifDrawer?.projectFilter || "all";
  const hideInfo = !!state.ui.notifDrawer?.hideInfo;
  const tab = state.ui.notifDrawer?.tab || "inbox";

  // hide current project chip if none
  const currentChip = document.getElementById("notifPF_current");
  if (currentChip) {
    currentChip.style.display = hasProject ? "inline-flex" : "none";
    currentChip.textContent = hasProject ? `Current: ${projectName}` : "Current project";
  }

  // active classes
  ["inbox","updates","done"].forEach(t => {
    const b = document.getElementById(`notifDT_${t}`);
    if (b) b.classList.toggle("active", tab === t);
  });
  const pfAll = document.getElementById("notifPF_all");
  if (pfAll) pfAll.classList.toggle("active", mode === "all");
  if (currentChip) currentChip.classList.toggle("active", mode === "current");
  const hideBtn = document.getElementById('notifHideInfoBtn');
  if (hideBtn) hideBtn.classList.toggle('active', !!state.ui.notifDrawer?.hideInfo);

  const sub = document.getElementById("notifDrawerSub");
  if (sub) {
    const base = (mode === "current" && hasProject)
      ? `Center • filtered to: ${projectName}`
      : "Center • all projects";
    const tabNow = state.ui.notifDrawer?.tab || "inbox";
    // Inbox gets a "New request" quick action
    if (tabNow === "inbox") {
      sub.innerHTML = `<div class="drawer-sub-row">
        <div>${escapeHtml(base)}</div>
        <div class="drawer-sub-actions">
          <span class="tiny muted">Create requests from Projects / Users pages</span>
        </div>
      </div>`;
    } else {
      sub.textContent = base;
    }
  }

  const box = document.getElementById("notifDrawerList");
  if (!box) return;

// Update tab counts (always, including when Inbox tab is active)
  try {
    const allN = getNotificationsForCurrentUser();
    const modePF = state.ui.notifDrawer?.projectFilter || "all";
    const curPid = state.currentProjectId;
    let base = allN;
    if (modePF === "current" && curPid) base = base.filter(n => (n.projectId || null) === curPid);

    const openReqs = (
      Array.isArray(state.notifications?.requests) ? state.notifications.requests : []
    ).filter(r => r && r.status !== "done" && r.status !== "archived");

    const openReqsFiltered = (modePF === "current" && curPid)
      ? openReqs.filter(r => (r.projectId || null) === curPid)
      : openReqs;

    const inboxCount = base.filter(n => n.requiresAction && !n.resolved).length + openReqsFiltered.length;
    const updatesCount = base.filter(n => !n.requiresAction && !n.resolved).length;
    const doneCount = base.filter(n => !!n.resolved).length;

    const btnInbox = document.getElementById("notifDrawerTabInbox");
    const btnUpdates = document.getElementById("notifDrawerTabUpdates");
    const btnDone = document.getElementById("notifDrawerTabDone");
    if (btnInbox) btnInbox.querySelector(".pill")?.replaceChildren(document.createTextNode(String(inboxCount)));
    if (btnUpdates) btnUpdates.querySelector(".pill")?.replaceChildren(document.createTextNode(String(updatesCount)));
    if (btnDone) btnDone.querySelector(".pill")?.replaceChildren(document.createTextNode(String(doneCount)));
  } catch (e) {
    // ignore
  }

  // INBOX TAB: show combined Requests + Action Alerts
  if ((state.ui.notifDrawer?.tab || "inbox") === "inbox") {
    renderNotifDrawerInboxCombined(box);
    updateNotificationsBadge();
    return;
  }

  const list = getNotifDrawerFilteredList();

  if (!list.length) {
    box.innerHTML = `<div class="tiny" style="padding:10px;">No notifications.</div>`;
    return;
  }

  
  box.innerHTML = "";

  // Group by project (default)
  const grouped = new Map();
  list.forEach((n) => {
    const pid = n.projectId || "__none__";
    if (!grouped.has(pid)) grouped.set(pid, []);
    grouped.get(pid).push(n);
  });

  const projects = Array.isArray(state.projects) ? state.projects : [];
  const blinds = Array.isArray(state.blinds) ? state.blinds : [];
  const projectLabel = (pid) => {
    if (pid === "__none__") return "No project";
    return projects.find(p => p.id === pid)?.name || pid;
  };

  const groupKeys = Array.from(grouped.keys()).sort((a, b) => projectLabel(a).localeCompare(projectLabel(b)));
  // Put current project first if filtering to all
  if (state.currentProjectId && mode === 'all') {
    const i = groupKeys.indexOf(state.currentProjectId);
    if (i > 0) {
      groupKeys.splice(i, 1);
      groupKeys.unshift(state.currentProjectId);
    }
  }

  groupKeys.forEach((pid) => {
    const items = grouped.get(pid) || [];
    const header = document.createElement('div');
    header.className = 'notif-group-header';
    header.innerHTML = `<span>${escapeHtml(projectLabel(pid))}</span><span class="notif-group-sub">${items.length}</span>`;
    box.appendChild(header);

    items.forEach((n) => {
      const project = n.projectId
        ? (state.projects.find(p => p.id === n.projectId)?.name || n.projectId)
        : (n.projectName || "-");
      const blind = n.blindId ? (state.blinds.find(b => b.id === n.blindId)?.name || n.blindId) : null;

      const pill = n.resolved ? { cls: "notif-pill-done", label: "Done" }
        : n.requiresAction ? { cls: "notif-pill-action", label: "Inbox" }
        : { cls: "notif-pill-info", label: "Update" };

      const wrap = document.createElement("div");
      wrap.className = "notif-item";
      wrap.innerHTML = `
        <div class="row">
          <div style="min-width:0;">
            <div style="font-weight:800; line-height:1.2;">${escapeHtml(n.title || "Notification")}</div>
            <div class="tiny" style="margin-top:4px;">${escapeHtml(n.message || "")}</div>
          </div>
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="notif-badge-pill ${pill.cls}">${pill.label}</span>
            ${notifSlaPill(n)}
          </div>
        </div>
        <div class="notif-meta">${getNotifContextLinks(n)}<span><i class="ph ph-clock"></i> ${escapeHtml(notifRelativeTime(sbtsToIso(n.ts||n.createdAt||Date.now())))}</span></div>
        <div class="notif-actions">
          <button class="secondary-btn tiny" onclick="openNotificationTarget('${n.id}')"><i class="ph ph-arrow-square-out"></i> Open</button>
          ${n.resolved ? "" : (n.requiresAction ? `<button class="primary-btn tiny" onclick="resolveNotification('${n.id}')"><i class="ph ph-check"></i> Mark done</button>` : `<button class="btn-neutral tiny" onclick="toggleNotificationRead('${n.id}')"><i class="ph ph-eye"></i> ${n.read ? "Mark unread" : "Mark read"}</button>`)}
          <button class="btn-neutral tiny" onclick="deleteNotification('${n.id}'); renderNotificationsDrawer();"><i class="ph ph-trash"></i> Delete</button>
        </div>
      `;
      if (!n.read) wrap.style.borderColor = "rgba(10,103,163,0.35)";
      box.appendChild(wrap);
    });
  });


  } catch (e) {
    showToast('Notifications failed to render: ' + (e?.message || e), 'error');
    console.error(e);
  }
}

// Legacy modal kept for compatibility (not used by bell anymore)
function openNotificationsModal() {
  if (!state.currentUser) return;
  renderNotificationsList();
  openModal("notificationsModal");
  updateNotificationsBadge();
}

function markNotificationRead(notifId, readValue = true) {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;

  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (n) n.read = !!readValue;

  saveState();
  renderNotificationsList();
  renderNotificationsInbox();
  updateNotificationsBadge();
}





function markNotificationUnread(notifId) {
  // Mark a notification as unread (local-only)
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;

  const u = (state.users || []).find(x => x.id === uid);
  const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";

  n.read = false;
  n.unreadTs = Date.now();
  n.activity = Array.isArray(n.activity) ? n.activity : [];
  n.activity.push({ ts: Date.now(), action: "Marked unread", by: byName });

  saveState();
  try { renderNotificationsList(); } catch (e) {}
  try { renderNotificationsInbox(); } catch (e) {}
  updateNotificationsBadge();
}

function toggleNotificationRead(notifId) {
  // Toggle read/unread, used by Details panel and legacy buttons
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;

  if (n.read) {
    markNotificationUnread(notifId);
  } else {
    // mark read
    const u = (state.users || []).find(x => x.id === uid);
    const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";
    n.read = true;
    n.readTs = Date.now();
    n.activity = Array.isArray(n.activity) ? n.activity : [];
    n.activity.push({ ts: Date.now(), action: "Marked read", by: byName });
    saveState();
    try { renderNotificationsList(); } catch (e) {}
    try { renderNotificationsInbox(); } catch (e) {}
    updateNotificationsBadge();
  }
}



function openNotifPreviewFromTitle(notifId){
  // Title click opens the reader view (Gmail-like) without leaving the page
  openInboxReader(notifId);
}

function closeNotifPreview(){
  // Backwards-compat: close reader if open
  try{ closeInboxReader(); }catch(_){ }
}

function openNotifModalFromList(notifId){
  // Open the message reader (Gmail-like)
  openInboxReader(notifId);
}

function deleteNotification(notifId) {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const idx = list.findIndex(x => x.id === notifId);
  if (idx >= 0) list.splice(idx, 1);
  saveState();
  renderNotificationsList();
  renderNotificationsInbox();
  updateNotificationsBadge();
}


function markAllNotificationsRead() {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;

  (state.notifications.byUser[uid] || []).forEach(n => n.read = true);
  saveState();
  renderNotificationsList();
  renderNotificationsInbox();
  updateNotificationsBadge();
}


function renderNotifQuickActions(n) {
  if (!n || !n.requiresAction || n.resolved) return "";
  // User approval (pending registration)
  if (n.actionKey && n.actionKey.startsWith("user_approval:")) {
    const userId = n.actionKey.split(":")[1];
    if (!canUser("adminApproveUsers")) return `<span class="notif-pill">Action required</span>`;
    return `
      <button class="btn-secondary btn-small" onclick="event.stopPropagation(); approveUser('${"${userId}"}');">Approve</button>
      <button class="btn-secondary btn-small" onclick="event.stopPropagation(); rejectUser('${"${userId}"}');">Reject</button>
    `.replaceAll("${userId}", userId);
  }
  return `<span class="notif-pill">Action required</span>`;
}

function renderNotificationsList() {
  const box = document.getElementById("notificationsList");
  if (!box) return;

  const list = getNotificationsForCurrentUser();
  if (list.length === 0) {
    box.innerHTML = `<div class="empty-state">No notifications yet.</div>`;
    return;
  }

  const fmt = (iso) => {
    try {
      const d = new Date(iso);
      return d.toLocaleString();
    } catch (_) { return iso; }
  };

  box.innerHTML = list.map(n => {
    const quick = renderNotifQuickActions(n);

    const unread = !n.read ? "unread" : "";
    const title = escapeHtml(n.title || "");
    const msg = escapeHtml(n.message || "");
    const time = escapeHtml(fmt(n.ts));
    return `
      <div class="notif-item ${unread}">
        <div class="notif-main" onclick="onNotificationOpen('${n.id}','${n.blindId || ""}')">
          <div class="notif-title-row">
            <div class="notif-title">${title}</div>
            <div class="notif-time">${time}</div>
          </div>
          <div class="notif-message">${msg}</div>
        </div>
        <div class="notif-actions">
          ${quick}
          ${!n.read ? `<button class="btn-secondary btn-small" onclick="event.stopPropagation(); markNotificationRead('${n.id}')">Mark read</button>` : ""}
        </div>
      </div>
    `;
  }).join("");
}

// ===== Notifications Inbox Page =====
state.ui = state.ui || {};
if (!state.ui.notifInboxTab) state.ui.notifInboxTab = "all";






function onNotificationOpen(notifId, blindId) {
  markNotificationRead(notifId);
  if (blindId) {
    closeModal("notificationsModal");
    openBlindDetails(blindId);
  }
}

// Trigger: Phase changed
function notifyOwnersOnPhaseChange(blind, fromPhase, toPhase) {
  const projectId = blind?.projectId || state.currentProjectId;
  const project = state.projects.find(p => p.id === projectId);
  if (!project) return;

  // Owners of the *target* phase
  const owners = (project.phaseOwners?.phases?.[toPhase] || []).slice();

  // If no owners set, do nothing (clean behavior)
  if (owners.length === 0) return;

  const actor = state.currentUser?.id || null;
  const recipients = owners.filter(uid => uid && uid !== actor);

  if (recipients.length === 0) return;

  const projectName = project.name || project.projectName || "Project";
  const blindTag = blind.tagNo || blind.tag || blind.blindTag || blind.id;

SBTS_ACTIVITY.pushNotification({
  category: "action",
  scope: "user",
  recipients,
  projectId,
  blindId: blind.id,
  phaseId: toPhase,
  fromPhase,
  toPhase,
  title: `Action required: ${phaseLabel(toPhase)}`,
  message: `${sbtsBlindOfficialText(blind)} — ${fromPhase ? phaseLabel(fromPhase) : "-"} → ${toPhase ? phaseLabel(toPhase) : "-"}`,
  requiresAction: true,
  resolved: false,
  actionKey: `phase:${blind.id}:${toPhase}`,
  actorId: actor
});


  // Optional: toast for current user (if they are also an owner, they won't get a notification)
  // We'll keep UI calm (no extra spam).
}


function openProjectSubTab(tab) {
  const blindsBox = document.getElementById("projectSubTabBlinds");
  const settingsBox = document.getElementById("projectSubTabSettings");
  const btnBlinds = document.getElementById("projTabBlindsBtn");
  const btnSettings = document.getElementById("projTabSettingsBtn");
  if (!blindsBox || !settingsBox) return;

  const isSettings = tab === "settings";
  blindsBox.classList.toggle("hidden", isSettings);
  settingsBox.classList.toggle("hidden", !isSettings);
  btnBlinds?.classList.toggle("active", !isSettings);
  btnSettings?.classList.toggle("active", isSettings);

  state.ui = state.ui || {};
  state.ui.projectSubTab = isSettings ? "settings" : "blinds";
  try { if (!NAV_SUPPRESS_PUSH) navReplace(navEntryForCurrent("projectDetailsPage")); } catch (e) {}

  if (isSettings) { renderProjectSettingsSummary(); updateProjectSettingsActionsVisibility(); }
}

function updateProjectSettingsActionsVisibility() {
  const row = document.getElementById("psActionsRow");
  if (!row) return;
  // Only show manage actions to authorized users.
  const allowed = canUser("managePhaseOwnership");
  row.style.display = allowed ? "flex" : "none";
}


function psGetInitials(name) {
  const s = String(name || "").trim();
  if (!s) return "U";
  const parts = s.split(/\s+/).filter(Boolean);
  const a = parts[0]?.[0] || "U";
  const b = parts.length > 1 ? parts[parts.length - 1][0] : (parts[0]?.[1] || "");
  return (a + b).toUpperCase();
}

// --- Smart binding for Registered Users (future-proof) ---
// We store selections as USER IDS in phaseOwners.phases arrays.
function psGetRegisteredUserDirectory() {
  const sources = [];

  if (Array.isArray(state?.users)) sources.push(state.users);

  // Fallback: read raw storage directly (in case state.users is not hydrated yet)
  try {
    const raw = SBTS_UTILS.LS.get("sbts_state");
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed?.users)) sources.push(parsed.users);
    }
  } catch (_) {}

  // Extra fallbacks (future connectors / refactors)
  try {
    if (Array.isArray(window?.appState?.users)) sources.push(window.appState.users);
    if (Array.isArray(window?.SBTS?.users)) sources.push(window.SBTS.users);
    if (Array.isArray(window?.users)) sources.push(window.users);
  } catch (_) {}

  const seen = new Set();
  const dir = [];

  sources.flat().forEach((u) => {
    if (!u || typeof u !== "object") return;
    const id = String(u.id || u.userId || u.uid || u.username || u.email || "").trim();
    if (!id) return;
    if (seen.has(id)) return;
    seen.add(id);

    const label = String(u.fullName || u.name || u.displayName || u.username || u.email || id).trim();
    dir.push({
      id,
      label,
      username: u.username || "",
      role: u.role || "",
      profileImage: u.profileImage || u.avatar || u.photo || null,
      _raw: u,
    });
  });

  dir.sort((a, b) => a.label.localeCompare(b.label));
  return dir;
}

function psResolveToUserId(value, dir) {
  if (!value) return null;
  const v = String(value).trim();
  const byId = dir.find((u) => u.id === v);
  if (byId) return byId.id;
  const low = v.toLowerCase();
  const byLabel = dir.find((u) => (u.label || "").toLowerCase() === low);
  if (byLabel) return byLabel.id;
  const byUsername = dir.find((u) => (u.username || "").toLowerCase() === low);
  if (byUsername) return byUsername.id;
  const byContains = dir.find((u) => (u.label || "").toLowerCase().includes(low));
  return byContains ? byContains.id : null;
}

function psNormalizeIdList(list, dir) {
  if (!Array.isArray(list)) return [];
  const out = [];
  list.forEach((x) => {
    const id = psResolveToUserId(x, dir);
    if (id && !out.includes(id)) out.push(id);
  });
  return out;
}

function psUserById(id, dir) {
  return dir.find((x) => x.id === id) || null;
}

function renderProjectSettingsSummary() {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  // Permission UI: Only authorized users can manage phase ownership.
  const canManage = canUser('managePhaseOwnership');
  const actionsRow = document.getElementById('psActionsRow');
  const reqRow = document.getElementById('psRequestsRow');
  if (actionsRow) actionsRow.style.display = canManage ? 'flex' : 'none';
  if (reqRow) reqRow.style.display = canManage ? 'none' : 'flex';

  // Set Mode radios
  const mode = project.phaseOwners.mode || "advanced";
  document.querySelectorAll('input[name="psAccessMode"]').forEach((r) => {
    r.checked = r.value === mode;
  });


  // Set Update Policy radios
  const policy = project.phaseOwners.policy || "hybrid";
  document.querySelectorAll('input[name="psUpdatePolicy"]').forEach((r) => {
    r.checked = r.value === policy;
  });

  // Render summary rows
  const box = document.getElementById("psPhaseSummary");
  if (!box) return;
  box.innerHTML = "";

  const dir = psGetRegisteredUserDirectory();
  const phaseIds = wfProjectPhaseIds(state.currentProjectId, { includeInactive: false });

  const blinds = state.blinds.filter((b) => b.projectId === project.id);

  phaseIds.forEach((ph) => {
    const row = document.createElement("div");
    row.className = "ps-phase-row";

    const count = blinds.filter((b) => (b.phase || "broken") === ph).length;
    const ownersRaw = project.phaseOwners.phases?.[ph] || [];
    const owners = psNormalizeIdList(ownersRaw, dir);
    project.phaseOwners.phases[ph] = owners;

    const supRaw = project.phaseOwners.support?.[ph] || [];
    const support = psNormalizeIdList(supRaw, dir);
    project.phaseOwners.support[ph] = support;

    const avatarsHtml = owners
      .map((uid) => {
        const u = psUserById(uid, dir);
        const name = u ? u.label : String(uid);
        if (u && u.profileImage) {
          return `<span class="ps-avatar"><img class="ps-avatar-img" src="${u.profileImage}" alt=""/><span>${escapeHtml(name)}</span></span>`;
        }
        return `<span class="ps-avatar"><span class="ps-avatar-initials">${escapeHtml(psGetInitials(name))}</span><span>${escapeHtml(name)}</span></span>`;
      })
      .join("");

    const supportHtml = support
      .map((uid) => {
        const u = psUserById(uid, dir);
        const name = u ? u.label : String(uid);
        if (u && u.profileImage) {
          return `<span class="ps-avatar ps-avatar--support"><img class="ps-avatar-img" src="${u.profileImage}" alt=""/><span>${escapeHtml(name)}</span></span>`;
        }
        return `<span class="ps-avatar ps-avatar--support"><span class="ps-avatar-initials">${escapeHtml(psGetInitials(name))}</span><span>${escapeHtml(name)}</span></span>`;
      })
      .join("");

    row.innerHTML = `
      <div class="ps-phase-left">
        <div>
          <div class="ps-phase-name">${escapeHtml(phaseLabel(ph))}</div>
          <div class="ps-phase-meta">${count} Blinds</div>
        </div>
      </div>
      <div class="ps-avatars">
        ${avatarsHtml || '<span class="muted">No owners</span>'}
        ${supportHtml ? `<div class="ps-support-line"><span class="badge">Support</span> ${supportHtml}</div>` : ''}
      </div>
    `;

    box.appendChild(row);
  });

  saveState();

  renderProjectPhaseActivationList();
}

function renderPhaseActivationTableInPSModal(){
  const tbody = document.getElementById("phaseActTbody");
  if (!tbody) return;
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  const q = (document.getElementById("phaseActSearch")?.value || "").trim().toLowerCase();

  const blinds = state.blinds.filter(b => b.projectId === project.id);
  const counts = {};
  blinds.forEach(b => {
    const ph = normalizePhaseId(b.phase || "broken");
    counts[ph] = (counts[ph] || 0) + 1;
  });

  const all = wfPhaseIds({ includeInactive: true });

  tbody.innerHTML = "";
  all
    .filter(ph => !q || phaseLabel(ph).toLowerCase().includes(q) || normalizePhaseId(ph).includes(q))
    .forEach(ph => {
      const enabled = isPhaseEnabledForProject(project.id, ph);
      const c = counts[normalizePhaseId(ph)] || 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>
          <div class="row" style="gap:10px; align-items:center;">
            <span class="phase-dot" style="background:${wfPhaseColor(ph)}"></span>
            <div style="display:flex;flex-direction:column;gap:2px;min-width:0">
              <b>${escapeHtml(phaseLabel(ph))}</b>
              <span class="tiny muted">${escapeHtml(normalizePhaseId(ph))}</span>
            </div>
          </div>
        </td>
        <td>${c}</td>
        <td>${enabled ? '<span class="badge badge-green">active</span>' : '<span class="badge badge-red">inactive</span>'}</td>
        <td style="text-align:center">
          <input type="checkbox" ${enabled ? "checked" : ""} onchange="phaseActToggleOne('${escapeJs(ph)}', this.checked)">
        </td>
      `;
      tbody.appendChild(tr);
    });

  if (!tbody.children.length){
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4" class="muted">No phases found.</td>`;
    tbody.appendChild(tr);
  }
  updatePhaseActivationSummaryUI();
}

function phaseActToggleOne(phaseId, enabled){
  setPhaseEnabledForProject(state.currentProjectId, phaseId, enabled);
  updatePhaseActivationSummaryUI();
}

function phaseActSetAll(enabled){
  const all = wfPhaseIds({ includeInactive: true });
  all.forEach(ph => setPhaseEnabledForProject(state.currentProjectId, ph, enabled));
  renderPhaseActivationTableInPSModal();
  updatePhaseActivationSummaryUI();
}

function psToggleProjectPhase(phaseId, enabled) {
  setPhaseEnabledForProject(state.currentProjectId, phaseId, enabled);
  renderProjectPhaseCards();
  renderProjectBlindsTable();
  renderProjectSettingsSummary();
}


function psSetProjectMode(mode) {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;
  project.phaseOwners.mode = mode;
  saveState();
  toast("Mode saved ✅");
}

function psSetProjectPolicy(policy) {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;
  project.phaseOwners.policy = policy;
  saveState();
  toast("Update policy saved ✅");
}

function getProjectUpdatePolicy(projectId) {
  const project = projectId ? getProjectByIdSafe(projectId) : ensureCurrentProjectPhaseOwners();
  const pol = project?.phaseOwners?.policy || "hybrid";
  const p = String(pol).toLowerCase().trim();
  if (p === "owners" || p === "owner" || p === "owners_only") return "owners";
  if (p === "roles" || p === "role" || p === "roles_only") return "roles";
  return "hybrid";
}

function getAllowedRolesForPhase(phaseId) {
  const cfg = getWorkflowConfig();
  const target = cfg.phases.find((p) => normalizePhaseId(p.id) === normalizePhaseId(phaseId));
  const allowedRoles = Array.isArray(target?.canUpdate) ? target.canUpdate : [];
  return allowedRoles.filter(Boolean);
}


/* ==========================
   Phase Ownership Modal
========================== */

function psSwitchTab(tab){
  const ownersBtn = document.getElementById("psTabOwners");
  const actBtn = document.getElementById("psTabActivation");
  const ownersPanel = document.getElementById("psTabPanelOwners");
  const actPanel = document.getElementById("psTabPanelActivation");
  if (!ownersBtn || !actBtn || !ownersPanel || !actPanel) return;

  if (tab === "activation"){
    ownersBtn.classList.remove("active");
    actBtn.classList.add("active");
    ownersPanel.style.display = "none";
    actPanel.style.display = "block";
    renderPhaseActivationTableInPSModal();
    const s = document.getElementById("phaseActSearch");
    if (s && !s._bound){
      s._bound = true;
      s.addEventListener("input", renderPhaseActivationTableInPSModal);
    }
  } else {
    actBtn.classList.remove("active");
    ownersBtn.classList.add("active");
    actPanel.style.display = "none";
    ownersPanel.style.display = "block";
  }
}

function openPhaseActivationTab(){
  openPhaseOwnershipModal();
  setTimeout(()=>psSwitchTab("activation"), 0);
}

function openPhaseOwnersTab(){
  openPhaseOwnershipModal();
  setTimeout(()=>psSwitchTab("owners"), 0);
}

function openPhaseOwnershipModal() {
  if (!requirePerm('managePhaseOwnership', 'No permission to manage phase ownership.')) return;
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  document.getElementById("psOverlay")?.classList.add("open");
  document.getElementById("psModal")?.classList.add("open");
  document.getElementById("psOverlay")?.setAttribute("aria-hidden", "false");
  document.getElementById("psModal")?.setAttribute("aria-hidden", "false");

  const ov = document.getElementById("psOverlay");
  if (ov && !ov._bound) {
    ov._bound = true;
    ov.addEventListener("click", closePhaseOwnershipModal);
  }

  window.addEventListener("keydown", psEscClosePhase);

  const sub = document.getElementById("psModalSub");
  if (sub) sub.textContent = `${project.name || "Project"} — Assign specific users to project phases`;

  psSwitchTab('owners');
  psRenderPhaseOwnershipModal();
}

function psEscClosePhase(e) {
  if (e.key === "Escape") updatePhaseActivationSummaryUI();
  closePhaseOwnershipModal();
}

function closePhaseOwnershipModal() {
  document.getElementById("psOverlay")?.classList.remove("open");
  document.getElementById("psModal")?.classList.remove("open");
  document.getElementById("psOverlay")?.setAttribute("aria-hidden", "true");
  document.getElementById("psModal")?.setAttribute("aria-hidden", "true");
  window.removeEventListener("keydown", psEscClosePhase);
}

function psRenderPhaseOwnershipModal() {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  const list = document.getElementById("psPhaseOwnerList");
  if (!list) return;
  list.innerHTML = "";

  const dir = psGetRegisteredUserDirectory();
  const phaseIds = wfProjectPhaseIds(state.currentProjectId, { includeInactive: false });
  const blinds = state.blinds.filter((b) => b.projectId === project.id);

  phaseIds.forEach((ph) => {
    const card = document.createElement("div");
    card.className = "ps-owner-card";

    const count = blinds.filter((b) => (b.phase || "broken") === ph).length;

    const owners = psNormalizeIdList(project.phaseOwners.phases[ph] || [], dir);
    project.phaseOwners.phases[ph] = owners;

    card.innerHTML = `
      <div class="ps-owner-phase">
        <div class="title">${escapeHtml(phaseLabel(ph))}</div>
        <div class="count">${count} Blinds</div>
      </div>

      <div class="ps-owner-right">
        <div class="ps-chips" data-chips="${ph}"></div>

        <div class="ps-search-wrap">
          <input class="ps-search" data-search="${ph}" type="text" placeholder="Type a name (e.g., Ab)…" autocomplete="off" />
          <div class="ps-suggest hidden" data-suggest="${ph}"></div>
        </div>

        <div class="tiny muted">Owners of this phase are allowed to update this phase inside this project.</div>
      </div>
    `;

    list.appendChild(card);

    // Render chips + wire search
    psRenderPhaseChips(ph);

    const input = card.querySelector(`[data-search="${ph}"]`);
    const suggest = card.querySelector(`[data-suggest="${ph}"]`);

    if (input) {
      input.addEventListener("input", () => {
        psRenderSuggestions(ph, input.value, suggest);
      });
      input.addEventListener("focus", () => {
        psRenderSuggestions(ph, input.value, suggest);
      });
      input.addEventListener("blur", () => {
        setTimeout(() => suggest?.classList.add("hidden"), 140);
      });

      input.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          suggest?.classList.add("hidden");
          input.blur();
        }
      });
    }
  });
}

function psRenderPhaseChips(phaseId) {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;
  const dir = psGetRegisteredUserDirectory();

  const box = document.querySelector(`[data-chips="${phaseId}"]`);
  if (!box) return;

  const owners = psNormalizeIdList(project.phaseOwners.phases[phaseId] || [], dir);
  project.phaseOwners.phases[phaseId] = owners;

  box.innerHTML = "";

  if (!owners.length) {
    const empty = document.createElement("div");
    empty.className = "muted";
    empty.textContent = "No owners";
    box.appendChild(empty);
    return;
  }

  owners.forEach((uid) => {
    const u = psUserById(uid, dir);
    const name = u ? u.label : String(uid);

    const chip = document.createElement("span");
    chip.className = "ps-chip";

    const avatar = u && u.profileImage
      ? `<img class="ps-avatar-img" src="${u.profileImage}" alt=""/>`
      : `<span class="ps-avatar-initials">${escapeHtml(psGetInitials(name))}</span>`;

    chip.innerHTML = `${avatar}<span>${escapeHtml(name)}</span><button type="button" title="Remove">×</button>`;

    chip.querySelector("button")?.addEventListener("click", () => {
      if (!requirePerm('managePhaseOwnership', 'No permission.')) return;
      project.phaseOwners.phases[phaseId] = owners.filter((x) => x !== uid);
      saveState();
      psRenderPhaseChips(phaseId);
      renderProjectSettingsSummary();
    });

    box.appendChild(chip);
  });
}

function psRenderSuggestions(phaseId, query, suggestBox) {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;
  const dir = psGetRegisteredUserDirectory();

  if (!suggestBox) return;

  const q = String(query || "").trim().toLowerCase();
  const owners = new Set(project.phaseOwners.phases[phaseId] || []);

  const candidates = dir
    .filter((u) => !owners.has(u.id))
    .filter((u) => {
      if (!q) return true;
      return (
        (u.label || "").toLowerCase().includes(q) ||
        (u.username || "").toLowerCase().includes(q)
      );
    })
    .slice(0, 12);

  if (!candidates.length) {
    suggestBox.innerHTML = `<div class="ps-suggest-item"><div class="muted">No matches</div></div>`;
    suggestBox.classList.remove("hidden");
    return;
  }

  suggestBox.innerHTML = "";

  candidates.forEach((u) => {
    const item = document.createElement("div");
    item.className = "ps-suggest-item";

    const avatar = u.profileImage
      ? `<img class="ps-avatar-img" src="${u.profileImage}" alt=""/>`
      : `<span class="ps-avatar-initials">${escapeHtml(psGetInitials(u.label))}</span>`;

    item.innerHTML = `
      ${avatar}
      <div>
        <div class="ps-suggest-name">${escapeHtml(u.label)}</div>
        <div class="ps-suggest-meta">${escapeHtml(u.username || (ROLE_LABELS[u.role] || u.role || ""))}</div>
      </div>
    `;

    item.addEventListener("click", () => {
      if (!requirePerm('managePhaseOwnership', 'No permission.')) return;
      const arr = project.phaseOwners.phases[phaseId] || [];
      if (!arr.includes(u.id)) arr.push(u.id);
      project.phaseOwners.phases[phaseId] = arr;
      saveState();

      // clear input and suggestions
      const input = document.querySelector(`[data-search="${phaseId}"]`);
      if (input) input.value = "";
      suggestBox.classList.add("hidden");

      psRenderPhaseChips(phaseId);
      renderProjectSettingsSummary();
    });

    suggestBox.appendChild(item);
  });

  suggestBox.classList.remove("hidden");
}

function saveProjectPhaseOwners() {
  if (!requirePerm('managePhaseOwnership', 'No permission.')) return;
  ensureCurrentProjectPhaseOwners();
  saveState();
  renderProjectSettingsSummary();
  toast("Phase owners saved ✅");
  updatePhaseActivationSummaryUI();
  closePhaseOwnershipModal();
}

function resetProjectPhaseOwnersToDefault() {
  if (!requirePerm('managePhaseOwnership', 'No permission.')) return;
  const project = state.projects.find((p) => p.id === state.currentProjectId);
  if (!project) return;
  if (!confirm("Reset phase owners for this project?")) return;
  sbtsAutoBackupBefore("Reset Phase Owners");
  project.phaseOwners = defaultProjectPhaseOwners();
  // ensure keys
  ensureCurrentProjectPhaseOwners();
  saveState();
  renderProjectSettingsSummary();
  toast("Reset done ✅");
}

function exportCurrentProjectPhaseOwners() {
  if (!requirePerm('managePhaseOwnership', 'No permission.')) return;
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;
  const data = JSON.stringify(project.phaseOwners, null, 2);
  const blob = new Blob([data], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `sbts-phase-owners-${project.name || project.id}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  toast("Exported ✅");
}

/* ==========================
   PHASE OWNERSHIP REQUESTS
   - Any user can submit a request.
   - Admin/Supervisor (manageRequests) can approve/reject.
========================== */

let psReqKind = "owner"; // owner | assist

function getRequestApproverIds() {
  const ids = [];
  (state.users || []).forEach((u) => {
    if (!u?.id) return;
    if (normalizeRole(u.role) === "admin") {
      ids.push(u.id);
      return;
    }
    if (userHasPerm(u.id, "manageRequests")) ids.push(u.id);
  });
  return Array.from(new Set(ids));
}

function openPSRequestModal(kind) {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  psReqKind = kind === "assist" ? "assist" : "owner";

  const title = document.getElementById("psRequestTitle");
  const sub = document.getElementById("psRequestSub");
  if (title) title.textContent = psReqKind === "owner" ? "Request owner change" : "Request assistance";
  if (sub) sub.textContent = psReqKind === "owner"
    ? "Ask Admin/Supervisor to add a phase owner in this project."
    : "Ask for extra support in a specific phase (adds Support tag, not ownership).";

  // populate phases
  const phaseSel = document.getElementById("psReqPhase");
  if (phaseSel) {
    phaseSel.innerHTML = "";
    wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
      const opt = document.createElement("option");
      opt.value = ph;
      opt.textContent = phaseLabel(ph);
      phaseSel.appendChild(opt);
    });
  }

  // suggested user list
  const sugSel = document.getElementById("psReqSuggestedUser");
  if (sugSel) {
    sugSel.innerHTML = "";
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "(Optional)";
    sugSel.appendChild(empty);

    const dir = psGetRegisteredUserDirectory();
    dir.forEach((u) => {
      const opt = document.createElement("option");
      opt.value = u.id;
      opt.textContent = `${u.label}${u.username ? " — " + u.username : ""}`;
      sugSel.appendChild(opt);
    });
  }

  // defaults
  const reason = document.getElementById("psReqReason");
  if (reason) reason.value = "";

  openModal("psRequestModal");
}

function submitPSRequest() {
  const project = ensureCurrentProjectPhaseOwners();
  if (!project) return;

  // Robust project identity (prevents "No project" notifications)
  const pid = project.id || state.currentProjectId || null;
  const pname = project.name || project.projectName || (pid ? (state.projects.find(p => p.id === pid)?.name || pid) : "");
  if (!pid) return toast("Open a project first");

  const phaseId = document.getElementById("psReqPhase")?.value || "";
  if (!phaseId) return toast("Choose a phase");

  const specialty = document.getElementById("psReqSpecialty")?.value || "";
  const suggestedUserId = document.getElementById("psReqSuggestedUser")?.value || "";
  const reason = (document.getElementById("psReqReason")?.value || "").trim();

  // Use stable user id (some accounts persist as username/email without `id`)
  const requesterId = getCurrentUserIdStable();
  const requesterName = (state.currentUser?.name || state.currentUser?.fullName || state.currentUser?.username || state.currentUser?.displayName) || userDisplayNameById(requesterId);

  const req = {
    id: uid("req"),
    kind: psReqKind,
    status: "pending", // pending | in_review | approved | rejected | canceled
    comments: [],
    createdAt: new Date().toISOString(),
    resolvedAt: null,
    projectId: pid,
    projectName: pname,
    phaseId,
    requestedBy: requesterId || null,
    requestedByName: requesterName || "User",
    specialty,
    suggestedUserId: suggestedUserId || null,
    reason,
  };

  state.requests = Array.isArray(state.requests) ? state.requests : [];
  state.requests.unshift(req);
  state.requests = state.requests.slice(0, 300);
  saveState();

  // Notify approvers
  const approvers = getRequestApproverIds();
  const kindLabel = req.kind === "owner" ? "Owner change" : "Assistance";
  addNotification(approvers, {
    requiresAction: true,
    type: "action",
    title: `New request: ${kindLabel}`,
    message: `${req.requestedByName} requested ${kindLabel.toLowerCase()} for phase “${phaseLabel(phaseId)}”.`,
    projectId: req.projectId,
    projectName: req.projectName,
    phaseId: req.phaseId,
    actionKey: `req:${req.id}`,
    actorId: req.requestedBy,
  });

  // Confirm to requester (so they can see it in their inbox even before approval)
  if (requesterId) {
    addNotification([requesterId], {
      type: "info",
      title: "Request submitted",
      message: `Your request was sent for phase “${phaseLabel(phaseId)}”. You will be notified when it is approved or rejected.`,
      projectId: req.projectId,
      projectName: req.projectName,
      phaseId: req.phaseId,
      resolved: true,
    });
  }

  closeModal("psRequestModal");
  toast("Request sent ✅");
}



/* ==========================
   INBOX ENGINE (Requests + Status + Comments)
   - Used by Notification Center Drawer (Tab: Inbox)
   - Local-only, stored in state.requests
========================== */
function ensureRequestsArray() {
  state.requests = Array.isArray(state.requests) ? state.requests : [];
}

// ===== Patch 46.4: Request Types UI (Create Request Modal) =====
const REQUEST_TYPE_DEFS = [
  { key: "NEW_USER", label: "New user access request" },
  { key: "TRANSFER_PHASE", label: "Transfer phase ownership" },
  { key: "ADD_HELPER", label: "Add helper to phase updates" },
  { key: "DATA_CHANGE", label: "Blind data change request" },
  { key: "ADMIN_NOTE", label: "Admin note" },
];

function requestTypeLabel(key){
  return (REQUEST_TYPE_DEFS.find(x => x.key === key)?.label) || (key || "Request");
}

function openCreateRequestModal(prefill = {}) {
  ensureRequestsArray();
  // build options
  const typeSel = document.getElementById("reqCreateType");
  const projSel = document.getElementById("reqCreateProject");
  const blindSel = document.getElementById("reqCreateBlind");
  const phaseSel = document.getElementById("reqCreatePhase");
  const tgtInp = document.getElementById("reqCreateTargetUser");


// Type-based UI: NEW_USER requests are global (no project/blind/phase)
const projWrap = projSel?.closest("div");
const blindWrap = blindSel?.closest("div");
const phaseWrap = phaseSel?.closest("div");
const applyReqTypeVisibility = () => {
  const t = typeSel?.value || "NEW_USER";
  const isNewUser = (t === "NEW_USER");
  if (projWrap) projWrap.style.display = isNewUser ? "none" : "";
  if (blindWrap) blindWrap.style.display = isNewUser ? "none" : "";
  if (phaseWrap) phaseWrap.style.display = isNewUser ? "none" : "";
  // also clear values when hidden
  if (isNewUser) {
    if (projSel) projSel.value = "";
    if (blindSel) blindSel.value = "";
    if (phaseSel) phaseSel.value = "";
  }
};
  const priSel = document.getElementById("reqCreatePriority");
  const reason = document.getElementById("reqCreateReason");
  const submit = document.getElementById("reqCreateSubmitBtn");

  if (!typeSel || !projSel || !blindSel || !phaseSel || !submit) return;

  // types
  typeSel.innerHTML = REQUEST_TYPE_DEFS.map(t => `<option value="${t.key}">${escapeHtml(t.label)}</option>`).join("");

  // projects
  const projects = Array.isArray(state.projects) ? state.projects.slice() : [];
  projSel.innerHTML = `<option value="">— No project —</option>` + projects
    .map(p => `<option value="${p.id}">${escapeHtml(p.name || p.id)}</option>`).join("");

  // phases
  phaseSel.innerHTML = `<option value="">— No phase —</option>` + (PHASES || []).map(pid => {
    return `<option value="${pid}">${escapeHtml(phaseLabel(pid))}</option>`;
  }).join("");

  // set defaults
  const defPid = prefill.projectId || state.currentProjectId || "";
  projSel.value = defPid;
  typeSel.value = prefill.type || "NEW_USER";

if (typeSel) {
  typeSel.onchange = () => {
    applyReqTypeVisibility();
    // refresh dependent options when switching away from NEW_USER
    try { buildReqCreateBlindOptions(); } catch(e){}
  };
}
applyReqTypeVisibility();
  phaseSel.value = prefill.phaseId || "";
  if (tgtInp) tgtInp.value = prefill.targetUser || "";
  if (priSel) priSel.value = prefill.priority || "normal";
  if (reason) reason.value = prefill.reason || "";

  // populate blinds based on project
  const repopBlinds = () => {
    const pid = projSel.value || "";
    const list = (state.blinds || []).filter(b => !pid || b.projectId === pid);
    const opts = [`<option value="">— No blind —</option>`].concat(list
      .slice()
      .sort((a,b)=>String(a.tagNo||a.tag_no||a.id).localeCompare(String(b.tagNo||b.tag_no||b.id)))
      .map(b => `<option value="${b.id}">${escapeHtml(formatBlindStoryInline(b, projects.find(p=>p.id===b.projectId)?.name||""))}</option>`));
    blindSel.innerHTML = opts.join("");
    blindSel.value = prefill.blindId || "";
  };

  repopBlinds();
  projSel.onchange = repopBlinds;

  // show
  openModal("reqCreateModal");

  // submit handler (replace any previous)
  submit.onclick = () => {
    const type = typeSel.value || "NEW_USER";
    const isNewUser = (type === "NEW_USER");
    const pid = isNewUser ? "" : (projSel.value || "");
    const pName = pid ? (state.projects.find(p => p.id === pid)?.name || pid) : "";
    const bid = blindSel.value || "";
    const phaseId = phaseSel.value || "";
    const targetUser = (tgtInp?.value || "").trim();
    const priority = priSel?.value || "normal";
    const reasonTxt = (reason?.value || "").trim();

    const uid_ = getCurrentUserIdStable();
    const uname_ = (state.currentUser?.name || state.currentUser?.fullName || state.currentUser?.username || state.currentUser?.displayName) || userDisplayNameById(uid_) || "User";

    // Decide who should receive the request
    let recipients = [];
    let scope = "user";

    if (type === "NEW_USER") {
      recipients = getUserApproverIds().filter(id => id && id !== uid_);
      scope = "user";
    } else if (type === "ADMIN_NOTE") {
      // Admin note is a notification (not an actionable request)
      const noteScope = pid ? "project" : "global";
      SBTS_ACTIVITY.pushNotification({
        category: "admin",
        scope: noteScope,
        projectId: pid || null,
        title: "Admin note",
        message: reasonTxt || "Admin note added.",
        requiresAction: false,
        resolved: true,
        actorId: uid_ || null
      });

      // Confirmation
      if (uid_) SBTS_ACTIVITY.pushNotification({
        category: "system",
        scope: "user",
        recipients: [uid_],
        title: "Note published",
        message: pid ? `Admin note posted to ${pName}.` : "Admin note posted globally.",
        resolved: true
      });

      closeModal("reqCreateModal");
      toast("Admin note published ✅");
      updateNotificationsBadge();
      try { renderRequestsInbox(); } catch(e){}
      return;
    } else {
      // Project-related approvals go to admins (for now)
      recipients = getAdminUserIds().filter(id => id && id !== uid_);
      scope = "user";
    }

    // Create Inbox request + notification for recipients
    const req = SBTS_ACTIVITY.pushRequest({
      requestType: type,
      title: `${requestTypeLabel(type)}`,
      message: `${uname_} submitted: ${requestTypeLabel(type)}${reasonTxt ? ` — ${reasonTxt}` : ""}`,
      scope,
      recipients,
      priority,
      projectId: pid || null,
      projectName: pName || null,
      blindId: bid || null,
      phaseId: phaseId || null,
      targetUser: targetUser || null,
      reason: reasonTxt || null,
      requestedBy: uid_ || null,
      requestedByName: uname_,
      meta: {
        source: "create_request_modal",
        targetUser: targetUser || null
      }
    });

    // Confirmation to requester
    if (uid_) {
      SBTS_ACTIVITY.pushNotification({
        category: "system",
        scope: "user",
        recipients: [uid_],
        title: "Request submitted",
        message: `${formatReqContextLine(req)} — ${req.kind}`,
        projectId: req.projectId,
        blindId: req.blindId,
        resolved: true
      });
    }

    closeModal("reqCreateModal");
    toast("Request created ✅");

    // Focus it inside drawer inbox
    const nd = ensureNotifDrawerUI();
    nd.selectedReqId = req.id;
    state.ui.notifDrawer = state.ui.notifDrawer || {};
    state.ui.notifDrawer.tab = "inbox";
    saveState();
    openNotificationsDrawer();
    renderNotificationsDrawer();
    updateNotificationsBadge();

    // Refresh inbox page if user is on it
    try { renderRequestsInbox(); } catch(e){}
  };
}



// ===== Patch 46.5: Contextual Requests (create from source pages) =====
function openReqTransferPhaseFromProject(){
  if(!state.currentProjectId) { toast("Open a project first"); return; }
  openCreateRequestModal({ type: "TRANSFER_PHASE", projectId: state.currentProjectId });
}
function openReqAddHelperFromProject(){
  if(!state.currentProjectId) { toast("Open a project first"); return; }
  openCreateRequestModal({ type: "ADD_HELPER", projectId: state.currentProjectId });
}
function openReqNewUserAccess(){
  openCreateRequestModal({ type: "NEW_USER" });
}

function requestStatusLabel(s){
  const m = {
    pending: "Pending",
    in_review: "In Review",
    approved: "Approved",
    rejected: "Rejected",
    canceled: "Canceled"
  };
  return m[s] || String(s || "pending");
}

function isRequestOpen(s){
  return (s === "pending" || s === "in_review");
}


function canEditOrDeleteRequest(req){
  const uid = getCurrentUserIdStable();
  if (!req || !uid) return false;
  if (canUser("manageRequests")) return true;
  return (req.requestedById && req.requestedById === uid);
}

function archiveRequest(reqId){
  ensureRequestsArray();
  const r = (state.requests||[]).find(x=>x && x.id===reqId);
  if (!r) return;
  if (!canEditOrDeleteRequest(r)) return toast("Not allowed.");
  r.status = "archived";
  r.updatedAt = new Date().toISOString();
  saveState();
  toast("Archived ✅");
}

function deleteRequest(reqId){
  ensureRequestsArray();
  const r = (state.requests||[]).find(x=>x && x.id===reqId);
  if (!r) return;
  if (!canEditOrDeleteRequest(r)) return toast("Not allowed.");
  if (!confirm("Delete this request? This cannot be undone.")) return;
  state.requests = (state.requests||[]).filter(x=>x && x.id!==reqId);
  saveState();
  toast("Deleted ✅");
}
function addRequestComment(reqId, text){
  ensureRequestsArray();
  const req = state.requests.find(r => r && r.id === reqId);
  if(!req) return;
  const t = (text || "").trim();
  if(!t) return;
  const uid_ = getCurrentUserIdStable();
  const uname_ = (state.currentUser?.name || state.currentUser?.fullName || state.currentUser?.username || state.currentUser?.displayName) || userDisplayNameById(uid_) || "User";
  req.comments = Array.isArray(req.comments) ? req.comments : [];
  const comment = {
    id: uid("c"),
    by: uid_ || null,
    byName: uname_,
    at: new Date().toISOString(),
    text: t
  };
  req.comments.push(comment);
  req.updatedAt = new Date().toISOString();
  saveState();

  // Notify other participants (requester + assignees) that a new message was posted
  try {
    const targets = Array.from(new Set([
      req.requestedBy,
      ...(Array.isArray(req.assigneeIds) ? req.assigneeIds : [])
    ].filter(Boolean))).filter(id => id !== uid_);
    if (targets.length) {
      const ctx = formatReqContextLine(req);
      const preview = (t.length > 90) ? (t.slice(0, 87) + "...") : t;
      SBTS_ACTIVITY.pushNotification({
        category: "system",
        scope: "user",
        recipients: targets,
        projectId: req.projectId || null,
        blindId: req.blindId || null,
        title: `New message: ${req.kind || "Request"}`,
        message: `${ctx} • ${uname_}: ${preview}`,
        requiresAction: false,
        resolved: true,
        actionKey: `request:${req.id}`
      });
      updateNotificationsBadge();
    }
  } catch(e) {
    console.warn("[addRequestComment] notify failed", e);
  }
}

function setRequestStatus(reqId, status){
  ensureRequestsArray();
  const req = state.requests.find(r => r && r.id === reqId);
  if(!req) return;
  req.status = status;
  req.updatedAt = new Date().toISOString();
  if(!isRequestOpen(status)) req.resolvedAt = new Date().toISOString();
  saveState();
}

function formatReqContextLine(req){
  // Use same story format used in Recent Activity when possible
  const pid = req.projectId || null;
  const projectName = req.projectName || (pid ? (state.projects.find(p => p.id === pid)?.name || pid) : "");
  // Some request kinds may not have blindId; keep it clean
  let ctx = "";
  if (req.blindId) {
    const blind = state.blinds.find(b => b.id === req.blindId);
    ctx = formatBlindStoryInline(blind, projectName);
  } else {
    const parts = [];
    if (projectName) parts.push(projectName);
    if (req.phaseId) parts.push(req.phaseId);
    ctx = parts.join(" | ");
  }
  return ctx || projectName || "Request";
}
function openRequestsInbox() {
  if (!requirePerm("manageRequests", "No permission.")) return;
  // populate project filter
  const sel = document.getElementById("reqInboxProject");
  if (sel) {
    sel.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "all";
    optAll.textContent = "All projects";
    sel.appendChild(optAll);
    (state.projects || []).forEach((p) => {
      const o = document.createElement("option");
      o.value = p.id;
      o.textContent = p.name || p.id;
      sel.appendChild(o);
    });
    sel.value = state.currentProjectId || "all";
  }
  const st = document.getElementById("reqInboxStatus");
  if (st && !st.value) st.value = "pending";

  openModal("requestsInboxModal");
  renderRequestsInbox();
}

function renderRequestsInbox() {
  if (!canUser("manageRequests")) return;
  const box = document.getElementById("requestsInboxList");
  if (!box) return;
  const q = (document.getElementById("reqInboxSearch")?.value || "").trim().toLowerCase();
  const status = document.getElementById("reqInboxStatus")?.value || "pending";
  const projectFilter = document.getElementById("reqInboxProject")?.value || "all";

  const list = Array.isArray(state.requests) ? state.requests : [];
  const filtered = list.filter((r) => {
    if (!r) return false;
    if (status !== "all" && (r.status || "pending") !== status) return false;
    if (projectFilter !== "all" && r.projectId !== projectFilter) return false;
    if (!q) return true;
    const hay = `${r.kind} ${r.status} ${r.projectId} ${r.phaseId} ${r.requestedByName} ${(r.reason || "")}`.toLowerCase();
    return hay.includes(q);
  });

  if (!filtered.length) {
    box.innerHTML = `<div class="muted" style="padding:12px;">No requests.</div>`;
    return;
  }

  box.innerHTML = "";
  filtered.forEach((r) => {
    const card = document.createElement("div");
    card.className = "notif-item";
    card.id = `req_${r.id}`;
    const proj = (state.projects || []).find((p) => p.id === r.projectId);
    const projName = proj ? (proj.name || proj.id) : (r.projectName || r.projectId || "-");
    const kindLabel = r.kind === "owner" ? "Owner change" : "Assistance";
    const st = r.status || "pending";

    const badge = st === "approved" ? "badge ok" : st === "rejected" ? "badge danger" : "badge warn";
    const badgeText = st.charAt(0).toUpperCase() + st.slice(1);

    card.innerHTML = `
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;">
        <div>
          <div style="font-weight:800;">${escapeHtml(kindLabel)} <span class="${badge}">${badgeText}</span></div>
          <div class="tiny muted" style="margin-top:4px;">
            Project: <b>${escapeHtml(projName)}</b> • Phase: <b>${escapeHtml(phaseLabel(r.phaseId))}</b>
          </div>
          <div class="tiny" style="margin-top:6px;">From: <b>${escapeHtml(r.requestedByName || "-")}</b>${r.specialty ? ` • Specialty: <b>${escapeHtml(r.specialty)}</b>` : ""}</div>
          ${r.reason ? `<div class="tiny" style="margin-top:6px;">Reason: ${escapeHtml(r.reason)}</div>` : ""}
        </div>
        <div class="notif-actions" style="display:flex;gap:8px;flex-wrap:wrap;">
          ${r.projectId ? `<button class="secondary-btn btn-small" onclick="openProjectSettingsForProject('${r.projectId}')">Open project</button>` : ``}
          ${st === "pending" ? `
            <button class="primary-btn btn-small" onclick="approveRequest('${r.id}')">Approve</button>
            <button class="secondary-btn btn-small" onclick="rejectRequest('${r.id}')">Reject</button>
          ` : ""}
        </div>
      </div>
    `;
    box.appendChild(card);
  });

  // Focus (when opened from a notification)
  const focusId = state.ui?.reqInboxFocusId;
  if (focusId) {
    const el = document.getElementById(`req_${focusId}`);
    if (el) {
      el.style.outline = "2px solid rgba(10,103,163,0.35)";
      el.style.outlineOffset = "2px";
      el.scrollIntoView({ block: "center", behavior: "smooth" });
    }
    // Clear after first use
    state.ui.reqInboxFocusId = null;
    saveState();
  }
}

function approveRequest(reqId) {
  if (!requirePerm("manageRequests", "No permission.")) return;
  const req = (state.requests || []).find((r) => r && r.id === reqId);
  if (!req || req.status !== "pending") return;

  const project = state.projects.find((p) => p.id === req.projectId);
  if (!project) return toast("Project not found");
  // Ensure keys for the targeted project (without changing user UI)
  const prevPid = state.currentProjectId;
  state.currentProjectId = req.projectId;
  ensureCurrentProjectPhaseOwners();
  state.currentProjectId = prevPid;

  // Choose who to add: suggested user or requester
  const userId = req.suggestedUserId || req.requestedBy;
  if (!userId) return toast("No user to add");

  if (req.kind === "owner") {
    project.phaseOwners = project.phaseOwners || defaultProjectPhaseOwners();
    project.phaseOwners.phases = project.phaseOwners.phases || {};
    const arr = Array.isArray(project.phaseOwners.phases[req.phaseId]) ? project.phaseOwners.phases[req.phaseId] : [];
    if (!arr.includes(userId)) arr.push(userId);
    project.phaseOwners.phases[req.phaseId] = arr;
  } else {
    project.phaseOwners = project.phaseOwners || defaultProjectPhaseOwners();
    project.phaseOwners.support = project.phaseOwners.support || {};
    const arr = Array.isArray(project.phaseOwners.support[req.phaseId]) ? project.phaseOwners.support[req.phaseId] : [];
    if (!arr.includes(userId)) arr.push(userId);
    project.phaseOwners.support[req.phaseId] = arr;
  }

  req.status = "approved";
  req.resolvedAt = new Date().toISOString();
  saveState();
  renderProjectSettingsSummary();
  renderRequestsInbox();

  // Notify requester
  if (req.requestedBy) {
    const kindLabel = req.kind === "owner" ? "owner change" : "assistance";
    addNotification([req.requestedBy], {
      type: "info",
      title: "Request approved ✅",
      message: `Your ${kindLabel} request for phase “${phaseLabel(req.phaseId)}” was approved.`,
      projectId: req.projectId,
      phaseId: req.phaseId,
      resolved: true,
    });
  }

  // Notify the person who was added (Owner/Support)
  // This is the key missing piece that made helpers not receive any notification.
  if (userId) {
    const roleLabel = req.kind === "owner" ? "Phase Owner" : "Support";
    const msg = req.kind === "owner"
      ? `You were assigned as <b>${roleLabel}</b> for phase “${phaseLabel(req.phaseId)}” in project “${escapeHtml(project.name || project.id)}”.`
      : `You were assigned as <b>${roleLabel}</b> for phase “${phaseLabel(req.phaseId)}” in project “${escapeHtml(project.name || project.id)}”.`;
    addNotification([userId], {
      requiresAction: false,
      type: "info",
      title: `${roleLabel} assigned`,
      message: msg,
      projectId: req.projectId,
      phaseId: req.phaseId,
      actorId: getCurrentUserIdStable(),
      resolved: true,
    });
  }

  toast("Approved ✅");
}

function rejectRequest(reqId) {
  if (!requirePerm("manageRequests", "No permission.")) return;
  const req = (state.requests || []).find((r) => r && r.id === reqId);
  if (!req || req.status !== "pending") return;

  req.status = "rejected";
  req.resolvedAt = new Date().toISOString();
  saveState();
  renderRequestsInbox();

  if (req.requestedBy) {
    const kindLabel = req.kind === "owner" ? "owner change" : "assistance";
    addNotification([req.requestedBy], {
      type: "info",
      title: "Request rejected",
      message: `Your ${kindLabel} request for phase “${phaseLabel(req.phaseId)}” was rejected.`,
      projectId: req.projectId,
      phaseId: req.phaseId,
      resolved: true,
    });
  }
  toast("Rejected");
}

function renderProjectPhaseCards() {
  const box = document.getElementById("projectPhaseCards");
  box.innerHTML = "";

  const blinds = state.blinds.filter((b) => b.projectId === state.currentProjectId);
  const counts = {};
  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((p) => (counts[p] = 0));
  blinds.forEach((b) => { counts[b.phase || "broken"]++; });

  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
    const card = document.createElement("div");
    card.className = "phase-card";
    card.setAttribute("data-phase", ph);
    card.innerHTML = `
      <div class="phase-left">
        <span class="phase-dot" style="background:${wfPhaseColor(ph)}"></span>
        <div class="phase-name">${phaseLabel(ph)}</div>
      </div>
      <div class="phase-count">${counts[ph] || 0}</div>
    `;
    box.appendChild(card);
  });
}

function renderProjectBlindsTable() {
  const tbody = document.getElementById("projectBlindsTableBody");
  const cards = document.getElementById("projectBlindsCards");

  // Mobile: cards view (<= 767px)
  if (cards && isMobileView()) {
    if (tbody) tbody.innerHTML = "";
    renderProjectBlindsCards(cards);
    return;
  }

  if (!tbody) return;
  tbody.innerHTML = "";

  const blinds = state.blinds.filter((b) => b.projectId === state.currentProjectId);
  blinds.forEach((b, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${b.name}</td>
      <td>${b.line || "-"}</td>
      <td>${b.type || "-"}</td>
      <td>${b.size || "-"}</td>
      <td>${b.rate || "-"}</td>
      <td>${phaseLabel(b.phase || "broken")}</td>
      <td>
        <div class="table-actions">
          <button class="secondary-btn tiny" onclick="openBlindDetails('${b.id}')">Details</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function isMobileView() {
  try {
    return window.matchMedia && window.matchMedia("(max-width: 767px)").matches;
  } catch (e) {
    return (window.innerWidth || 1024) <= 767;
  }
}

function blindStatusLabel(blind) {
  const ph = (blind?.phase || "broken");
  if (ph === "inspectionReady") return "Completed";
  if (ph === "finalTight") return "In review";
  return "In progress";
}

function renderProjectBlindsCards(container) {
  const blinds = state.blinds.filter((b) => b.projectId === state.currentProjectId);
  if (!blinds.length) {
    container.innerHTML = `<div class="muted" style="padding:12px;">No blinds found.</div>`;
    return;
  }

  container.innerHTML = "";
  blinds.forEach((b) => {
    const card = document.createElement("div");
    card.className = "blind-card";
    const status = blindStatusLabel(b);
    const stClass = status === "Completed" ? "badge ok" : status === "In review" ? "badge warn" : "badge";

    card.innerHTML = `
      <div class="blind-card-head">
        <div>
          <div class="blind-card-title">${escapeHtml(b.name || "-")}</div>
          <div class="blind-card-sub">${escapeHtml(b.line || "-")}</div>
        </div>
        <span class="${stClass}">${escapeHtml(status)}</span>
      </div>
      <div class="blind-card-meta">
        <span class="pill">Phase: <b>${escapeHtml(phaseLabel(b.phase || "broken"))}</b></span>
      </div>
      <div class="blind-card-actions">
        <button class="secondary-btn" onclick="openBlindQrModal('${b.id}')"><i class="ph ph-qr-code"></i> QR</button>
        <button class="primary-btn" onclick="openBlindDetails('${b.id}')"><i class="ph ph-folder-open"></i> Open</button>
        <button class="btn-neutral" onclick="openBlindMoreModal('${b.id}')"><i class="ph ph-dots-three-outline"></i> More</button>
      </div>
    `;
    container.appendChild(card);
  });
}

function openBlindMoreModal(blindId) {
  const blind = state.blinds.find((b) => b.id === blindId);
  if (!blind) return;
  const proj = state.projects.find((p) => p.id === blind.projectId);
  const area = state.areas.find((a) => a.id === blind.areaId);
  const title = document.getElementById("blindMoreTitle");
  const sub = document.getElementById("blindMoreSub");
  const body = document.getElementById("blindMoreBody");
  const openBtn = document.getElementById("blindMoreOpenBtn");

  if (title) title.textContent = `Blind: ${blind.name || "-"}`;
  if (sub) sub.textContent = `${proj ? proj.name : "-"} • ${area ? area.name : "-"}`;
  if (body) {
    body.innerHTML = `
      <div class="kv"><div class="k">Line</div><div class="v">${escapeHtml(blind.line || "-")}</div></div>
      <div class="kv"><div class="k">Type</div><div class="v">${escapeHtml(blind.type || "-")}</div></div>
      <div class="kv"><div class="k">Size</div><div class="v">${escapeHtml(blind.size || "-")}</div></div>
      <div class="kv"><div class="k">Rate</div><div class="v">${escapeHtml(blind.rate || "-")}</div></div>
    `;
  }
  if (openBtn) {
    openBtn.onclick = () => {
      closeModal("blindMoreModal");
      openBlindDetails(blindId);
    };
  }
  openModal("blindMoreModal");
}

function openBlindQrModal(blindId) {
  const blind = state.blinds.find((b) => b.id === blindId);
  if (!blind) return;
  const link = buildPublicBlindUrl(blind.id);
  const img = document.getElementById("blindQrModalImg");
  const linkEl = document.getElementById("blindQrModalLink");
  if (img) img.src = buildQrImageUrl(link);
  if (linkEl) linkEl.textContent = link;
  openModal("blindQrModal");
}

function copyBlindQrModalLink() {
  const el = document.getElementById("blindQrModalLink");
  if (!el) return;
  const txt = (el.textContent || "").trim();
  if (!txt) return;
  navigator.clipboard?.writeText(txt).then(() => toast("Copied"))
    .catch(() => {
      try {
        const ta = document.createElement("textarea");
        ta.value = txt;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        toast("Copied");
      } catch (e) {}
    });
}

/* ==========================
   ADD BLINDS
========================== */
function openAddSingleBlindModal() {
  if (!canUser("manageBlinds")) return alert("No permission.");
  document.getElementById("singleBlindName").value = "";
  document.getElementById("singleBlindLine").value = "";
  document.getElementById("singleBlindType").value = "Isolation Blind";
  document.getElementById("singleBlindSize").value = "";
  openModal("addSingleBlindModal");
}

function confirmAddSingleBlind() {
  const project = state.projects.find((p) => p.id === state.currentProjectId);
  if (!project) return;

  const name = document.getElementById("singleBlindName").value.trim();
  const line = document.getElementById("singleBlindLine").value.trim();
  const type = document.getElementById("singleBlindType").value;
  const size = document.getElementById("singleBlindSize").value.trim();
  const rate = (document.getElementById("singleBlindRate")?.value || "").trim();

  if (!name) return alert("Blind name required.");

  state.blinds.push({
    id: crypto.randomUUID(),
    areaId: project.areaId,
    projectId: project.id,
    name,
    line,
    type,
    size,
    rate,
    phase: "broken",
    history: [],
    finalApprovals: {},
  });

  // Notify admins (info)
  try {
    const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
    if (admins.length) {
      const blind = state.blinds[state.blinds.length - 1];
      addNotification(admins, {
        type: "blind",
        title: "New blind created",
        message: `${sbtsBlindOfficialText(blind)} — Created`,
        projectId: blind?.projectId || state.currentProjectId || null,
        blindId: blind?.id || null,
        requiresAction: false,
        actorId: state.currentUser?.id || null,
      });
    }
  } catch(e) {}


  saveState();
  closeModal("addSingleBlindModal");
  openProjectDetails(project.id);
  renderDashboard();
}

let bulkTempRows = [];
function openBulkAddBlindsModal() {
  if (!canUser("manageBlinds")) return alert("No permission.");
  bulkTempRows = [
    { name: "", line: "", type: "Isolation Blind", size: "" },
    { name: "", line: "", type: "Isolation Blind", size: "" },
  ];
  renderBulkBlindsRows();
  openModal("bulkAddBlindsModal");
}

function addBulkBlindRow() {
  bulkTempRows.push({ name: "", line: "", type: "Isolation Blind", size: "", rate: "" });
  renderBulkBlindsRows();
}

function bulkUpdateRow(index, field, value) {
  if (!bulkTempRows[index]) return;
  bulkTempRows[index][field] = value;
}

function bulkRemoveRow(index) {
  bulkTempRows.splice(index, 1);
  renderBulkBlindsRows();
}

function renderBulkBlindsRows() {
  const tbody = document.getElementById("bulkBlindsTableBody");
  tbody.innerHTML = "";

  const typeOptions = `
    <option value="Isolation Blind">Isolation Blind</option>
    <option value="Drop Spool">Drop Spool</option>
    <option value="Slip Blind">Slip Blind</option>
  `;

  bulkTempRows.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td><input type="text" value="${row.name || ""}" oninput="bulkUpdateRow(${index}, 'name', this.value)" /></td>
      <td><input type="text" value="${row.line || ""}" oninput="bulkUpdateRow(${index}, 'line', this.value)" /></td>
      <td>
        <select onchange="bulkUpdateRow(${index}, 'type', this.value)">
          ${typeOptions}
        </select>
      </td>
      <td><input type="text" value="${row.size || ""}" oninput="bulkUpdateRow(${index}, 'size', this.value)" /></td>
      <td><input type="text" value="${row.rate || ""}" oninput="bulkUpdateRow(${index}, 'rate', this.value)" /></td>
      <td><button class="secondary-btn tiny" onclick="bulkRemoveRow(${index})">X</button></td>
    `;
    tbody.appendChild(tr);

    // set selected
    tr.querySelector("select").value = row.type || "Isolation Blind";
  });
}

function confirmBulkAddBlinds() {
  const project = state.projects.find((p) => p.id === state.currentProjectId);
  if (!project) return;

  const clean = bulkTempRows.filter((r) => (r.name || "").trim() !== "");
  if (clean.length === 0) return alert("Add at least one blind with a name.");

  clean.forEach((r) => {
    state.blinds.push({
      id: crypto.randomUUID(),
      areaId: project.areaId,
      projectId: project.id,
      name: r.name.trim(),
      line: (r.line || "").trim(),
      type: (r.type || "Isolation Blind"),
      size: (r.size || "").trim(),
      rate: (r.rate || "").trim(),
      phase: "broken",
      history: [],
      finalApprovals: {},
    });
  });

  
  // Notify admins (info) - bulk create summary
  try {
    const admins = getAdminUserIds().filter(uid => uid && uid !== (state.currentUser?.id || null));
    if (admins.length) {
      addNotification(admins, {
        type: "blind",
        title: "Bulk blinds created",
        message: `${clean.length} blinds added to ${(project?.name || "project")}.`,
        projectId: project.id,
        requiresAction: false,
        actorId: state.currentUser?.id || null,
      });
    }
  } catch(e) {}
saveState();
  closeModal("bulkAddBlindsModal");
  openProjectDetails(project.id);
  renderDashboard();
}

/* ==========================
   BLIND DETAILS + WORKFLOW UI
========================== */
function openBlindDetails(blindId) {
  const blind = state.blinds.find((b) => b.id === blindId);
  if (!blind) return;

  state.currentBlindId = blindId;

  // Open page first so navigation still works even if a render error happens.
  openPage("blindDetailsPage");

  const area = state.areas.find((a) => a.id === blind.areaId);
  const project = state.projects.find((p) => p.id === blind.projectId);

  document.getElementById("blindDetailsTitle").textContent = "Blind: " + blind.name;
  document.getElementById("blindDetailArea").textContent = area ? area.name : "-";
  document.getElementById("blindDetailProject").textContent = project ? project.name : "-";
  document.getElementById("blindDetailLine").textContent = blind.line || "-";
  document.getElementById("blindDetailType").textContent = blind.type || "-";
  document.getElementById("blindDetailSize").textContent = blind.size || "-";
  const rateEl = document.getElementById("blindDetailRate");
  if (rateEl) rateEl.textContent = blind.rate || "-";

  // QR (Public / Visitor link)
  const publicLink = buildPublicBlindUrl(blind.id);
  const linkEl = document.getElementById("blindQrLink");
  const imgEl = document.getElementById("blindQrImg");
  if (linkEl) linkEl.textContent = publicLink;
  if (imgEl) imgEl.src = buildQrImageUrl(publicLink);

  document.getElementById("blindPhaseBadge").textContent = phaseLabel(blind.phase || "broken");

  try {
    renderWorkflowSteps(blind);
    renderBlindHistory(blind);
    renderApprovalSummary(blind);
  } catch (err) {
    console.error("SBTS: failed to render blind details", err);
    const box = document.getElementById("workflowBox");
    if (box) {
      box.innerHTML = `<div class="card" style="padding:12px;border:1px solid #f0c8c8;background:#fff7f7;border-radius:12px;">` +
        `<b>Unable to load details</b><div style="margin-top:6px;">Open DevTools → Console to see the error.</div>` +
        `</div>`;
    }
  }
}

function findLastChangeToPhase(blind, phase) {
  const hist = blind.history || [];
  const matches = hist.filter((h) => h.toPhase === phase);
  if (matches.length === 0) return null;
  matches.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return matches[0];
}

function renderWorkflowSteps(blind) {
  const box = document.getElementById("workflowSteps");
  box.innerHTML = "";

  const cfg = getWorkflowConfig();
  const phaseObjs = (cfg.phases || []).slice().sort((a, b) => (a.order || 0) - (b.order || 0));
  const activePhaseObjs = phaseObjs.filter(p => p && p.active !== false && isPhaseEnabledForProject(blind.projectId, p.id));

  const ids = activePhaseObjs.map(p => p.id);
  // Patch40: If workflow was modified and new phases were inserted earlier,
  // we must NOT mark them as DONE unless there is a real history record.
  // For blinds that have no history at all, we reset the current phase to the new first phase.
  if (!Array.isArray(blind.history) || blind.history.length === 0) {
    if (ids[0] && blind.phase !== ids[0]) { blind.phase = ids[0]; try{saveState();}catch(e){} }
  }
  // If stored phase is no longer valid, fallback to first active phase.
  const currentId = (blind.phase && ids.includes(blind.phase)) ? blind.phase : (ids[0] || "broken");
  let currentIndex = ids.indexOf(currentId);
  if (currentId === "__final__") currentIndex = ids.length;
  if (currentIndex < 0) currentIndex = 0;

  const isPhaseActuallyDone = (phaseId) => {
    const hist = Array.isArray(blind.history) ? blind.history : [];
    // Consider a phase done only if there is a history record that started from it.
    // (i.e., user clicked Next at least once from that phase)
    return hist.some(h => h && h.type === 'phase' && h.fromPhase === phaseId);
  };

  activePhaseObjs.forEach((pObj, idx) => {
    const ph = pObj.id;
    const statusClass = idx < currentIndex ? (isPhaseActuallyDone(ph) ? "done" : "pending") : (idx === currentIndex ? "current" : "pending");

    const last = findLastChangeToPhase(blind, ph);
    const byText = last ? `${last.workerName || "-"} (${last.workerId || "-"})` : "-";
    const whenText = last?.date ? new Date(last.date).toLocaleString() : "-";

    const canUpdateRoles = Array.isArray(pObj.canUpdate) && pObj.canUpdate.length
      ? pObj.canUpdate.map(r => roleLabel(r)).join(", ")
      : roleLabel(WORKFLOW_REQUIRED_ROLE[ph] || "-");

    const approvalEnabled = !!(pObj.approval && pObj.approval.enabled);
    const approvalRoles = approvalEnabled
      ? (pObj.approval.roles || []).map(r => roleLabel(r)).join(", ") || "-"
      : "Not required";

    const extraEnabled = !!(pObj.extra && pObj.extra.enabled);
    const extraCount = extraEnabled ? ((pObj.extra.items || []).length) : 0;

    const div = document.createElement("div");
    div.className = `workflow-step ${statusClass}`;
    // Use per-phase color for CURRENT (DONE is always green, FUTURE is gray)
    div.style.setProperty("--phase-color", pObj.color || wfPhaseColor(ph));
    div.innerHTML = `
      <div class="workflow-left">
        <div class="workflow-bullet"></div>
        <div>
          <div class="workflow-title">${phaseLabel(ph)}</div>
          <div class="workflow-sub">
            Can update: ${canUpdateRoles}
            • Approval: ${approvalRoles}
            • Extra: ${extraEnabled ? (extraCount + " item(s)") : "None"}
            • Last: ${byText} • ${whenText}
          </div>
        </div>
      </div>
      <div class="workflow-status ${idx < currentIndex ? "done" : (idx === currentIndex ? "current" : "pending")}">
        ${idx < currentIndex ? "✓ DONE" : (idx === currentIndex ? "CURRENT" : "")}
      </div>
    `;
    box.appendChild(div);
  });

  // Final approvals pseudo-step (after all phases)
  const finalDiv = document.createElement("div");
  const finalDone = isAllFinalApprovalsDone(blind);
  const finalCurrent = (blind.phase === "__final__") && !finalDone;

  // progress (required only)
  const _finalList = getFinalApprovalsForBlind(blind) || [];
  const _finalRequired = _finalList.filter(a => (a && a.required !== false) && (a.status !== "disabled") && (a.status !== "archived"));
  // Count approved required final approvals using the approval's resolved key.
  // Use the blind's final approvals store (keyed by approval.key).
  const _faStore = blind.finalApprovals || {};
  const _finalApprovedCount = _finalRequired.filter(a => faIsIdApprovedForKey(_faStore, a.key || a.id)).length;
  const _finalTotal = _finalRequired.length;

  finalDiv.className = "workflow-step " + (finalDone ? "done" : (finalCurrent ? "current" : "pending"));

  // Final approvals color: Grey (0 approved) → Orange (partial) → Green (done)
  let _finalColor = "#6b7280"; // pending
  if (_finalTotal > 0 && _finalApprovedCount > 0 && !finalDone) _finalColor = "#f59e0b"; // in progress
  if (finalDone) _finalColor = "#16a34a"; // done
  finalDiv.style.setProperty("--phase-color", _finalColor);

  const _finalStage = finalDone ? "Completed" : (_finalApprovedCount > 0 ? "In progress" : "Pending");
  const _finalMeta = (_finalTotal === 0)
    ? "No required approvals configured • Managed in Final Approvals Manager"
    : `Required: ${_finalStage} • ${_finalApprovedCount}/${_finalTotal} approved • Managed in Final Approvals Manager`;

  // Match the standard workflow row layout (bullet aligned with other phases)
  finalDiv.innerHTML = `
    <div class="workflow-left">
      <div class="workflow-bullet"></div>
    </div>
    <div class="workflow-step-body">
      <div class="workflow-title">Final approvals</div>
      <div class="workflow-meta">${_finalMeta}</div>
    </div>
    <div class="workflow-status ${finalDone ? "done" : (finalCurrent ? (_finalApprovedCount > 0 ? "inprogress" : "current") : "pending")}">
      ${finalDone ? "✓ DONE" : (finalCurrent ? (_finalApprovedCount > 0 ? "IN PROGRESS" : "FINAL") : "PENDING")}
    </div>
  `;
  box.appendChild(finalDiv);


  // If no phases configured (edge case), fallback to legacy list
  if (!activePhaseObjs.length) {
    PHASES.forEach((ph) => {
      const div = document.createElement("div");
      div.className = "workflow-step pending";
      div.innerHTML = `<div class="workflow-step-body"><div class="workflow-title">${phaseLabel(ph)}</div></div>`;
      box.appendChild(div);
    });
  }
}


function renderBlindHistory(blind) {
  const tbody = document.getElementById("blindHistoryTableBody");
  tbody.innerHTML = "";

  const hist = (blind.history || []).slice().sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  hist.forEach((h) => {
    const tr = document.createElement("tr");
    const isFA = h.type === "finalApproval";
    const fromTxt = isFA ? "Final approvals" : phaseLabel(h.fromPhase || "-");
    const toTxt = isFA ? (h.approvalName || h.approvalKey || "-") : phaseLabel(h.toPhase || "-");
    const whoTxt = (h.workerName || "-") + " (" + (h.workerId || "-") + ")";
    tr.innerHTML = `
      <td>${h.date ? new Date(h.date).toLocaleString() : "-"}</td>
      <td>${fromTxt}</td>
      <td>${toTxt}</td>
      <td>${whoTxt}</td>
    `;
    tbody.appendChild(tr);
  });
}


function renderApprovalSummary(blind) {
  const el = document.getElementById("approvalSummaryText");
  const fa = blind.finalApprovals || {};
  const model = getFinalApprovalsModel(blind);
  const list = (model.required||[]);
  const lines = list.map((a) => {
    const ok = faIsIdApprovedForKey(fa, a.key);
    const rec = (fa[a.key] || null);
    // if approved via redirected old key
    let via = "";
    if(!rec){
      const redirects = faLoadRedirects();
      for(const oldKey in redirects){
        if(faResolveId(oldKey)===a.key && fa[oldKey] && fa[oldKey].status==="approved"){ via = ` (via ${oldKey})`; break; }
      }
    }
    const name = (rec?.name) || (()=>{ const redirects = faLoadRedirects(); for(const oldKey in redirects){ if(faResolveId(oldKey)===a.key && fa[oldKey]) return fa[oldKey].name||"-";} return "-"; })();
    const dt = (rec?.date) ? new Date(rec.date).toLocaleString() : "";
    return `${a.label}: ${ok ? "✅" : "⏳"} (${name}) ${ok ? dt : ""}${via}`;
  });
  const status = isCertificateApproved(blind) ? "APPROVED ✅" : "PENDING ⏳";
  el.textContent = `Certificate status: ${status} • ${lines.join(" • ")}`;
}

/* ==========================
   PHASE CHANGE
========================== */
/**
 * Open custom confirm modal for phase change.
 * payload: { blindId, toPhase, workerName, fromPhase, closeAfter: [modalIds], defaultSigId }
 */
function openPhaseConfirmModal(payload) {
  const blind = state.blinds.find((b) => b.id === payload.blindId);
  if (!blind) return;

  phaseConfirmPayload = {
    blindId: payload.blindId,
    toPhase: payload.toPhase,
    fromPhase: payload.fromPhase || (blind.phase || "broken"),
    workerName: payload.workerName || (state.currentUser?.fullName || "-"),
    defaultSigId: payload.defaultSigId || (state.currentUser?.username || ""),
    closeAfter: payload.closeAfter || [],
  };

  // Fill UI
  document.getElementById("pc_blindName").textContent = blind.name || "-";
  document.getElementById("pc_fromPhase").textContent = phaseLabel(phaseConfirmPayload.fromPhase);
  document.getElementById("pc_toPhase").textContent = phaseLabel(phaseConfirmPayload.toPhase);
  document.getElementById("pc_workerName").textContent = phaseConfirmPayload.workerName;

  const sig = document.getElementById("pc_signatureId");
  sig.value = phaseConfirmPayload.defaultSigId || "";
  openModal("phaseConfirmModal");
  setTimeout(() => sig.focus(), 50);
}

function cancelPhaseConfirm() {
  phaseConfirmPayload = null;
  closeModal("phaseConfirmModal");
}

function confirmPhaseConfirm() {
  if (!phaseConfirmPayload) return;

  const blind = state.blinds.find((b) => b.id === phaseConfirmPayload.blindId);
  if (!blind) return cancelPhaseConfirm();

  const sigId = document.getElementById("pc_signatureId").value.trim();
  if (!sigId) return alert("ID is required.");

  applyPhaseChange(blind, phaseConfirmPayload.toPhase, phaseConfirmPayload.workerName, sigId);

  closeModal("phaseConfirmModal");
  (phaseConfirmPayload.closeAfter || []).forEach((m) => closeModal(m));
  phaseConfirmPayload = null;
}

function openChangePhaseModal() {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  if (!canUser("changePhases") && state.currentUser.role !== "admin") return alert("No permission.");

  const sel = document.getElementById("changePhaseSelect");
  sel.innerHTML = "";

  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
    const allowed = canTransition(blind.phase, ph, state.currentUser, blind);
    if (allowed || ph === blind.phase || state.currentUser.role === "admin") {
      const opt = document.createElement("option");
      opt.value = ph;
      opt.textContent = phaseLabel(ph);
      sel.appendChild(opt);
    }
  });

  sel.value = nextPhaseOf(blind.phase || "broken");
  document.getElementById("changePhaseTechName").value = state.currentUser?.fullName || "";
  document.getElementById("changePhaseTechId").value = state.currentUser?.username || "";
  openModal("changePhaseModal");
}

function quickNextPhase() {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  const to = nextPhaseOf(blind.phase || "broken");
  if (!canTransition(blind.phase, to, state.currentUser, blind) && state.currentUser.role !== "admin") {
    const policy = getProjectUpdatePolicy(blind?.projectId);
    const allowedRoles = getAllowedRolesForPhase(to);
    const legacyReq = WORKFLOW_REQUIRED_ROLE[to];
    const roleList = allowedRoles.length ? allowedRoles : (legacyReq ? [legacyReq] : []);

    if (policy === "owners") {
      return alert(`Not allowed. You must be assigned as an owner for "${phaseLabel(to)}" in this project. Go to Project Settings → Phase Ownership.`);
    }

    if (policy === "roles") {
      if (roleList.length) return alert(`Not allowed. Required role for "${phaseLabel(to)}" is ${roleList.map(roleLabel).join(" / ")}.`);
      return alert(`Not allowed. Role policy is enabled but no allowed roles are configured for "${phaseLabel(to)}". Ask admin to set canUpdate in Workflow Control.`);
    }

    // Hybrid: Owner OR Role
    if (roleList.length) {
      return alert(`Not allowed. You must be an owner OR have role ${roleList.map(roleLabel).join(" / ")} for "${phaseLabel(to)}".`);
    }
    return alert(`Not allowed. You must be assigned as an owner for "${phaseLabel(to)}" in this project. Go to Project Settings → Phase Ownership.`);
  }

  openPhaseConfirmModal({
    blindId: blind.id,
    fromPhase: blind.phase || "broken",
    toPhase: to,
    workerName: state.currentUser?.fullName || "-",
    defaultSigId: state.currentUser?.username || "",
    closeAfter: [], // no other modals to close
  });
}

function confirmChangePhase() {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  const newPhase = document.getElementById("changePhaseSelect").value;
  const workerName = document.getElementById("changePhaseTechName").value.trim();
  const workerId = document.getElementById("changePhaseTechId").value.trim();

  if (!newPhase) return alert("Select new phase.");
  if (!workerName || !workerId) return alert("Enter worker name and ID.");

  if (!canTransition(blind.phase, newPhase, state.currentUser, blind) && state.currentUser.role !== "admin") {
    const req = WORKFLOW_REQUIRED_ROLE[newPhase];
    if (req) return alert(`Not allowed. Required role for "${phaseLabel(newPhase)}" is ${roleLabel(req)}.`);
    return alert(`Not allowed. You must be assigned as an owner for phase "${phaseLabel(newPhase)}" in this project. Go to Project Settings → Phase Ownership.`);
  }

  // Use the custom modal for confirmation + signature
  openPhaseConfirmModal({
    blindId: blind.id,
    fromPhase: blind.phase || "broken",
    toPhase: newPhase,
    workerName: workerName,
    defaultSigId: workerId,
    closeAfter: ["changePhaseModal"],
  });
}

function applyPhaseChange(blind, newPhase, workerName, workerId) {
  const now = new Date().toISOString();
  const fromPhase0 = blind.phase || null;

  // NOTE: we do NOT write a history record when moving into Final approvals (__final__).
  // We only log the actual Final approval actions themselves.
  if (newPhase !== "__final__") {
    // Keep only one history record per target phase (latest wins)
    blind.history = (blind.history || []).filter(h => h.toPhase !== newPhase);
    blind.history.push({
      date: now,
      type: "phase",
      fromPhase: blind.phase || null,
      toPhase: newPhase,
      workerName,
      workerId,
      userId: state.currentUser?.id || null,
      role: state.currentUser?.role || null,
    });
  }

  const fromIdx = PHASES.indexOf(blind.phase || "broken");
  const toIdx = newPhase === "__final__" ? PHASES.length : PHASES.indexOf(newPhase);

  blind.phase = newPhase;
  try { notifyOwnersOnPhaseChange(blind, fromPhase0, newPhase); } catch(e) {}

  // Info notification to admins (audit visibility) - local only
  try {
    const actorId = state.currentUser?.id || null;
    const admins = getAdminUserIds().filter(uid => uid && uid !== actorId);
    if (admins.length) {
      const pid = blind.projectId || state.currentProjectId || null;
      const fromLbl = fromPhase0 ? phaseLabel(fromPhase0) : "-";
      const toLbl = newPhase ? phaseLabel(newPhase) : "-";
      SBTS_ACTIVITY.pushNotification({
  category: "system",
  scope: "project",
  projectId: pid,
  title: "Phase updated",
  message: `${sbtsBlindOfficialText(blind)} — ${fromLbl} → ${toLbl}`,
  blindId: blind.id,
  requiresAction: false,
  resolved: true,
  actorId: actorId
});
    }
  } catch(e) {}


  // if moved backward (admin), reset approvals
  if (state.currentUser?.role === "admin" && toIdx < fromIdx) {
    blind.finalApprovals = {};
  }
  // if not fully complete (i.e., not in __final__), approvals reset
  if (!isAllPhasesComplete(blind)) {
    blind.finalApprovals = {};
  }

  saveState();
  openBlindDetails(blind.id);
  renderDashboard();
}


/* ==========================
   FINAL APPROVALS
========================== */
function openApprovalsModal() {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  if (!isAllPhasesComplete(blind)) {
    return alert("Finish all phases first (reach the last phase).");
  }

  // Build dynamic approval rows from Workflow Control (last active phase)
  const box = document.getElementById("finalApprovalsBox");
  if (box) {
    const model = getFinalApprovalsModel(blind);
    const required = model.required || [];
    const optional = model.optional || [];
    const fa = blind.finalApprovals || {};

    const renderRow = (a, isOptional=false) => {
      const done = fa[a.key] && fa[a.key].status === "approved";
      const roleLabel = ROLE_LABELS[a.role] || a.role;
      const canApprove = (state.currentUser?.role === "admin") || (state.currentUser?.role === a.role) || (Array.isArray(a.roles) && a.roles.includes(state.currentUser?.role));
      const btnDisabled = done || !canApprove;
      const meta = done ? `<div class="tiny">Approved by: ${fa[a.key].by} • ${fa[a.key].date}</div>` : `<div class="tiny">Role: ${roleLabel}${isOptional ? " • Optional" : ""}</div>`;
      return `
        <div class="approval-row ${done ? "approved" : ""}">
          <div><b>${faLookupName((a.key||"").toUpperCase(), a.label)}</b>${meta}</div>
          <button class="btn-neutral" ${btnDisabled ? "disabled" : ""} onclick="approveStep('${a.key}')">${done ? "Approved" : "Approve"}</button>
        </div>`;
    };

    const reqHtml = required.length ? required.map(a=>renderRow(a,false)).join("") : `<div class="tiny">No required approvals configured.</div>`;
    const optHtml = optional.length ? `<div class="tiny" style="margin-top:10px;"><b>Optional approvals</b></div>` + optional.map(a=>renderRow(a,true)).join("") : "";

    box.innerHTML = reqHtml + optHtml;
  }

  openModal("approvalsModal");
}

function approveStep(stepKey) {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  if (!isAllPhasesComplete(blind)) return alert("Finish all phases first.");

  const model = getFinalApprovalsModel(blind);
  const all = (model.required || []).concat(model.optional || []);
  const def = all.find((x) => x.key === stepKey);
  if (!def) return;

  const userRole = state.currentUser?.role;
  const canApprove = (userRole === "admin") || (userRole === def.role) || (Array.isArray(def.roles) && def.roles.includes(userRole));
  if (!canApprove) {
    return alert(`Not allowed. Required role: ${roleLabel(def.role)}`);
  }

  blind.finalApprovals = blind.finalApprovals || {};
  blind.finalApprovals[stepKey] = {
    status: "approved",
    name: state.currentUser.fullName,
    username: state.currentUser.username,
    role: userRole,
    date: formatDateDDMMYYYY(new Date())
  };

  // Log the approval itself (instead of logging a fake "Final phase")
  blind.history = blind.history || [];
  blind.history.push({
    date: new Date().toISOString(),
    type: "finalApproval",
    approvalKey: stepKey,
    approvalName: def.name,
    workerName: state.currentUser.fullName,
    workerId: state.currentUser.username,
    userId: state.currentUser.id || null,
    role: userRole || null,
  });

  saveState();

  // Notify admins (audit trail)
  try {
    const users = Array.isArray(state.users) ? state.users : [];
    const actorId = state.currentUser?.id || null;
    const admins = users.filter(u=>u.role==="admin" && u.id && u.id!==actorId).map(u=>u.id);
    if (admins.length) {
      const pid = blind.projectId || state.currentProjectId || null;
      addNotification(admins, {
        type: "approval",
        title: "Final approval updated",
        message: `Blind ${blind.name || blind.id}: ${def.name} approved by ${state.currentUser.fullName || state.currentUser.username}.`,
        projectId: pid,
        blindId: blind.id,
        requiresAction: false,
      });
    }
    if (admins.length && isCertificateApproved(blind)) {
      const pid = blind.projectId || state.currentProjectId || null;
      addNotification(admins, {
        type: "certificate",
        title: "Certificate approved",
        message: `Blind ${blind.name || blind.id} is now fully approved and ready to print.`,
        projectId: pid,
        blindId: blind.id,
        requiresAction: false,
      });
    }
  } catch(e) {}


  // re-render modal rows + summary
  if (document.getElementById("finalApprovalsBox")) openApprovalsModal();
  renderApprovalSummary(blind);
  // update workflow final row immediately
  if (document.getElementById('workflowSteps')) renderWorkflowSteps(blind);
  // update history table immediately if visible
  if (document.getElementById('blindHistoryTableBody')) renderBlindHistory(blind);

  if (isCertificateApproved(blind)) alert("✅ Certificate is now APPROVED.");
  else alert("Approval saved.");
}

/* ==========================
   CERTIFICATE (RIGHT COMPANY LOGO, LEFT PROGRAM, STATUS CENTER)
========================== */
function getCertificateStyles() {
  const cfg = state.certificate;
  return `
<style>
  @page { size: A4; margin: 12mm; }
  body { font-family: system-ui, -apple-system, "Segoe UI", sans-serif; background:#fff; color:#111827; }
  .cert { border: 2px solid #111827; border-radius: 18px; padding: 14px 16px; }
  .top {
    display:flex; align-items:center; justify-content:space-between;
    border-bottom: 2px solid #111827; padding-bottom: 10px; margin-bottom: 12px;
    background:${cfg.headerBg || "#fff"};
    border-radius: 14px;
    padding: 12px 12px 10px;
  }
  .brandL, .brandR { display:flex; align-items:center; gap:10px; width: 33%; }
  .brandR { justify-content:flex-end; }
  .logoCircle {
    width:44px; height:44px; border-radius:999px; background:#174c7e; color:#fff;
    display:flex; align-items:center; justify-content:center; font-weight:900;
  }
  .companyLogo {
    width:46px; height:46px; border-radius:14px; background:#e5e7eb;
    background-size:cover; background-position:center;
  }
  .center {
    width:34%;
    text-align:center;
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    gap:6px;
  }
  .cert-title { font-weight:1000; font-size:16px; }
  .statusBig { font-weight:1100; letter-spacing:0.12em; font-size:18px; }
  .pill {
    border:1px solid #e5e7eb; border-radius:999px; padding:6px 10px; font-size:12px; background:#f9fafb;
  }
  .pending { color:#b45309; }
  .approved { color:#16a34a; }
  .grid { display:grid; grid-template-columns:repeat(2, minmax(0,1fr)); gap:8px; }
  table { width:100%; border-collapse: collapse; margin-top:10px; font-size:12px; }
  th, td { border-bottom:1px solid #e5e7eb; padding:6px 6px; text-align:left; }
  th { background:#f3f4f6; font-size:11px; color:#374151; }
  .foot { margin-top:10px; font-size:10px; color:#6b7280; display:flex; justify-content:space-between; gap:10px; }
</style>
`;
}

function openCertificatePage() {
  const blind = state.blinds.find((b) => b.id === state.currentBlindId);
  if (!blind) return;

  const area = state.areas.find((a) => a.id === blind.areaId);
  const project = state.projects.find((p) => p.id === blind.projectId);

  const hist = (blind.history || []).slice().sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const fa = blind.finalApprovals || {};
  const statusApproved = isCertificateApproved(blind);

  const approvalsRows = getFinalApprovalsForBlind(blind).map((a) => {
    const x = fa[a.key];
    const ok = x?.status === "approved";
    return `
      <tr>
        <td>${a.label}</td>
        <td>${ok ? "YES" : "NO"}</td>
        <td>${ok ? (x.name || "-") : "-"}</td>
        <td>${ok ? new Date(x.date).toLocaleString() : "-"}</td>
      </tr>
    `;
  }).join("");

  const cfg = state.certificate;
  const b = state.branding;

  const statusHtml = cfg.statusStyle === "pill"
    ? `<div class="pill ${statusApproved ? "approved" : "pending"}">${statusApproved ? "APPROVED" : "PENDING"}</div>`
    : `<div class="statusBig ${statusApproved ? "approved" : "pending"}">${statusApproved ? "APPROVED" : "PENDING"}</div>`;

  const win = window.open("", "_blank");
  win.document.write(`<html><head><title>Certificate - ${blind.name}</title>${getCertificateStyles()}</head><body>`);
  win.document.write(`
    <div class="cert">
      <div class="top">
        <!-- LEFT: Program -->
        <div class="brandL">
          <div class="logoCircle">SB</div>
          <div>
            <div style="font-weight:900;">${(b.programTitle || "").trim() || ""}</div>
            <div style="font-size:11px;color:#374151;">${(b.programSubtitle || "").trim() || ""}</div>
          </div>
        </div>

        <!-- CENTER: Status -->
        <div class="center">
          <div class="cert-title">${cfg.title || "Smart Blind Tag System Certificate"}</div>
          ${statusHtml}
        </div>

        <!-- RIGHT: Company -->
        <div class="brandR">
          <div style="text-align:right;">
            <div style="font-weight:900;">${(b.companyName || "").trim() || ""}</div>
            <div style="font-size:11px;color:#374151;">${(b.companySub || "").trim() || ""}</div>
          </div>
          <div class="companyLogo" style="${b.companyLogo ? `background-image:url(${b.companyLogo});` : ""}"></div>
        </div>
      </div>

      <div class="grid">
        <div class="pill"><b>Area:</b> ${area ? area.name : "-"}</div>
        <div class="pill"><b>Project:</b> ${project ? project.name : "-"}</div>
        <div class="pill"><b>Blind:</b> ${blind.name || "-"}</div>
        <div class="pill"><b>Line:</b> ${blind.line || "-"}</div>
        <div class="pill"><b>Type:</b> ${blind.type || "-"}</div>
        <div class="pill"><b>Size:</b> ${blind.size || "-"}</div>
        <div class="pill"><b>Current phase:</b> ${phaseLabel(blind.phase || "broken")}</div>
      </div>

      ${cfg.showWorkflow ? `
      <h3 style="margin:12px 0 6px;">Workflow log</h3>
      <table>
        <thead><tr><th>Date</th><th>From</th><th>To</th><th>Worker</th><th>Role</th></tr></thead>
        <tbody>
          ${
            hist.length === 0
              ? `<tr><td colspan="5">No changes yet.</td></tr>`
              : hist.map(h => `
                <tr>
                  <td>${h.date ? new Date(h.date).toLocaleString() : "-"}</td>
                  <td>${phaseLabel(h.fromPhase || "-")}</td>
                  <td>${phaseLabel(h.toPhase || "-")}</td>
                  <td>${h.workerName || "-"} (${h.workerId || "-"})</td>
                  <td>${roleLabel(h.role || "-")}</td>
                </tr>
              `).join("")
          }
        </tbody>
      </table>` : ""}

      ${cfg.showApprovals ? `
      <h3 style="margin:12px 0 6px;">Final approvals</h3>
      <table>
        <thead><tr><th>Approval</th><th>Approved</th><th>By</th><th>Date</th></tr></thead>
        <tbody>${approvalsRows}</tbody>
      </table>` : ""}

      <div class="foot">
        <div>${cfg.footerText || ""}</div>
        <div>Generated from SBTS local data.</div>
      </div>
    </div>
  `);
  win.document.write(`</body></html>`);
  win.document.close();
  win.print();
}

function generateProjectCertificates() {
  const projectId = state.currentProjectId;
  if (!projectId) return;

  const blinds = state.blinds.filter((b) => b.projectId === projectId);
  if (blinds.length === 0) return alert("No blinds in this project.");

  // Better UX: open ALL certificates in ONE printable window (approved + pending)
  // This avoids popup blockers and lets the user print everything for the project.
  openPrintCertsPicker(projectId);
}


function printProjectTagsFromProjectPage(){
  const projectId = state.currentProjectId;
  if (!projectId) return;
  const project = state.projects.find(p=>p.id===projectId);
  if (!project) return;
  const blinds = state.blinds.filter(b=>b.projectId===projectId);
  if (blinds.length===0) return alert("No blinds in this project.");
  ensureTagThemeDefaults();
  ensureTagTemplateDefaults();
  const cards = blinds.map(b=>({
    area: (state.areas.find(a=>a.id===b.areaId)?.name) || "",
    line: b.line || "",
    blindId: b.name || "",
    qrValue: buildPublicBlindUrl(b.id),
    publicLink: buildPublicBlindUrl(b.id),
  }));
  openPrintWindow(cards);
}

function duplicateCurrentProject(){
  if (!canUser("manageProjects") && state.currentUser.role !== "admin") {
    return alert("No permission.");
  }
  const srcId = state.currentProjectId;
  const src = state.projects.find(p=>p.id===srcId);
  if (!src) return;
  const suggested = (src.name||"Project") + " (Copy)";
  const name = prompt("New project name:", suggested);
  if (!name) return;
  const newId = crypto.randomUUID();
  const clone = { ...src, id: newId, name };
  state.projects.push(clone);

  // duplicate blinds (structure only)
  const srcBlinds = state.blinds.filter(b=>b.projectId===srcId);
  srcBlinds.forEach(b=>{
    const nb = {
      ...b,
      id: crypto.randomUUID(),
      projectId: newId,
      areaId: clone.areaId,
      phase: "broken",
      history: [],
      finalApprovals: {},
    };
    state.blinds.push(nb);
  });

  saveState();

  // Notify admins
  try {
    const users = Array.isArray(state.users) ? state.users : [];
    const actorId = state.currentUser?.id || null;
    const admins = users.filter(u=>u.role==="admin" && u.id && u.id!==actorId).map(u=>u.id);
    if (admins.length) {
      addNotification(admins, {
        type: "system",
        title: "Project duplicated",
        message: `Project "${name}" was created from "${src.name}" by ${state.currentUser?.fullName || state.currentUser?.username || "user"}.`,
        projectId: newId,
        requiresAction: false,
      });
    }
  } catch(e) {}

  toast("Project duplicated.");
  // Notify admins
  try {
    const users = Array.isArray(state.users) ? state.users : [];
    const actorId = state.currentUser?.id || null;
    const admins = users.filter(u=>u.role==="admin" && u.id && u.id!==actorId).map(u=>u.id);
    if (admins.length) {
      addNotification(admins, {
        type: "system",
        title: "Project duplicated",
        message: `New project "${name}" was created from "${src.name}".`,
        projectId: newId,
        requiresAction: false,
      });
    }
  } catch(e) {}

  openProjectDetails(newId);
  renderDashboard();
}

function openPrintCertsPicker(projectId) {
  const project = state.projects.find((p) => p.id === projectId);
  const blindsAll = state.blinds
    .filter((b) => b.projectId === projectId)
    .slice();

  if (!project) return alert("Project not found.");
  if (blindsAll.length === 0) return alert("No blinds in this project.");

  // Cache for modal actions & filtering
  state._printCertsProjectId = projectId;
  state._printCertsAllBlinds = blindsAll;

  // Default filter state
  state._printCertsFilter = state._printCertsFilter || { q: "", view: "all", sort: "name_asc" };

  const listEl = document.getElementById("printCertsList");
  const sumEl  = document.getElementById("printCertsSummary");
  const searchEl = document.getElementById("printCertsSearch");
  const viewEl   = document.getElementById("printCertsView");
  const sortEl   = document.getElementById("printCertsSort");

  if (!listEl || !sumEl) {
    // Fallback: print all if modal is missing for any reason
    openCombinedProjectCertificates(projectId);
    return;
  }

  // Summary
  const approvedCount = blindsAll.filter(isCertificateApproved).length;
  const pendingCount  = blindsAll.length - approvedCount;

  sumEl.innerHTML = `
    <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
      <div><b>${escapeHtml(project.name || "Project")}</b></div>
      <span class="pill" style="background:#e5e7eb; color:#111827;">Total: ${blindsAll.length}</span>
      <span class="pill" style="background:#dcfce7; color:#166534;">Approved: ${approvedCount}</span>
      <span class="pill" style="background:#ffedd5; color:#9a3412;">Pending: ${pendingCount}</span>
    </div>
    <div class="tiny muted" style="margin-top:6px;">Select the certificates you want to print, then click “Print selected”.</div>
  `;

  // Bind filter controls (once per open)
  if (searchEl) searchEl.value = state._printCertsFilter.q || "";
  if (viewEl)   viewEl.value   = state._printCertsFilter.view || "all";
  if (sortEl)   sortEl.value   = state._printCertsFilter.sort || "name_asc";

  const onFilterChanged = () => {
    state._printCertsFilter.q    = (searchEl ? searchEl.value : "").trim();
    state._printCertsFilter.view = (viewEl ? viewEl.value : "all");
    state._printCertsFilter.sort = (sortEl ? sortEl.value : "name_asc");
    pcRenderPrintCertsList();
  };

  if (searchEl) {
    searchEl.oninput = onFilterChanged;
  }
  if (viewEl) {
    viewEl.onchange = onFilterChanged;
  }
  if (sortEl) {
    sortEl.onchange = onFilterChanged;
  }

  // Initial render
  pcRenderPrintCertsList();

  // Apply user's last maximize choice (local preference)
  try {
    const wantMax = SBTS_UTILS.LS.get("sbts_printCerts_max") === "1";
    const modal = document.getElementById("printCertsModal");
    const content = modal?.querySelector(".modal-content");
    if (content) content.classList.toggle("pc-maximized", wantMax);
  } catch (_) {}

  openModal("printCertsModal");
}

// Patch 9: Maximize / Restore Print Certificates window (user preference)
function pcTogglePrintCertsMax() {
  const modal = document.getElementById("printCertsModal");
  const content = modal?.querySelector(".modal-content");
  if (!content) return;
  const next = !content.classList.contains("pc-maximized");
  content.classList.toggle("pc-maximized", next);
  try {
    SBTS_UTILS.LS.set("sbts_printCerts_max", next ? "1" : "0");
  } catch (_) {}
}


// Patch 8: Print picker filtering / sorting (Search + View + Sort)
function pcGetCurrentCheckedIds() {
  const wrap = document.getElementById("printCertsList");
  const set = new Set();
  if (!wrap) return set;
  wrap.querySelectorAll('input.pcCheck[type="checkbox"]').forEach((cb) => {
    if (cb.checked) set.add(cb.getAttribute("data-id"));
  });
  return set;
}

function pcApplyPrintCertsFilter(blinds) {
  const f = state._printCertsFilter || { q: "", view: "all", sort: "name_asc" };
  const q = String(f.q || "").toLowerCase();

  let out = blinds.slice();

  // View filter (Approved/Pending)
  if (f.view === "approved") {
    out = out.filter((b) => isCertificateApproved(b));
  } else if (f.view === "pending") {
    out = out.filter((b) => !isCertificateApproved(b));
  }

// Extra filters (in addition to left tabs)
if (unreadOnly) list = list.filter(n => !n.read);
if (projectFilter !== "all") list = list.filter(n => n.projectId === projectFilter);
if (typeFilter !== "all") {
  if (typeFilter === "action") list = list.filter(n => n.requiresAction && !n.resolved);
  else list = list.filter(n => notifKind(n) === typeFilter);
}

  // Search by name (and allow searching by id/type/size lightly)
  if (q) {
    out = out.filter((b) => {
      const name = String(b.name || "").toLowerCase();
      const id   = String(b.id || "").toLowerCase();
      const type = String(b.type || "").toLowerCase();
      const size = String(b.size || "").toLowerCase();
      return name.includes(q) || id.includes(q) || type.includes(q) || size.includes(q);
    });
  }

  // Sorting
  const sort = f.sort || "name_asc";
  if (sort === "name_desc") {
    out.sort((a, b) => String(b.name || "").localeCompare(String(a.name || "")));
  } else if (sort === "status") {
    // Approved first, then by name
    out.sort((a, b) => {
      const aa = isCertificateApproved(a) ? 0 : 1;
      const bb = isCertificateApproved(b) ? 0 : 1;
      if (aa !== bb) return aa - bb;
      return String(a.name || "").localeCompare(String(b.name || ""));
    });
  } else if (sort === "phase") {
    out.sort((a, b) => String(phaseLabel(a.phase || "broken")).localeCompare(String(phaseLabel(b.phase || "broken"))) || String(a.name || "").localeCompare(String(b.name || "")));
  } else {
    // name_asc default
    out.sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));
  }

  return out;
}

function pcRenderPrintCertsList() {
  const listEl = document.getElementById("printCertsList");
  if (!listEl) return;

  const all = state._printCertsAllBlinds || [];
  const keepChecked = pcGetCurrentCheckedIds();

  const blinds = pcApplyPrintCertsFilter(all);

  listEl.innerHTML = blinds.map((b) => {
    const ok = isCertificateApproved(b);
    const badge = ok
      ? `<span class="pill" style="background:#dcfce7; color:#166534;">Approved</span>`
      : `<span class="pill" style="background:#ffedd5; color:#9a3412;">Pending</span>`;

    const checked = keepChecked.has(String(b.id)) ? "checked" : "";
    return `
      <label class="pc-item">
        <input type="checkbox" class="pcCheck pc-checkbox" data-id="${b.id}" ${checked} />
        <div class="pc-content">
          <div class="pc-top">
            <div class="pc-title" title="${escapeHtml(b.name || "Blind")}">${escapeHtml(b.name || "Blind")}</div>
            ${badge}
          </div>
          <div class="pc-meta tiny muted">
            <span>Phase: <b>${escapeHtml(phaseLabel(b.phase || "broken"))}</b></span>
            <span>Type: <b>${escapeHtml(b.type || "-")}</b></span>
            <span>Size: <b>${escapeHtml(b.size || "-")}</b></span>
          </div>
        </div>
      </label>
    `;
  }).join("");

  // Update the "Tip" / empty state if nothing matched
  if (blinds.length === 0) {
    listEl.innerHTML = `
      <div style="padding:14px 12px; border:1px dashed #e5e7eb; border-radius:14px; color:#6b7280;">
        No certificates match your search/filter.
      </div>
    `;
  }
}

function pcSelectAll(val) {
  document.querySelectorAll("#printCertsList .pcCheck").forEach((c) => { c.checked = !!val; });
}
function pcSelectApproved() {
  const pid = state._printCertsProjectId;
  document.querySelectorAll("#printCertsList .pcCheck").forEach((c) => {
    const b = state.blinds.find((x) => x.id === c.dataset.id && x.projectId === pid);
    c.checked = b ? isCertificateApproved(b) : false;
  });
}
function pcSelectPending() {
  const pid = state._printCertsProjectId;
  document.querySelectorAll("#printCertsList .pcCheck").forEach((c) => {
    const b = state.blinds.find((x) => x.id === c.dataset.id && x.projectId === pid);
    c.checked = b ? !isCertificateApproved(b) : false;
  });
}

function pcPrintSelected() {
  const projectId = state._printCertsProjectId;
  const selected = [...document.querySelectorAll("#printCertsList .pcCheck")]
    .filter((c) => c.checked)
    .map((c) => c.dataset.id);

  if (!projectId) return;
  if (selected.length === 0) {
    alert("Please select at least one certificate to print.");
    return;
  }

  closeModal("printCertsModal");
  openCombinedProjectCertificates(projectId, selected);
}


function openCombinedProjectCertificates(projectId, selectedIds) {
  const project = state.projects.find((p) => p.id === projectId);
  const areaName = project ? (state.areas.find((a) => a.id === project.areaId)?.name || "-") : "-";

  // Build the candidate list (all blinds in project)
  let blinds = state.blinds
    .filter((b) => b.projectId === projectId)
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));

  // If a selection is provided, print ONLY the selected certificates
  if (Array.isArray(selectedIds) && selectedIds.length) {
    const set = new Set(selectedIds.map((x) => String(x)));
    blinds = blinds.filter((b) => set.has(String(b.id)));
  }

  
if (blinds.length === 0) return alert("No blinds in this project.");

  // Certificate issued tracking (once per blind, local-only)
  try {
    const actorId = state.currentUser?.id || null;
    const admins = getAdminUserIds().filter(uid => uid && uid !== actorId);
    blinds.forEach((b0) => {
      if (!b0) return;
      if (!isCertificateApproved(b0)) return;
      if (b0.certificateIssuedAt) return;
      b0.certificateIssuedAt = new Date().toISOString();
      b0.history = b0.history || [];
      b0.history.push({
        date: b0.certificateIssuedAt,
        type: "certificateIssued",
        workerName: state.currentUser?.fullName || state.currentUser?.username || "User",
        workerId: state.currentUser?.username || "",
        userId: actorId,
        role: state.currentUser?.role || null,
      });

      if (admins.length) {
        SBTS_ACTIVITY.pushNotification({
  category: "system",
  scope: "project",
  projectId: projectId,
  title: "Certificate issued",
  message: `${sbtsBlindOfficialText(b0)} — certificate generated.`,
  blindId: b0.id,
  requiresAction: false,
  resolved: true,
  actorId
});
      }
      // Inform the current user too (confirmation)
      SBTS_ACTIVITY.pushNotification({
  category: "system",
  scope: "user",
  recipients: [getCurrentUserIdStable()].filter(Boolean),
  projectId: projectId,
  title: "Certificate generated",
  message: `${sbtsBlindOfficialText(b0)} — ready for printing.`,
  blindId: b0.id,
  requiresAction: false,
  resolved: true,
  actorId
});
    });
    saveState();
    try { renderDashboardRecentActivity(); } catch(_) {}
    try { updateNotificationsBadge(); } catch(_) {}
  } catch(_) {}

  const cfg = state.certificate;
  const b = state.branding;

  const win = window.open("", "_blank");
  if (!win) {
    alert("Popup blocked. Please allow popups to print project certificates.");
    return;
  }

  // Reuse the exact same certificate styles
  win.document.write(`<html><head><title>Project Certificates - ${(project?.name || "Project")}</title>${getCertificateStyles()}</head><body>`);

  // Simple cover / index section
  win.document.write(`
    <div style="margin:0 0 14px; padding:10px 12px; border:1px solid #e5e7eb; border-radius:14px;">
      <div style="display:flex; align-items:flex-start; justify-content:space-between; gap:12px;">
        <div>
          <div style="font-weight:900; font-size:16px;">Project certificates</div>
          <div style="font-size:12px; color:#374151; margin-top:4px;">Area: <b>${areaName}</b> &nbsp; | &nbsp; Project: <b>${project?.name || "-"}</b></div>
          <div style="font-size:12px; color:#6b7280; margin-top:2px;">Total: <b>${blinds.length}</b> certificate(s)${(Array.isArray(selectedIds) && selectedIds.length) ? " (Selected)" : ""}</div>
        </div>
        <div style="text-align:right; font-size:12px; color:#6b7280;">
          <div>Generated from SBTS local data.</div>
          <div>${new Date().toLocaleString()}</div>
        </div>
      </div>
      <div style="margin-top:10px; font-size:12px; color:#111827;">
        ${(Array.isArray(selectedIds) && selectedIds.length)
          ? "Tip: You are printing only the selected certificates."
          : "Tip: Use your printer settings to select pages if you don’t want to print all certificates."}
      </div>
    </div>
  `);

  // Render each certificate with a page break
  blinds.forEach((blind, idx) => {
    const area = state.areas.find((a) => a.id === blind.areaId);
    const proj = state.projects.find((p) => p.id === blind.projectId);

    const hist = (blind.history || []).slice().sort((x, y) => (x.date || "").localeCompare(y.date || ""));
    const fa = blind.finalApprovals || {};
    const statusApproved = isCertificateApproved(blind);

    const approvalsRows = getFinalApprovalsForBlind(blind).map((a) => {
      const x = fa[a.key];
      const ok = x?.status === "approved";
      return `
        <tr>
          <td>${a.label}</td>
          <td>${ok ? "YES" : "NO"}</td>
          <td>${ok ? (x.name || "-") : "-"}</td>
          <td>${ok ? new Date(x.date).toLocaleString() : "-"}</td>
        </tr>
      `;
    }).join("");

    const statusHtml = cfg.statusStyle === "pill"
      ? `<div class="pill ${statusApproved ? "approved" : "pending"}">${statusApproved ? "APPROVED" : "PENDING"}</div>`
      : `<div class="statusBig ${statusApproved ? "approved" : "pending"}">${statusApproved ? "APPROVED" : "PENDING"}</div>`;

    win.document.write(`
      <div class="cert">
        <div class="top">
          <div class="brandL">
            <div class="logoCircle">SB</div>
            <div>
              <div style="font-weight:900;">${(b.programTitle || "").trim() || ""}</div>
              <div style="font-size:11px;color:#374151;">${(b.programSubtitle || "").trim() || ""}</div>
            </div>
          </div>
          <div class="center">
            <div class="cert-title">${cfg.title || "Smart Blind Tag System Certificate"}</div>
            ${statusHtml}
          </div>
          <div class="brandR">
            <div style="text-align:right;">
              <div style="font-weight:900;">${(b.companyName || "").trim() || ""}</div>
              <div style="font-size:11px;color:#374151;">${(b.companySub || "").trim() || ""}</div>
            </div>
            <div class="companyLogo" style="${b.companyLogo ? `background-image:url(${b.companyLogo});` : ""}"></div>
          </div>
        </div>

        <div class="grid">
          <div class="pill"><b>Area:</b> ${area ? area.name : "-"}</div>
          <div class="pill"><b>Project:</b> ${proj ? proj.name : "-"}</div>
          <div class="pill"><b>Blind:</b> ${blind.name || "-"}</div>
          <div class="pill"><b>Line:</b> ${blind.line || "-"}</div>
          <div class="pill"><b>Type:</b> ${blind.type || "-"}</div>
          <div class="pill"><b>Size:</b> ${blind.size || "-"}</div>
          <div class="pill"><b>Current phase:</b> ${phaseLabel(blind.phase || "broken")}</div>
        </div>

        ${cfg.showWorkflow ? `
        <h3 style="margin:12px 0 6px;">Workflow log</h3>
        <table>
          <thead><tr><th>Date</th><th>From</th><th>To</th><th>Worker</th><th>Role</th></tr></thead>
          <tbody>
            ${
              hist.length === 0
                ? `<tr><td colspan="5">No changes yet.</td></tr>`
                : hist.map(h => `
                  <tr>
                    <td>${h.date ? new Date(h.date).toLocaleString() : "-"}</td>
                    <td>${phaseLabel(h.fromPhase || "-")}</td>
                    <td>${phaseLabel(h.toPhase || "-")}</td>
                    <td>${h.workerName || "-"} (${h.workerId || "-"})</td>
                    <td>${roleLabel(h.role || "-")}</td>
                  </tr>
                `).join("")
            }
          </tbody>
        </table>` : ""}

        ${cfg.showApprovals ? `
        <h3 style="margin:12px 0 6px;">Final approvals</h3>
        <table>
          <thead><tr><th>Approval</th><th>Approved</th><th>By</th><th>Date</th></tr></thead>
          <tbody>${approvalsRows}</tbody>
        </table>` : ""}

        <div class="foot">
          <div>${cfg.footerText || ""}</div>
          <div>Generated from SBTS local data.</div>
        </div>
      </div>
      ${idx < blinds.length - 1 ? '<div style="page-break-after:always;"></div>' : ''}
    `);
  });

  win.document.write(`</body></html>`);
  win.document.close();
  win.focus();
  // Let the user decide when to print (more control)
  // They can click print from browser or Ctrl+P.
}

function printSelectedProjectCertificates() {
  const projectId = document.getElementById("reportProjectSelect").value;
  if (!projectId) return alert("Select a project first.");
  state.currentProjectId = projectId;
  generateProjectCertificates();
}


/* ==========================
   SLIP BLIND (Area -> Project -> Blinds)
========================== */
function getSlipBlinds() {
  return state.blinds.filter((b) => String(b.type || "").toLowerCase().includes("slip"));
}

function renderSlipBlindPage() {
  if (!canUser("viewProjects") && state.currentUser.role !== "admin") {
    alert("No permission to view Slip Blind.");
    openPage("dashboardPage");
    return;
  }

  // reset selection
  state.slip = state.slip || { areaId: "", projectId: "" };

  const crumb = document.getElementById("slipBreadcrumb");
  const areasBox = document.getElementById("slipAreaCards");
  const projectsBox = document.getElementById("slipProjectCards");
  const tableBox = document.getElementById("slipBlindsBox");

  if (!crumb || !areasBox || !projectsBox || !tableBox) return;

  const slipBlindsRaw = getSlipBlinds();

  // Optional quick filter (disabled for UX v2)
  state.slip = state.slip || {};
  state.slip.quickFilter = "";
  const qf = "";
  const slipBlinds = slipBlindsRaw;

  // KPI summary (all projects)
  try {
    const kpi = document.getElementById("slipDashboardStats");
    if (kpi) {
      const total = slipBlinds.length;
      const done = slipBlinds.filter(b => String(b.phase||"").toLowerCase().includes("final")).length;
      const active = total - done;
      const filterLabel = qf ? qf.replace(/_/g, " ").toUpperCase() : "";
      const filterBar = ``; // filters disabled (PATCH 45.1.5)
      kpi.innerHTML = `${filterBar}
        <div class="phase-grid slip-kpi-grid">
          <div class="kpi-pill">
            <div class="phase-left"><span class="phase-dot" style="background:#2D7FF9;"></span><div class="phase-name">Total Slip Blinds (All Projects)</div></div>
            <div class="phase-count">${total}</div>
          </div>
          <div class="kpi-pill">
            <div class="phase-left"><span class="phase-dot" style="background:#18B26A;"></span><div class="phase-name">Completed (Final / Certificate)</div></div>
            <div class="phase-count">${done}</div>
          </div>
          <div class="kpi-pill">
            <div class="phase-left"><span class="phase-dot" style="background:#F4B400;"></span><div class="phase-name">In Progress (No Certificate)</div></div>
            <div class="phase-count">${active}</div>
          </div>
        </div>
      `;
      // KPI filtering disabled (PATCH 45.1.5)
    }
  } catch(e){}

  // breadcrumb text
  const areaName = state.areas.find((a) => a.id === state.slip.areaId)?.name;
  const projectName = state.projects.find((p) => p.id === state.slip.projectId)?.name;

  crumb.innerHTML = `
    <span class="crumb-link" onclick="slipBackToAreas()">Areas</span>
    ${areaName ? `<span class="crumb-sep">/</span><span class="crumb-link" onclick="slipBackToProjects()">${areaName}</span>` : ""}
    ${projectName ? `<span class="crumb-sep">/</span><span class="crumb-current">${projectName}</span>` : ""}
  `;

  // --- Areas view ---
  if (!state.slip.areaId) {
    projectsBox.classList.add("hidden");
    tableBox.classList.add("hidden");
    areasBox.classList.remove("hidden");

    areasBox.innerHTML = "";
    state.areas.forEach((a) => {
      const count = slipBlinds.filter((b) => b.areaId === a.id).length;
      const card = document.createElement("div");
      card.className = "slip-card";
      const done = slipBlinds.filter((b) => b.areaId === a.id && String(b.phase||"").toLowerCase().includes("final")).length;
      const active = count - done;
      card.innerHTML = `
        <div class="slip-card-title">${a.name}</div>
        <div class="slip-card-meta">${count} Slip Blinds</div>
        <div class="tiny" style="margin-top:6px; display:flex; gap:10px; flex-wrap:wrap;">
          <span>Done: <b>${done}</b></span>
          <span>Active: <b>${active}</b></span>
        </div>
      `;
      card.onclick = () => slipSelectArea(a.id);
      areasBox.appendChild(card);
    });
    return;
  }

  // --- Projects view (within selected area) ---
  if (!state.slip.projectId) {
    areasBox.classList.add("hidden");
    tableBox.classList.add("hidden");
    projectsBox.classList.remove("hidden");

    projectsBox.innerHTML = "";
    const projects = state.projects.filter((p) => p.areaId === state.slip.areaId);

    projects.forEach((p) => {
      const count = slipBlinds.filter((b) => b.projectId === p.id).length;
      const card = document.createElement("div");
      card.className = "slip-card";
      card.innerHTML = `
        <div class="slip-card-title">${p.name}</div>
        <div class="slip-card-meta">${count} Slip Blinds</div>
      `;
      card.onclick = () => slipSelectProject(p.id);
      projectsBox.appendChild(card);
    });
    return;
  }

  // --- Table view (within selected project) ---
  areasBox.classList.add("hidden");
  projectsBox.classList.add("hidden");
  tableBox.classList.remove("hidden");

  const list = slipBlinds.filter((b) => b.projectId === state.slip.projectId);

  const tbody = document.getElementById("slipBlindsTableBody");
  tbody.innerHTML = "";
  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8"><div class="empty-state"><div class="empty-title">No slip blinds found</div><div class="empty-sub">Create a slip blind or select another project.</div></div></td></tr>`;
    return;
  }

  // selection state
  state.slip = state.slip || {};
  if (!Array.isArray(state.slip.selectedIds)) state.slip.selectedIds = [];
  const sel = new Set(state.slip.selectedIds);

  const selCountEl = document.getElementById("slipSelectedCount");
  const selAllEl = document.getElementById("slipSelectAll");

  list.forEach((b, idx) => {
    const tr = document.createElement("tr");
    const checked = sel.has(b.id) ? "checked" : "";
    tr.innerHTML = `
      <td><input type="checkbox" ${checked} onchange="slipToggleSelect('${b.id}', this.checked)" /></td>
      <td>${idx + 1}</td>
      <td>${b.name || "-"}</td>
      <td>${b.line || "-"}</td>
      <td>${b.size || "-"}</td>
      <td>${b.rate || "-"}</td>
      <td><span class="phase-badge">${phaseLabel(b.phase || "broken")}</span></td>
      <td>
        <button class="btn-neutral tiny" onclick="openBlindDetails('${b.id}')"><i class="ph ph-eye"></i> Details</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // update counters
  if (selCountEl) selCountEl.textContent = String(sel.size);
  if (selAllEl) selAllEl.checked = (list.length > 0 && sel.size === list.length);
}

function slipSelectArea(areaId) {
  state.slip = state.slip || {};
  state.slip.areaId = areaId;
  state.slip.projectId = "";
  renderSlipBlindPage();
}

function slipSelectProject(projectId) {
  state.slip = state.slip || {};
  state.slip.projectId = projectId;
  renderSlipBlindPage();
}


function slipOpenDashboard(filterKey=""){
  // Sidebar entry: always open Slip Blind dashboard (Areas level)
  state.slip = state.slip || {};
  state.slip.areaId = "";
  state.slip.projectId = "";
  state.slip.selectedIds = [];
  state.slip.quickFilter = "";
  saveState();
  openPage("slipBlindPage");
}

function slipClearQuickFilter(){
  state.slip = state.slip || {};
  state.slip.quickFilter = "";
  saveState();
  renderSlipBlindPage();
}

function slipBackToAreas() {
  state.slip = { areaId: "", projectId: "" };
  renderSlipBlindPage();
}

function slipBackToProjects() {
  if (!state.slip) state.slip = { areaId: "", projectId: "" };
  state.slip.projectId = "";
  renderSlipBlindPage();
}

/* ==========================
   REPORTS
========================== */

const REPORT_CARDS_DEFAULT_ORDER = [
  "total",
  "create",
  "type_isolation",
  "type_slip",
  "type_drop",
  "ph_broken",
  "ph_assembly",
  "ph_tightTorque",
  "ph_finalTight",
  "ph_inspectionReady",
  "approved",
  "pending",
  "slip_mf_demolish",
];

function getReportCardDefs() {
  return [
    { key: "total", label: "Total Blinds", kind: "all" },
    { key: "create", label: "Create Blind", kind: "create" },
    { key: "type_isolation", label: "Isolation blind", kind: "type", value: "Isolation Blind" },
    { key: "type_slip", label: "Slip blind", kind: "type", value: "Slip Blind" },
    { key: "type_drop", label: "Drop spool", kind: "type", value: "Drop Spool" },

    { key: "ph_broken", label: "Broken / Preparation", kind: "phase", value: "broken" },
    { key: "ph_assembly", label: "Assembly", kind: "phase", value: "assembly" },
    { key: "ph_tightTorque", label: "Tight & Torque", kind: "phase", value: "tightTorque" },
    { key: "ph_finalTight", label: "Final Tight", kind: "phase", value: "finalTight" },
    { key: "ph_inspectionReady", label: "Inspection & Ready", kind: "phase", value: "inspectionReady" },

    { key: "approved", label: "Completed approval", kind: "approval", value: "approved" },
    { key: "pending", label: "Pending Approval", kind: "approval", value: "pending" },

    { key: "slip_mf_demolish", label: "Slip blind MF approved (Demolish)", kind: "special", value: "slip_mf_demolish" },
  ];
}

function ensureReportsCardsConfig() {
  state.reports = state.reports || {};
  if (!Array.isArray(state.reports.cardOrder) || state.reports.cardOrder.length === 0) {
    state.reports.cardOrder = REPORT_CARDS_DEFAULT_ORDER.slice();
  }
  state.reports.cardVisibility = state.reports.cardVisibility || {};
  const defs = getReportCardDefs();
  defs.forEach((d) => {
    if (typeof state.reports.cardVisibility[d.key] !== "boolean") state.reports.cardVisibility[d.key] = true;
  });
  if (typeof state.reports.activeCardKey === "undefined") state.reports.activeCardKey = null;
  if (!state.reports.preset) state.reports.preset = "management";
}

function computeReportsCounts(blinds) {
  const counts = {
    total: blinds.length,
    create: 0,
    byType: { "Isolation Blind": 0, "Slip Blind": 0, "Drop Spool": 0 },
    byPhase: {},
    approved: 0,
    pending: 0,
    slip_mf_demolish: 0,
  };

  PHASES.forEach((p) => (counts.byPhase[p] = 0));

  blinds.forEach((b) => {
    const ph = b.phase || "broken";
    counts.byPhase[ph] = (counts.byPhase[ph] || 0) + 1;

    const t = b.type || "";
    if (counts.byType[t] !== undefined) counts.byType[t]++;

    // "Create Blind" = created but never moved phase (no history entries)
    if ((b.history || []).length === 0) counts.create++;

    if (isCertificateApproved(b)) counts.approved++;

    // Slip MF demolish approval
    const mf = (b.finalApprovals || {}).metal_foreman_demolish;
    if ((b.type || "") === "Slip Blind" && mf && mf.status === "approved") counts.slip_mf_demolish++;
  });

  counts.pending = Math.max(0, counts.total - counts.approved);
  return counts;
}

function applyReportsCardFilter(blinds, cardKey) {
  if (!cardKey) return blinds;

  const def = getReportCardDefs().find((d) => d.key === cardKey);
  if (!def) return blinds;

  if (def.kind === "all") return blinds;

  if (def.kind === "create") {
    return blinds.filter((b) => (b.history || []).length === 0);
  }

  if (def.kind === "type") {
    return blinds.filter((b) => (b.type || "") === def.value);
  }

  if (def.kind === "phase") {
    return blinds.filter((b) => (b.phase || "broken") === def.value);
  }

  if (def.kind === "approval") {
    return def.value === "approved"
      ? blinds.filter((b) => isCertificateApproved(b))
      : blinds.filter((b) => !isCertificateApproved(b));
  }

  if (def.kind === "special" && def.value === "slip_mf_demolish") {
    return blinds.filter((b) => (b.type || "") === "Slip Blind" && (b.finalApprovals || {}).metal_foreman_demolish?.status === "approved");
  }

  return blinds;
}

function toggleReportsCardFilter(cardKey) {
  ensureReportsCardsConfig();
  state.reports.activeCardKey = (state.reports.activeCardKey === cardKey) ? null : cardKey;
  saveState();
  renderReports();
}

function hydrateReportsCardsSettingsUI() {
  ensureReportsCardsConfig();
  const section = document.getElementById("reportsCardsSettingsSection");
  const list = document.getElementById("reportsCardsControl");

  // Admin only
  const isAdmin = state.currentUser?.role === "admin";
  if (section) section.classList.toggle("hidden", !isAdmin);
  if (!isAdmin || !list) return;

  const defs = getReportCardDefs();
  const defsByKey = Object.fromEntries(defs.map((d) => [d.key, d]));

  // Ensure order contains any new keys
  defs.forEach((d) => {
    if (!state.reports.cardOrder.includes(d.key)) state.reports.cardOrder.push(d.key);
  });

  list.innerHTML = "";
  state.reports.cardOrder.forEach((key) => {
    const d = defsByKey[key];
    if (!d) return;

    const item = document.createElement("div");
    item.className = "card-control-item";
    item.draggable = true;
    item.dataset.key = key;

    const checked = !!state.reports.cardVisibility[key];

    item.innerHTML = `
      <div class="card-control-left">
        <div class="drag-handle" title="Drag to reorder">⋮⋮</div>
        <div class="card-control-texts">
          <div class="card-control-label">${d.label}</div>
          <div class="card-control-meta">${d.kind}</div>
        </div>
      </div>
      <div class="card-control-right">
        <label class="toggle-switch">
          <input type="checkbox" ${checked ? "checked" : ""} data-toggle="${key}" />
          Show
        </label>
      </div>
    `;

    // Drag events
    item.addEventListener("dragstart", (e) => {
      e.dataTransfer.setData("text/plain", key);
      item.classList.add("dragging");
    });
    item.addEventListener("dragend", () => item.classList.remove("dragging"));
    item.addEventListener("dragover", (e) => e.preventDefault());
    item.addEventListener("drop", (e) => {
      e.preventDefault();
      const fromKey = e.dataTransfer.getData("text/plain");
      const toKey = key;
      if (!fromKey || fromKey === toKey) return;

      const order = state.reports.cardOrder.slice();
      const fromIdx = order.indexOf(fromKey);
      const toIdx = order.indexOf(toKey);
      if (fromIdx < 0 || toIdx < 0) return;

      order.splice(fromIdx, 1);
      order.splice(toIdx, 0, fromKey);

      state.reports.cardOrder = order;
      state.reports.preset = "custom";
      saveState();
      hydrateReportsCardsSettingsUI();
    });

    list.appendChild(item);
  });

  // Toggle handlers
  list.querySelectorAll("input[type='checkbox'][data-toggle]").forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const k = e.target.getAttribute("data-toggle");
      state.reports.cardVisibility[k] = e.target.checked;
      state.reports.preset = "custom";
      saveState();
      // update reports view if open
      if (!document.getElementById("reportsPage")?.classList.contains("hidden")) renderReports();
    });
  });
}

function applyReportsPreset(name) {
  ensureReportsCardsConfig();

  const showAll = () => {
    getReportCardDefs().forEach((d) => state.reports.cardVisibility[d.key] = true);
  };

  if (name === "management") {
    // Focus on totals, types, approvals, and inspection readiness
    showAll();
    const hideKeys = ["ph_assembly", "ph_tightTorque", "ph_finalTight", "slip_mf_demolish"];
    hideKeys.forEach((k) => state.reports.cardVisibility[k] = false);
    state.reports.preset = "management";
  } else if (name === "safety") {
    // Focus on slip, pending, demolish approval, and phase readiness
    showAll();
    const hideKeys = ["type_drop"];
    hideKeys.forEach((k) => state.reports.cardVisibility[k] = false);
    state.reports.preset = "safety";
  }

  saveState();
  hydrateReportsCardsSettingsUI();
  if (!document.getElementById("reportsPage")?.classList.contains("hidden")) renderReports();
}

function resetReportsCardsDefaults() {
  ensureReportsCardsConfig();
  state.reports.cardOrder = REPORT_CARDS_DEFAULT_ORDER.slice();
  getReportCardDefs().forEach((d) => state.reports.cardVisibility[d.key] = true);
  state.reports.activeCardKey = null;
  state.reports.preset = "management";
  saveState();
  hydrateReportsCardsSettingsUI();
  if (!document.getElementById("reportsPage")?.classList.contains("hidden")) renderReports();
}

function initReportsUI() {
  if (!canUser("viewReports")) {
    alert("No permission to view reports.");
    openPage("dashboardPage");
    return;
  }

  const areaSel = document.getElementById("reportAreaSelect");
  const projectSel = document.getElementById("reportProjectSelect");

  areaSel.innerHTML = `<option value="">All areas</option>`;
  state.areas.forEach((a) => {
    const opt = document.createElement("option");
    opt.value = a.id;
    opt.textContent = a.name;
    areaSel.appendChild(opt);
  });

  projectSel.innerHTML = `<option value="">All projects</option>`;
  state.projects.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.name;
    projectSel.appendChild(opt);
  });

  renderReports();
}

function onReportAreaChange() {
  const areaId = document.getElementById("reportAreaSelect").value;
  const projectSel = document.getElementById("reportProjectSelect");

  projectSel.innerHTML = `<option value="">All projects</option>`;
  const list = areaId ? state.projects.filter((p) => p.areaId === areaId) : state.projects;

  list.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.name;
    projectSel.appendChild(opt);
  });

  renderReports();
}

function getFilteredBlindsForReports() {
  const areaId = document.getElementById("reportAreaSelect").value;
  const projectId = document.getElementById("reportProjectSelect").value;

  let blinds = state.blinds.slice();
  if (areaId) blinds = blinds.filter((b) => b.areaId === areaId);
  if (projectId) blinds = blinds.filter((b) => b.projectId === projectId);

  return blinds;
}


function renderReports() {
  ensureReportsCardsConfig();

  const baseBlinds = getFilteredBlindsForReports();
  const counts = computeReportsCounts(baseBlinds);

  const activeKey = state.reports.activeCardKey;
  const displayBlinds = applyReportsCardFilter(baseBlinds, activeKey);

  const displayCounts = computeReportsCounts(displayBlinds);

  renderReportsKpis({ counts, activeKey });

  // Progress bars should reflect the current filter
  renderPhaseProgressBars(displayCounts.byPhase, displayCounts.total);

  // Summary
  const summary = document.getElementById("reportsSummary");
  if (summary) {
    summary.innerHTML = "";
    [
      ["Total blinds", displayCounts.total],
      ["Completed approval", displayCounts.approved],
      ["Pending approval", displayCounts.pending],
      ["Active filter", activeKey ? (getReportCardDefs().find(d => d.key === activeKey)?.label || "-") : "None"],
    ].forEach(([label, value]) => {
      const div = document.createElement("div");
      div.className = "summary-item";
      div.innerHTML = `<div class="summary-label">${label}</div><div class="summary-value">${value}</div>`;
      summary.appendChild(div);
    });
  }


  // Management & Operational panels (v1)
  try {
    const mg = document.getElementById("reportsMgmt");
    const ops = document.getElementById("reportsOps");

    const srcBlinds = baseBlinds || [];
    const pendingBlinds = srcBlinds.filter(b => !isCertificateApproved(b));

    // --- Management: top projects by pending approvals ---
    if (mg) {
      const byProject = new Map();
      for (const b of srcBlinds) {
        const pid = b.projectId || "__none__";
        if (!byProject.has(pid)) byProject.set(pid, { total:0, pending:0, approved:0 });
        const row = byProject.get(pid);
        row.total += 1;
        if (isCertificateApproved(b)) row.approved += 1;
        else row.pending += 1;
      }

      const projName = (pid) => {
        if (pid === "__none__") return "(No project)";
        return state.projects.find(p => p.id === pid)?.name || pid;
      };

      const rows = Array.from(byProject.entries())
        .map(([pid, v]) => ({ pid, name: projName(pid), total: v.total, pending: v.pending, approved: v.approved }))
        .sort((a, b) => (b.pending - a.pending) || a.name.localeCompare(b.name))
        .slice(0, 8);

      mg.innerHTML = `
        <div class="tiny" style="margin-bottom:8px;">Top projects with pending approvals (based on current Reports filters).</div>
        <div class="table-wrap">
          <table class="main-table" style="margin:0;">
            <thead><tr><th>Project</th><th>Total</th><th>Approved</th><th>Pending</th></tr></thead>
            <tbody>
              ${rows.map(r => `
                <tr>
                  <td>${escapeHtml(r.name)}</td>
                  <td>${r.total}</td>
                  <td>${r.approved}</td>
                  <td><b>${r.pending}</b></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;
    }

    // --- Operational: pending by phase + quick list ---
    if (ops) {
      const byPhase = {};
      for (const p of PHASES) byPhase[p] = 0;
      for (const b of pendingBlinds) {
        const ph = b.phase || "broken";
        byPhase[ph] = (byPhase[ph] || 0) + 1;
      }

      const phaseRows = Object.entries(byPhase)
        .filter(([_, c]) => c > 0)
        .sort((a, b) => b[1] - a[1]);

      const list = pendingBlinds.slice(0, 10).map(b => {
        const project = state.projects.find(p => p.id === b.projectId)?.name || "-";
        const name = b.name || b.tagNumber || b.id;
        return { project, phase: b.phase || "-", name };
      });

      ops.innerHTML = `
        <div class="tiny" style="margin-bottom:8px;">Pending approvals by phase + quick list (first 10).</div>
        <div class="table-wrap" style="margin-bottom:10px;">
          <table class="main-table" style="margin:0;">
            <thead><tr><th>Phase</th><th>Pending</th></tr></thead>
            <tbody>
              ${phaseRows.map(([ph, c]) => `<tr><td>${escapeHtml(ph)}</td><td><b>${c}</b></td></tr>`).join('') || `<tr><td colspan="2">No pending items.</td></tr>`}
            </tbody>
          </table>
        </div>

        <div class="table-wrap">
          <table class="main-table" style="margin:0;">
            <thead><tr><th>Project</th><th>Blind</th><th>Phase</th></tr></thead>
            <tbody>
              ${list.map(r => `<tr><td>${escapeHtml(r.project)}</td><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.phase)}</td></tr>`).join('') || `<tr><td colspan="3">No pending items.</td></tr>`}
            </tbody>
          </table>
        </div>
      `;
    }
  } catch (e) {
    console.warn("Reports panels render failed", e);
  }

  // Optional canvas chart remains disabled by default
}

function renderReportsKpis({ counts, activeKey }) {
  const box = document.getElementById("reportsKpis");
  if (!box) return;

  const defs = getReportCardDefs();
  const defsByKey = Object.fromEntries(defs.map((d) => [d.key, d]));

  // Build cards based on visibility + order
  const order = (state.reports.cardOrder || []).filter((k) => defsByKey[k]);
  const visible = order.filter((k) => state.reports.cardVisibility?.[k] !== false);

  box.innerHTML = "";

  const valueForKey = (key) => {
    const d = defsByKey[key];
    if (!d) return 0;
    if (key === "total") return counts.total;
    if (key === "create") return counts.create;
    if (d.kind === "type") return counts.byType[d.value] || 0;
    if (d.kind === "phase") return counts.byPhase[d.value] || 0;
    if (key === "approved") return counts.approved;
    if (key === "pending") return counts.pending;
    if (key === "slip_mf_demolish") return counts.slip_mf_demolish;
    return 0;
  };

  const subForKey = (key) => {
    const d = defsByKey[key];
    if (!d) return "";
    if (d.kind === "phase") return "Click to filter by phase";
    if (d.kind === "type") return "Click to filter by type";
    if (d.kind === "approval") return "Click to filter by approval";
    if (d.kind === "create") return "Created (no history yet)";
    if (d.kind === "special") return "Slip blind with MF demolish approval";
    return "Click to filter";
  };

  // If no cards visible, show a gentle hint for admin
  if (visible.length === 0) {
    const div = document.createElement("div");
    div.className = "tiny";
    div.textContent = "No report cards enabled. (Admin: Reports Cards page)";
    box.appendChild(div);
    return;
  }

  visible.forEach((key) => {
    const d = defsByKey[key];
    const card = document.createElement("div");
    card.className = "report-kpi-card clickable" + (activeKey === key ? " active" : "");
    card.innerHTML = `
      <div class="report-kpi-title">${d.label}</div>
      <div class="report-kpi-value">${valueForKey(key)}</div>
      <div class="report-kpi-sub">${subForKey(key)}</div>
    `;
    card.addEventListener("click", () => toggleReportsCardFilter(key));
    box.appendChild(card);
  });
}


function renderPhaseProgressBars(byPhase, total) {
  const el = document.getElementById("phaseProgressBars");
  if (!el) return;

  // Colors from CSS variables
  const css = getComputedStyle(document.documentElement);
  const phaseColor = (ph) => {
    const key = `--ph-${ph}`;
    const v = (css.getPropertyValue(key) || "").trim();
    return v || "#174c7e";
  };

  el.innerHTML = "";
  wfProjectPhaseIds(state.currentProjectId, { includeInactive: false }).forEach((ph) => {
    const count = byPhase[ph] || 0;
    const pct = total ? Math.round((count / total) * 100) : 0;

    const row = document.createElement("div");
    row.className = "phase-progress-row";
    row.innerHTML = `
      <div class="phase-progress-label">${phaseLabel(ph)}</div>
      <div class="phase-progress-bar">
        <div class="phase-progress-fill" style="width:${pct}%; background:${phaseColor(ph)};"></div>
      </div>
      <div class="phase-progress-count">${count}</div>
    `;
    el.appendChild(row);
  });
}

function buildReportsDocumentHtml({ baseBlinds, displayBlinds }) {
  const b = state.branding || {};
  const today = formatDateDDMMYYYY(new Date());

  const areaId = document.getElementById("reportAreaSelect")?.value || "";
  const projectId = document.getElementById("reportProjectSelect")?.value || "";

  const areaName = areaId ? (state.areas.find(a => a.id === areaId)?.name || "-") : "All areas";
  const projectName = projectId ? (state.projects.find(p => p.id === projectId)?.name || "-") : "All projects";

  const baseCounts = computeReportsCounts(baseBlinds || []);

  // Cards follow admin selection (order + visibility) exactly like the Reports page
  const defs = getReportCardDefs();
  const defsByKey = Object.fromEntries(defs.map((d) => [d.key, d]));
  const order = (state.reports?.cardOrder || []).filter((k) => defsByKey[k]);
  const visible = order.filter((k) => state.reports?.cardVisibility?.[k] !== false);

  const valueForKey = (key) => {
    const d = defsByKey[key];
    if (!d) return 0;
    if (key === "total") return baseCounts.total;
    if (key === "create") return baseCounts.create;
    if (d.kind === "type") return baseCounts.byType[d.value] || 0;
    if (d.kind === "phase") return baseCounts.byPhase[d.value] || 0;
    if (key === "approved") return baseCounts.approved;
    if (key === "pending") return baseCounts.pending;
    if (key === "slip_mf_demolish") return baseCounts.slip_mf_demolish;
    return 0;
  };

  const cardsHtml = (visible.length ? visible : REPORT_CARDS_DEFAULT_ORDER)
    .filter((k) => defsByKey[k])
    .map((key) => {
      const d = defsByKey[key];
      return `<div class="card"><div class="t">${escapeHtml(d.label)}</div><div class="v">${valueForKey(key)}</div></div>`;
    })
    .join("");

  const totalDisplay = (displayBlinds || []).length;

  const byPhase = {};
  PHASES.forEach((p) => (byPhase[p] = 0));
  (displayBlinds || []).forEach((x) => { byPhase[x.phase || "broken"] = (byPhase[x.phase || "broken"] || 0) + 1; });

  const phaseRows = PHASES.map((ph) => {
    const count = byPhase[ph] || 0;
    const pct = totalDisplay ? Math.round((count / totalDisplay) * 100) : 0;
    return `
      <tr>
        <td style="font-weight:700;">${phaseLabel(ph)}</td>
        <td>
          <div style="height:12px;border-radius:999px;background:rgba(15,23,42,0.08);overflow:hidden;">
            <div style="height:100%;width:${pct}%;border-radius:999px;background:${getComputedStyle(document.documentElement).getPropertyValue(`--ph-${ph}`) || "#174c7e"};"></div>
          </div>
        </td>
        <td style="text-align:right;font-weight:800;">${count}</td>
      </tr>
    `;
  }).join("");

  const tableRows = (displayBlinds || []).map((x, i) => {
    const a = state.areas.find((t) => t.id === x.areaId)?.name || "-";
    const p = state.projects.find((t) => t.id === x.projectId)?.name || "-";
    const approved = isCertificateApproved(x) ? "YES" : "NO";
    return `
      <tr>
        <td>${i + 1}</td>
        <td>${escapeHtml(x.name || "")}</td>
        <td>${escapeHtml(a)}</td>
        <td>${escapeHtml(p)}</td>
        <td>${escapeHtml(x.line || "")}</td>
        <td>${escapeHtml(x.type || "")}</td>
        <td>${escapeHtml(x.size || "")}</td>
        <td>${escapeHtml(x.rate || "")}</td>
        <td>${escapeHtml(phaseLabel(x.phase || "broken"))}</td>
        <td style="text-align:center;font-weight:800;">${approved}</td>
      </tr>
    `;
  }).join("");

  const logo = b.companyLogo ? `<div style="width:54px;height:54px;border-radius:14px;background:#eef2ff;background-image:url('${b.companyLogo}');background-size:cover;background-position:center;"></div>` : "";

  return `
<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>SBTS Report</title>
<style>
  *{box-sizing:border-box;font-family:system-ui,-apple-system,Segoe UI,Arial,sans-serif;}
  body{margin:0;background:#fff;color:#111827;}
  .wrap{padding:22px;}
  .head{display:flex;align-items:center;justify-content:space-between;gap:12px;border:1px solid #e5e7eb;border-radius:18px;padding:14px 16px;}
  .h-left{display:flex;align-items:center;gap:12px;}
  .h-title{font-size:18px;font-weight:900;}
  .h-sub{margin-top:4px;font-size:12px;color:#6b7280;}
  .badge{padding:6px 12px;border-radius:999px;background:#eef2ff;font-weight:900;}
  .grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:10px;margin-top:12px;}
  .card{border:1px solid #e5e7eb;border-radius:16px;padding:10px 12px;}
  .t{font-size:12px;color:#6b7280;font-weight:700;}
  .v{font-size:22px;font-weight:900;margin-top:6px;}
  h2{font-size:14px;margin:18px 0 8px;}
  table{width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:14px;overflow:hidden;}
  th,td{padding:8px 10px;border-bottom:1px solid #e5e7eb;font-size:11px;text-align:left;}
  th{background:#f3f4ff;color:#6b7280;font-weight:800;}
  tr:nth-child(even) td{background:#fafbff;}
  @page{size:A4; margin:12mm;}
</style>
</head>
<body>
  <div class="wrap">
    <div class="head">
      <div class="h-left">
        ${logo}
        <div>
          <div class="h-title">${escapeHtml(b.siteTitle || "SBTS Reports")}</div>
          <div class="h-sub">${escapeHtml(b.programTitle || "Smart Blind Tag System")} • ${escapeHtml(areaName)} • ${escapeHtml(projectName)}</div>
        </div>
      </div>
      <div class="badge">${today}</div>
    </div>

    <div class="grid">
      ${cardsHtml}
    </div>

    <h2>All phases (progress)</h2>
    <table>
      <thead><tr><th style="width:220px;">Phase</th><th>Progress</th><th style="width:80px;text-align:right;">Count</th></tr></thead>
      <tbody>${phaseRows}</tbody>
    </table>

    <h2>Blinds list (filtered)</h2>
    <table>
      <thead>
        <tr>
          <th style="width:40px;">#</th>
          <th>Blind</th><th>Area</th><th>Project</th><th>Line</th><th>Type</th>
          <th style="width:70px;">Size</th><th style="width:70px;">Rate</th><th>Phase</th><th style="width:90px;text-align:center;">Approved</th>
        </tr>
      </thead>
      <tbody>${tableRows || `<tr><td colspan="10" style="text-align:center;color:#6b7280;">No records</td></tr>`}</tbody>
    </table>
  </div>
</body>
</html>
  `.trim();
}

function openReportPrintWindow(html) {
  const w = window.open("", "_blank");
  if (!w) return alert("Popup blocked. Please allow popups for this site.");
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(() => w.print(), 450);
}

function printReportsA4() {
  ensureReportsCardsConfig();
  const baseBlinds = getFilteredBlindsForReports();
  const displayBlinds = applyReportsCardFilter(baseBlinds, state.reports.activeCardKey);
  const html = buildReportsDocumentHtml({ baseBlinds, displayBlinds });
  openReportPrintWindow(html);
}

function exportReportsPDF() {
  // Same as print (user chooses Save as PDF)
  ensureReportsCardsConfig();
  const baseBlinds = getFilteredBlindsForReports();
  const displayBlinds = applyReportsCardFilter(baseBlinds, state.reports.activeCardKey);
  const html = buildReportsDocumentHtml({ baseBlinds, displayBlinds });
  openReportPrintWindow(html);
}

function shareReports() {
  ensureReportsCardsConfig();
  const blinds = applyReportsCardFilter(getFilteredBlindsForReports(), state.reports.activeCardKey);
  const total = blinds.length;
  const approved = blinds.filter((b) => isCertificateApproved(b)).length;

  const areaId = document.getElementById("reportAreaSelect")?.value || "";
  const projectId = document.getElementById("reportProjectSelect")?.value || "";
  const areaName = areaId ? (state.areas.find(a => a.id === areaId)?.name || "-") : "All areas";
  const projectName = projectId ? (state.projects.find(p => p.id === projectId)?.name || "-") : "All projects";

  const text = `SBTS Report
Area: ${areaName}
Project: ${projectName}
Total Blinds: ${total}
Completed approval: ${approved}
Pending approval: ${Math.max(0, total-approved)}
Date: ${formatDateDDMMYYYY(new Date())}`;

  if (navigator.share) {
    navigator.share({ title: "SBTS Report", text }).catch(() => {});
    return;
  }

  // fallback: copy to clipboard
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => alert("Report summary copied. Paste it in WhatsApp/Email."));
  } else {
    alert(text);
  }
}

// Small helper for safe HTML output in report export
function escapeHtml(str) {
  return String(str || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// Short alias used in some render code
function esc(str) { return escapeHtml(str); }


// ----------------------------------------------------
// Notification category meta (hoisted-safe)
// NOTE: use var to avoid temporal-dead-zone issues when referenced before declaration.
var NOTIF_CATEGORY_META = (typeof window !== 'undefined' && window.NOTIF_CATEGORY_META) ? window.NOTIF_CATEGORY_META : {
  action: { label: "Action Required", color: "red", icon: "sbts-ico-action", badge: "Action" },
  system: { label: "System", color: "blue", icon: "sbts-ico-system", badge: "System" },
  warning:{ label: "Warnings", color: "amber", icon: "sbts-ico-warning", badge: "Warning" },
  admin:  { label: "Admin Notes", color: "gray", icon: "sbts-ico-admin", badge: "Admin" },
};
if (typeof window !== 'undefined') window.NOTIF_CATEGORY_META = NOTIF_CATEGORY_META;

// ------------------------------------------------------------
// Compatibility shims (older patches expect these globals)
// ------------------------------------------------------------
try {
  if (typeof window !== 'undefined' && !window.esc) window.esc = escapeHtml;
} catch (_) {}
// esc(...) was used in a few render paths.
// showToast(...) is referenced by the global error handler.
try {
  if (typeof window !== "undefined") {
    if (!window.esc) window.esc = escapeHtml;
    // If toast(...) exists later, we'll re-point showToast then as well.
    if (!window.showToast && typeof window.toast === "function") window.showToast = window.toast;
  }
} catch (e) {
  // ignore
}

function escapeJs(str){
  return String(str ?? "")
    .replaceAll("\\", "\\\\")
    .replaceAll("'", "\\'")
    .replaceAll("\n"," ")
    .replaceAll("\r"," ");
}


function toast(message) {
  // Lightweight toast without external libs
  const msg = String(message || "");
  let el = document.getElementById("sbtsToast");
  if (!el) {
    el = document.createElement("div");
    el.id = "sbtsToast";
    el.style.position = "fixed";
    el.style.left = "50%";
    el.style.bottom = "18px";
    el.style.transform = "translateX(-50%)";
    el.style.padding = "10px 14px";
    el.style.borderRadius = "12px";
    el.style.background = "rgba(0,0,0,0.75)";
    el.style.color = "#fff";
    el.style.fontSize = "14px";
    el.style.zIndex = "9999";
    el.style.maxWidth = "90vw";
    el.style.textAlign = "center";
    el.style.display = "none";
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.display = "block";
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.display = "none"; }, 1700);
}

function showToast(message) { toast(message); }

// Ensure legacy callers (and the global error handler) can use showToast(...)
try {
  if (typeof window !== 'undefined' && !window.showToast) window.showToast = toast;
} catch (_) {}


function renderReportsChart(blinds) {
  const canvas = document.getElementById("reportsChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  // Make canvas responsive to container width
  const parent = canvas.parentElement;
  const targetW = Math.max(520, Math.min(1100, (parent?.clientWidth || canvas.width) - 10));
  canvas.width = targetW;
  canvas.height = 320;

  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  // Count by phase
  const byPhase = {};
  PHASES.forEach((p) => (byPhase[p] = 0));
  blinds.forEach((b) => { byPhase[b.phase || "broken"] = (byPhase[b.phase || "broken"] || 0) + 1; });

  const values = PHASES.map((p) => byPhase[p] || 0);
  const max = Math.max(1, ...values);

  // Colors from CSS variables (fallback to a safe default)
  const css = getComputedStyle(document.documentElement);
  const phaseColor = (ph) => {
    const key = `--ph-${ph}`;
    const v = (css.getPropertyValue(key) || "").trim();
    return v || "#174c7e";
  };

  // Layout
  const padL = 170;
  const padR = 30;
  const padT = 46;
  const padB = 28;
  const rowH = (h - padT - padB) / PHASES.length;
  const barH = Math.max(10, Math.min(18, rowH * 0.55));
  const barMaxW = w - padL - padR;

  // Title
  ctx.fillStyle = "#111827";
  ctx.font = "700 13px system-ui";
  ctx.fillText("Slip/Blind Phase distribution (count)", 12, 22);

  // Grid / baseline
  ctx.strokeStyle = "rgba(15,23,42,0.08)";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const x = padL + (barMaxW * i) / 4;
    ctx.beginPath();
    ctx.moveTo(x, padT - 6);
    ctx.lineTo(x, h - padB + 6);
    ctx.stroke();
    ctx.fillStyle = "#6b7280";
    ctx.font = "11px system-ui";
    ctx.fillText(String(Math.round((max * i) / 4)), x - 8, h - 8);
  }

  // Bars
  PHASES.forEach((ph, i) => {
    const val = byPhase[ph] || 0;
    const yMid = padT + i * rowH + rowH / 2;
    const y = yMid - barH / 2;
    const bw = (val / max) * barMaxW;

    // Label
    ctx.fillStyle = "#111827";
    ctx.font = "12px system-ui";
    ctx.fillText(phaseLabel(ph), 12, yMid + 4);

    // Background bar
    ctx.fillStyle = "rgba(15,23,42,0.06)";
    ctx.fillRect(padL, y, barMaxW, barH);

    // Value bar
    ctx.fillStyle = phaseColor(ph);
    ctx.fillRect(padL, y, bw, barH);

    // Count text
    ctx.fillStyle = "#111827";
    ctx.font = "700 12px system-ui";
    ctx.fillText(String(val), padL + bw + 8, yMid + 4);
  });
}

function printReports() { window.print(); }

function exportReports() {
  const blinds = getFilteredBlindsForReports();
  const rows = [
    ["Blind", "Area", "Project", "Line", "Type", "Size", "Rate", "Phase", "CertApproved"].join(","),
    ...blinds.map((b) => {
      const a = state.areas.find((x) => x.id === b.areaId)?.name || "";
      const p = state.projects.find((x) => x.id === b.projectId)?.name || "";
      const ca = isCertificateApproved(b) ? "YES" : "NO";
      return [b.name, a, p, b.line || "", b.type || "", b.size || "", b.rate || "", phaseLabel(b.phase || "broken"), ca]
        .map((s) => `"${String(s).replaceAll('"', '""')}"`)
        .join(",");
    }),
  ].join("\n");

  const blob = new Blob([rows], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "sbts_reports.csv";
  a.click();
  URL.revokeObjectURL(url);
}

/* ==========================
   USERS / PERMISSIONS
========================== */
function renderUsers() {
  const tbody = document.getElementById("usersTableBody");
  tbody.innerHTML = "";

  state.users.forEach((u, index) => {
    const tr = document.createElement("tr");
    tr.id = `user_${u.id}`;
    // 47.10 focus/highlight when coming from notifications
    if (state.ui && state.ui.focusUserId && state.ui.focusUserId === u.id) {
      tr.classList.add('row-focus');
    }
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${u.fullName || "-"}</td>
      <td>${u.username}</td>
      <td>${roleLabel(u.role)}</td>
      <td>${u.status}</td>
      <td>${(u.projectAssignments && u.projectAssignments.length) ? `${u.projectAssignments.length} project(s)` : "—"}</td>
      <td>
        <div class="table-actions">
          ${
            u.status === "pending" && canUser("adminApproveUsers")
              ? `<button class="secondary-btn tiny" onclick="approveUser('${u.id}')">Approve</button>
              <button class="secondary-btn tiny" onclick="rejectUser('${u.id}')">Reject</button>`
              : ""
          }
          ${
            canUser("adminApproveUsers")
              ? `<button class="secondary-btn tiny" onclick="openEditPermissionsModal('${u.id}')">Permissions</button>`
              : ""
          }

${
  canUser("adminApproveUsers") && u.status === "active"
    ? `<button class="secondary-btn tiny" onclick="openAssignProjectsModal('${u.id}')">Assign Projects</button>`
    : ""
}
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Focus a user row when opened from a notification
  try {
    const focusId = state.ui?.usersFocusId;
    if (focusId) {
      const el = document.getElementById(`user_${focusId}`);
      if (el) {
        el.style.outline = "2px solid rgba(10,103,163,0.35)";
        el.style.outlineOffset = "2px";
        el.scrollIntoView({ block: "center", behavior: "smooth" });
      }
      state.ui.usersFocusId = null;
      saveState();
    }
  } catch(_) {}

}

function approveUser(userId) {
  if (!canUser("adminApproveUsers")) return alert("No permission.");
  const user = state.users.find((u) => u.id === userId);
  if (!user) return;

  user.status = "active";
  saveState();

// Sync related NEW_USER request in Inbox (if exists)
try {
  ensureRequestsArray();
  const r = (state.requests || []).find(x => x && x.type === "NEW_USER" && x.meta && x.meta.userId === userId && (x.status || "pending") !== "approved");
  if (r) setRequestStatus(r.id, "approved");
} catch(e) {}

// Remind approver to assign projects (account approved, but access comes from assignments)
try {
  const approverId = getCurrentUserIdStable();
  if (approverId) {
    SBTS_ACTIVITY.pushNotification({
      category: "action",
      scope: "user",
      recipients: [approverId],
      title: "Assign projects",
      message: `Approved user ${user?.fullName || user?.username || ""}. Assign projects to grant access.`,
      requiresAction: true,
      actionKey: `assign_projects:${userId}`,
      resolved: false,
      actorId: approverId
    });
  }
} catch(e) {}

  // Resolve pending approval notifications for approvers
  resolveNotificationsByActionKey(`user_approval:${userId}`);

  // Notify the user
  SBTS_ACTIVITY.pushNotification({
    category: "system",
    scope: "user",
    recipients: [userId],
    title: "Account approved",
    message: "Your account has been approved. Admin will assign you to projects before you can work.",
    requiresAction: false,
    resolved: true,
    actorId: state.currentUser?.id || null
  });

  renderUsers();
  updateNotificationsBadge();
}


/* ==========================
   PATCH 46.5.1 - User Project Assignments (Real Always-on)
   - Users register once (no project at registration)
   - Admin assigns projects after approval
========================== */
let __assignProjectsUserId = null;

function openAssignProjectsModal(userId){
  if (!canUser("adminApproveUsers")) return alert("No permission.");
  const u = state.users.find(x => x.id === userId);
  if(!u) return;
  __assignProjectsUserId = userId;

  const line = document.getElementById("assignProjectsUserLine");
  if(line){
    const nm = (u.fullName || u.username || "User");
    line.textContent = `Assign projects for: ${nm} (@${u.username || "-"})`;
  }

  // Ensure array
  u.projectAssignments = Array.isArray(u.projectAssignments) ? u.projectAssignments : [];
  renderAssignProjectsList();
  openModal("assignProjectsModal");
}

function renderAssignProjectsList(){
  const box = document.getElementById("assignProjectsList");
  if(!box) return;
  const q = (document.getElementById("assignProjectsSearch")?.value || "").trim().toLowerCase();
  const u = state.users.find(x => x.id === __assignProjectsUserId);
  const assigned = new Set(Array.isArray(u?.projectAssignments) ? u.projectAssignments : []);
  box.innerHTML = "";

  const list = Array.isArray(state.projects) ? state.projects : [];
  const filtered = list.filter(p => {
    if(!p) return false;
    const name = (p.name || p.id || "").toLowerCase();
    return !q || name.includes(q);
  });

  if(!filtered.length){
    box.innerHTML = `<div class="empty-state" style="padding:18px;border-radius:12px;">
      <div class="empty-title">No projects found</div>
      <div class="empty-subtitle">Create a project first, then assign it to users.</div>
    </div>`;
    return;
  }

  filtered.forEach(p => {
    const row = document.createElement("label");
    row.style.display = "flex";
    row.style.alignItems = "center";
    row.style.gap = "10px";
    row.style.padding = "10px 8px";
    row.style.borderRadius = "10px";
    row.style.cursor = "pointer";
    row.style.border = "1px solid rgba(0,0,0,.06)";
    row.style.marginBottom = "8px";
    row.onmouseenter = () => row.style.background = "rgba(0,0,0,.03)";
    row.onmouseleave = () => row.style.background = "transparent";

    const checked = assigned.has(p.id);
    row.innerHTML = `
      <input type="checkbox" data-pid="${p.id}" ${checked ? "checked" : ""} />
      <div style="display:flex;flex-direction:column;gap:2px;min-width:0;">
        <div style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(p.name || p.id)}</div>
        <div class="tiny" style="opacity:.8;">Project ID: ${escapeHtml(p.id)}</div>
      </div>
    `;
    box.appendChild(row);
  });
}

function toggleAssignProjectsSelectAll(val){
  const box = document.getElementById("assignProjectsList");
  if(!box) return;
  box.querySelectorAll('input[type="checkbox"][data-pid]').forEach(cb => cb.checked = !!val);
}

function saveAssignProjects(){
  if (!canUser("adminApproveUsers")) return alert("No permission.");
  const u = state.users.find(x => x.id === __assignProjectsUserId);
  if(!u) return;

  const box = document.getElementById("assignProjectsList");
  const selected = [];
  if(box){
    box.querySelectorAll('input[type="checkbox"][data-pid]').forEach(cb => {
      if(cb.checked) selected.push(cb.getAttribute("data-pid"));
    });
  }

  u.projectAssignments = selected;
  saveState();

  // Resolve "assign projects" reminder notifications for this user
  try { resolveNotificationsByActionKey(`assign_projects:${u.id}`); } catch(e){}

  // Notify the user (system message)
  try {
    const names = selected.map(pid => (state.projects.find(p => p.id === pid)?.name || pid)).slice(0, 6);
    const more = selected.length > 6 ? ` (+${selected.length-6} more)` : "";
    SBTS_ACTIVITY.pushNotification({
      category: "system",
      scope: "user",
      recipients: [u.id],
      title: "Project assignments updated",
      message: selected.length ? `You were assigned to: ${names.join(", ")}${more}` : "You currently have no assigned projects.",
      requiresAction: false,
      resolved: true,
      actorId: getCurrentUserIdStable()
    });
  } catch(e){}

  closeModal("assignProjectsModal");
  toast("Assignments saved ✅");
  renderUsers();
  updateNotificationsBadge();
}

function rejectUser(userId) {
  if (!canUser("adminApproveUsers")) return alert("No permission.");
  const user = state.users.find((u) => u.id === userId);
  if (!user) return;

  const reason = prompt("Reject reason (required):", "") || "";
  if (!reason.trim()) return alert("Reject reason is required.");

  user.status = "rejected";
  user.rejectionReason = reason.trim();
  saveState();

// Sync related NEW_USER request in Inbox (if exists)
try {
  ensureRequestsArray();
  const r = (state.requests || []).find(x => x && x.type === "NEW_USER" && x.meta && x.meta.userId === userId && isRequestOpen(x.status || "pending"));
  if (r) setRequestStatus(r.id, "rejected");
} catch(e) {}

  // Resolve pending approval notifications for approvers
  resolveNotificationsByActionKey(`user_approval:${userId}`);

  // Notify the user
  addNotification([userId], {
    type: "user_status",
    title: "Account rejected",
    message: `Your registration request was rejected. Reason: ${reason.trim()}`,
    requiresAction: false,
    actorId: state.currentUser?.id || null,
  });

  renderUsers();
  updateNotificationsBadge();
}


function openEditPermissionsModal(userId) {
  if (!canUser("adminApproveUsers")) return alert("No permission.");
  const user = state.users.find((u) => u.id === userId);
  if (!user) return;

  currentPermissionsUserId = userId;
  document.getElementById("editPermissionsUserInfo").textContent =
    `${user.fullName || user.username} (${roleLabel(user.role)})`;

  const tbody = document.getElementById("permissionsMatrixBody");
  tbody.innerHTML = "";

  const perms = state.permissions[userId] || {};
  PERMISSION_DEFS.forEach((p) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${p.label}</td>
      <td><input type="checkbox" data-perm-key="${p.key}" ${perms[p.key] ? "checked" : ""} /></td>
    `;
    tbody.appendChild(tr);
  });

  openModal("editPermissionsModal");
}

function saveUserPermissions() {
  if (!currentPermissionsUserId) return;

  const checkboxes = document.querySelectorAll("#permissionsMatrixBody input[type='checkbox']");
  const perms = {};
  checkboxes.forEach((cb) => {
    const key = cb.getAttribute("data-perm-key");
    perms[key] = cb.checked;
  });

  state.permissions[currentPermissionsUserId] = perms;
  saveState();
  closeModal("editPermissionsModal");
  renderUsers();
}

/* ==========================
   SETTINGS + BRANDING
========================== */
function applyThemePreset(key, silent = false) {
  const p = THEME_PRESETS[key] || THEME_PRESETS.classic;
  state.themePreset = key;
  state.ui = state.ui || {};
  state.ui.themePreset = state.themePreset;

  document.documentElement.style.setProperty("--primary", p.primary);
  document.documentElement.style.setProperty("--primary-strong", p.primaryStrong);
  document.documentElement.style.setProperty("--app-bg", p.bodyBg);
  document.documentElement.style.setProperty("--thead-bg", p.tableHead);

  const themePicker = document.getElementById("themePicker");
  if (themePicker) themePicker.value = p.primary;

  state.themeColor = p.primary;
  if (!silent) saveState();

  applyBrandingToHeader();
}

function applyThemeColor(color) {
  state.themeColor = color;
  state.ui = state.ui || {};
  state.ui.themeColor = state.themeColor;
  document.documentElement.style.setProperty("--primary", color);
}

function updateThemeColor() {
  const color = document.getElementById("themePicker").value;
  state.themePreset = "custom";
  state.ui = state.ui || {};
  state.ui.themePreset = state.themePreset;
  applyThemeColor(color);
  saveState();
  applyBrandingToHeader();
}

function updateProfileImage(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    if (!state.currentUser) return;
    state.currentUser.profileImage = e.target.result;
    const u = state.users.find(x => x.id === state.currentUser.id);
    if (u) u.profileImage = state.currentUser.profileImage;
    saveState();
    applySidebarUserBox();
    alert("Profile image saved.");
  };
  reader.readAsDataURL(file);
}

function applyFontSize(px) {
  const prevPx = Number(state.fontSize) || 14;
  const prevScale = prevPx / 14;

  state.fontSize = Number(px) || 14;
  const newScale = state.fontSize / 14;

  // Keep root font-size stable-ish, but apply a true global scale across the UI.
  // This ensures font size affects buttons, tables, modals, etc. (even if some CSS uses px).
  applyGlobalFontScale(prevScale, newScale);

  const v = document.getElementById("fontSizeValue");
  if (v) v.textContent = String(state.fontSize);
}

/**
 * Applies font scaling across the whole app without compounding.
 * Stores each element's "base" font size (at scale=1) in data-base-font.
 */
function applyGlobalFontScale(prevScale, newScale) {
  try {
    const root = document.getElementById("appRoot") || document.body;
    const all = root.querySelectorAll("*");
    all.forEach(el => {
      // Skip script/style/meta, and SVG text is rare; safe to skip svg entirely.
      const tag = (el.tagName || "").toLowerCase();
      if (!tag || tag === "script" || tag === "style" || tag === "meta" || tag === "link") return;
      if (tag === "svg") return;

      const cs = window.getComputedStyle(el);
      const cur = parseFloat(cs.fontSize);
      if (!cur || Number.isNaN(cur)) return;

      let base = el.dataset.baseFont ? parseFloat(el.dataset.baseFont) : NaN;
      if (!base || Number.isNaN(base)) {
        // Derive base by reversing the previous scale
        base = cur / (prevScale || 1);
        el.dataset.baseFont = String(base);
      }
      const next = base * (newScale || 1);
      el.style.fontSize = `${next}px`;
    });

    // Also set root to match (helps for newly rendered elements that inherit)
    document.documentElement.style.fontSize = `${14 * (newScale || 1)}px`;
  } catch (e) {
    console.warn("applyGlobalFontScale failed", e);
  }
}

function updateFontSize(px) {
  applyFontSize(px);
  saveState();
}



function toggleEnforcePhaseOrder(enabled) {
  // Admin only
  if (normalizeRole(state.currentUser?.role) !== 'admin') {
    toast('Admin only.');
    hydrateSettingsInputs();
    return;
  }
  state.ui.enforcePhaseOrder = !!enabled;
  saveState();
  toast('Enforce phase order ' + (enabled ? 'enabled' : 'disabled') + '.');
}
function updateJobTitle(value) {
  if (!state.currentUser) return;
  state.currentUser.jobTitle = value;
  const u = state.users.find(x => x.id === state.currentUser.id);
  if (u) u.jobTitle = value;
  saveState();
  applySidebarUserBox();
}

function hydrateSettingsInputs() {
  // Branding inputs
  const b = state.branding;
  const setVal = (id, value) => { const el = document.getElementById(id); if (el) el.value = value || ""; };

  const canBrand = canUser("editBranding");

  // Training page visibility (permission toggle)
  const tglTraining = document.getElementById("toggleTrainingPage");

  // Workflow behavior (permission)
  const tglEnforce = document.getElementById("toggleEnforcePhaseOrder");
  if (tglEnforce) {
    tglEnforce.checked = state.ui?.enforcePhaseOrder !== false;
    tglEnforce.disabled = !canUser("manageWorkflowControl");
    tglEnforce.onchange = () => toggleEnforcePhaseOrder(tglEnforce.checked);
  }

  if (tglTraining) {
    tglTraining.checked = state.ui?.showTrainingPage !== false;
    tglTraining.disabled = !canUser("manageTrainingVisibility");
    tglTraining.onchange = () => setShowTrainingPage(tglTraining.checked);
  }

  // Backup weekly local toggle (Admin only)
  const tglWeekly = document.getElementById("toggleWeeklyBackup");
  if (tglWeekly) {
    tglWeekly.checked = sbtsIsWeeklyBackupEnabled();
    tglWeekly.disabled = !sbtsIsAdmin();
    tglWeekly.onchange = () => sbtsToggleWeeklyBackup(tglWeekly.checked);
  }




  setVal("brandingProgramTitle", b.programTitle);
  setVal("brandingProgramSub", b.programSubtitle);
  setVal("brandingSiteTitle", b.siteTitle);
  setVal("brandingSiteSub", b.siteSubtitle);
  setVal("brandingCompanyName", b.companyName);
  setVal("brandingCompanySub", b.companySub);

  // Lock Branding section unless permitted
  try {
    const sec = document.getElementById("brandingSection");
    if (sec) {
      const controls = Array.from(sec.querySelectorAll("input"));
      controls.forEach((el) => {
        if (!canBrand) {
          el.disabled = true;
          el.title = "No permission";
        } else {
          el.disabled = false;
        }
      });
      let note = sec.querySelector(".perm-note");
      if (!canBrand) {
        if (!note) {
          note = document.createElement("div");
          note.className = "tiny perm-note";
          note.style.marginTop = "8px";
          note.textContent = "Locked: Admin must grant 'Edit app/project branding' permission.";
          sec.appendChild(note);
        }
      } else if (note) {
        note.remove();
      }
    }
  } catch (e) {}

  // font
  const fontSlider = document.getElementById("fontSizeSlider");
  if (fontSlider) fontSlider.value = String(state.fontSize || 14);

  const jobTitleInput = document.getElementById("jobTitleInput");
  if (jobTitleInput && state.currentUser) jobTitleInput.value = state.currentUser.jobTitle || "";

  applyBrandingToHeader();

  // Apply section visibility based on permissions
  applySettingsPermissionsUI();

// Roles & Specialties manager
try {
  renderRolesCatalogManager();
  const btn = document.getElementById("btnAddRole");
  if (btn) btn.onclick = addCatalogRole;

  // Patch 47.0: Admin inspector tools
  sbtsBindNotifInspectorButtons();

} catch (e) { console.warn(e); }


  try { hydrateNotifRulesSettings(); } catch(e){}
}

function applySettingsPermissionsUI() {
  // Hide admin/system sections unless permitted
  const setVis = (id, ok) => {
    const el = document.getElementById(id);
    if (el) el.style.display = ok ? "" : "none";
  };
  setVis("workflowLiteSection", canUser("manageWorkflowControl"));
  setVis("reportsCardsShortcutSection", canUser("manageReportsCards"));
  setVis("certificateShortcutSection", canUser("manageCertificateSettings"));
  setVis("trainingVisibilitySection", canUser("manageTrainingVisibility"));
  // Backup & Restore is Admin only
  setVis("backupRestoreSection", state.currentUser?.role === "admin");
  setVis("rolesCatalogCard", canUser("manageRolesCatalog"));
  setVis("finalApprovalsManagerSection", canUser("manageFinalApprovals"));
  setVis("faSafeModeSection", canUser("manageFinalApprovals"));
  // Branding is admin/system-level; hide entirely unless permitted
  setVis("brandingSection", canUser("editBranding"));
}

function updateBranding(key, value) {
  if (!canUser("editBranding")) return toast("No permission.");
  state.branding[key] = value;
  saveState();
  applyBrandingToHeader();
}

function updateCompanyLogo(event) {
  if (!canUser("editBranding")) return toast("No permission.");
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    state.branding.companyLogo = e.target.result;
    saveState();
    applyBrandingToHeader();
    alert("Company logo saved.");
  };
  reader.readAsDataURL(file);
}



/* ==========================
   TAG SETTINGS PAGE + TAG PRINTING
========================== */
function ensureTagThemeDefaults(){
  if (!state.tagTheme) state.tagTheme = { color: '#0F6D8C', audit: [] };
  if (!state.tagTheme.color) state.tagTheme.color = '#0F6D8C';
  if (!Array.isArray(state.tagTheme.audit)) state.tagTheme.audit = [];
}

function getTagMonthYear(){
  const d = new Date();
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yy = String(d.getFullYear());
  return `${mm}/${yy}`;
}

function hydrateTagSettingsUI(){
  if (!canUser('manageTagSettings')){ toast('No permission'); openPage('dashboardPage'); return; }
  ensureTagThemeDefaults();
  const picker = document.getElementById('tagColorPicker');
  const hex = document.getElementById('tagColorHex');
  if (picker){ picker.value = state.tagTheme.color || '#0F6D8C'; }
  if (hex){ hex.textContent = state.tagTheme.color || '#0F6D8C'; }
  renderTagSettingsPreview();
  renderTagThemeAudit();
}

function saveTagThemeFromUI(){
  if (!canUser('manageTagSettings')) return toast('No permission');
  ensureTagThemeDefaults();
  const picker = document.getElementById('tagColorPicker');
  if (!picker) return;
  const newColor = picker.value || '#0F6D8C';
  const oldColor = state.tagTheme.color || '#0F6D8C';
  state.tagTheme.color = newColor;
  const by = state.currentUser?.fullName || state.currentUser?.username || 'Admin';
  state.tagTheme.audit.unshift({ ts: new Date().toISOString(), by, old: oldColor, new: newColor });
  state.tagTheme.audit = state.tagTheme.audit.slice(0, 200);
  saveState();
  toast('Tag theme saved');
  const hex = document.getElementById('tagColorHex');
  if (hex) hex.textContent = newColor;
  renderTagSettingsPreview();
  renderTagThemeAudit();
}

function renderTagSettingsPreview(){
  ensureTagThemeDefaults();
  const wrap = document.getElementById('tagPreviewWrap');
  if (!wrap) return;
  // dummy preview based on current project if exists, else static
  const sample = {
    area: 'SRU-3',
    equipment: 'D-301',
    blindId: 'BL-001'
  };
  wrap.innerHTML = buildTagCardHTMLFromTemplate(sample, state.tagTheme.color, getTagMonthYear());
}

function renderTagThemeAudit(){
  const box = document.getElementById('tagAuditTable');
  if (!box) return;
  ensureTagThemeDefaults();
  const rows = (state.tagTheme.audit||[]).map(a => {
    const dt = new Date(a.ts);
    const dstr = isNaN(dt.getTime()) ? (a.ts||'') : dt.toLocaleString();
    return `<tr>
      <td>${escapeHtml(dstr)}</td>
      <td>${escapeHtml(a.by||'')}</td>
      <td><span style="display:inline-block;width:14px;height:14px;border:1px solid #cbd5e1;border-radius:3px;background:${a.old||''};"></span> ${escapeHtml(a.old||'')}</td>
      <td><span style="display:inline-block;width:14px;height:14px;border:1px solid #cbd5e1;border-radius:3px;background:${a.new||''};"></span> ${escapeHtml(a.new||'')}</td>
    </tr>`;
  }).join('');
  box.innerHTML = `
    <table class="main-table">
      <thead><tr><th>Date</th><th>By</th><th>Old</th><th>New</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="4">No changes yet.</td></tr>'}</tbody>
    </table>
  `;
}


// ==========================
// TAG TEMPLATE (DESIGNER)
// ==========================
// Patch40: Support 3 editable templates + 1 locked default.
// Backward-compatible: if only the old single template exists, we seed Template 1 from it.
const TAG_TEMPLATE_KEY = "sbts_tag_template_v1"; // legacy
const TAG_TEMPLATES_KEY = "sbts_tag_templates_v2";
const TAG_ACTIVE_TEMPLATE_KEY = "sbts_tag_active_template_v2";

function getDefaultTemplateId(){ return 'default'; }
function getEditableTemplateIds(){ return ['tpl1','tpl2','tpl3']; }

function loadTagTemplates(){
  try {
    const raw = SBTS_UTILS.LS.get(TAG_TEMPLATES_KEY);
    if (raw){
      const obj = JSON.parse(raw);
      if (obj && obj.slots){
        // Ensure required slots exist
        const slots = obj.slots;
        slots[getDefaultTemplateId()] = slots[getDefaultTemplateId()] || getDefaultTagTemplate();
        getEditableTemplateIds().forEach(id => { slots[id] = slots[id] || getDefaultTagTemplate(); });
        return { slots, updatedAt: obj.updatedAt || new Date().toISOString() };
      }
    }
  } catch(e){}

  // Seed from legacy single-template, if any
  let seed = null;
  try {
    const legacyRaw = SBTS_UTILS.LS.get(TAG_TEMPLATE_KEY);
    if (legacyRaw) seed = JSON.parse(legacyRaw);
  } catch(e){}

  const slots = {};
  slots[getDefaultTemplateId()] = getDefaultTagTemplate();
  slots['tpl1'] = seed && seed.elements ? seed : getDefaultTagTemplate();
  slots['tpl2'] = getDefaultTagTemplate();
  slots['tpl3'] = getDefaultTagTemplate();

  const obj = { slots, updatedAt: new Date().toISOString() };
  try { SBTS_UTILS.LS.set(TAG_TEMPLATES_KEY, JSON.stringify(obj)); } catch(e){}
  // Set active
  try { if (!SBTS_UTILS.LS.get(TAG_ACTIVE_TEMPLATE_KEY)) SBTS_UTILS.LS.set(TAG_ACTIVE_TEMPLATE_KEY,'tpl1'); } catch(e){}
  return obj;
}

function saveTagTemplatesObj(obj){
  try { SBTS_UTILS.LS.set(TAG_TEMPLATES_KEY, JSON.stringify(obj)); } catch(e){}
}

function getActiveTemplateId(){
  const id = SBTS_UTILS.LS.get(TAG_ACTIVE_TEMPLATE_KEY) || 'tpl1';
  if (id === getDefaultTemplateId()) return getDefaultTemplateId();
  return getEditableTemplateIds().includes(id) ? id : 'tpl1';
}

function setActiveTemplateId(id){
  const safe = (id === getDefaultTemplateId()) ? getDefaultTemplateId() : (getEditableTemplateIds().includes(id) ? id : 'tpl1');
  SBTS_UTILS.LS.set(TAG_ACTIVE_TEMPLATE_KEY, safe);
}

function ensureTagTemplateDefaults(){
  if (!state.tagTemplate) {
    const t = loadTagTemplate();
    state.tagTemplate = t;
  }
}

function getDefaultTagTemplate(){
  // Patch39: Tag template uses centimeters internally (7cm x 11cm),
  // but print output is locked to 70mm x 110mm (label printer ready).
  return {
    sizeCm: { w: 7, h: 11 },
    fontFamily: 'Arial',
    // Elements positions/sizes in cm
    elements: {
      // hole is center-based
      hole: { x: 3.5, y: 0.85, d: 1.2 },
      logo: { x: 5.55, y: 0.45, w: 1.05, h: 1.05 },
      title: { x: 0.35, y: 2.05, w: 6.3, h: 1.4, fontPx: 32, weight: 900, align: "center" },
      qr:    { x: 1.05, y: 3.55, w: 4.9, h: 4.9, pad: 0.18, radius: 0.25 },

      // Patch39: each row is its own box (independent font/position)
      areaLine: { x: 0.55, y: 8.40, w: 5.9, h: 0.65, fontPx: 22, align: "left", keyText: "Area", keyWeight: 700, valWeight: 700 },
      equLine:  { x: 0.55, y: 9.10, w: 5.9, h: 0.65, fontPx: 22, align: "left", keyText: "Line", keyWeight: 700, valWeight: 700 },
      idLine:   { x: 0.55, y: 9.80, w: 5.9, h: 0.65, fontPx: 22, align: "left", keyText: "ID",   keyWeight: 900, valWeight: 900 },

      // date: bottom-left (MM/YYYY)
      date:  { x: 0.25, y: 10.55, w: 2.4, h: 0.50, fontPx: 11, weight: 700, align: "left" }
    }
  };
}

function loadTagTemplate(){
  // Active slot
  const obj = loadTagTemplates();
  const activeId = getActiveTemplateId();
  const slots = obj?.slots || {};
  const t = slots[activeId] || slots['tpl1'] || getDefaultTagTemplate();
  // light validation
  if (!t || !t.sizeCm || !t.elements) return getDefaultTagTemplate();
  return t;
}

function saveTagTemplate(t){
  // Save into the active editable template slot (never overwrite locked default)
  const activeId = getActiveTemplateId();
  if (activeId === getDefaultTemplateId()) {
    // locked
    state.tagTemplate = t;
    saveState();
    return;
  }
  const obj = loadTagTemplates();
  obj.slots = obj.slots || {};
  obj.slots[activeId] = t;
  obj.updatedAt = new Date().toISOString();
  saveTagTemplatesObj(obj);
  // Keep legacy key updated for backward compatibility/tools
  try { SBTS_UTILS.LS.set(TAG_TEMPLATE_KEY, JSON.stringify(t)); } catch(e){}
  state.tagTemplate = t;
  saveState();
}

function resetTagTemplate(){
  const activeId = getActiveTemplateId();
  if (activeId === getDefaultTemplateId()) {
    // Switch to Template 1 to allow editing
    setActiveTemplateId('tpl1');
  }
  const t = getDefaultTagTemplate();
  saveTagTemplate(t);
  return t;
}

function tdSelectTemplate(slotId){
  const safe = (slotId===getDefaultTemplateId()) ? getDefaultTemplateId() : (getEditableTemplateIds().includes(slotId)?slotId:'tpl1');
  setActiveTemplateId(safe);
  // Load template into state
  if (safe === getDefaultTemplateId()) {
    state.tagTemplate = getDefaultTagTemplate();
  } else {
    const obj = loadTagTemplates();
    state.tagTemplate = (obj?.slots && obj.slots[safe]) ? obj.slots[safe] : getDefaultTagTemplate();
  }
  saveState();
  hydrateTagDesignerUI();
}

function cm(v){
  return (Number(v)||0) + "cm";
}

function safeCssFont(f){
  const s = String(f||"Arial");
  return s.replace(/[^a-zA-Z0-9 ,\-'"]+/g, "");
}

function styleBox(el){
  return `left:${cm(el.x)};top:${cm(el.y)};width:${cm(el.w)};height:${cm(el.h)};`;
}

function buildTagCardHTMLFromTemplate(data, color, monthYear){
  ensureTagTemplateDefaults();
  const t = state.tagTemplate || getDefaultTagTemplate();
  const bg = color || '#0F6D8C';

  const qrValue = data.qrValue || data.publicLink || '';
  const qrImg = qrValue ? `<img alt="QR" src="${buildQrImageUrl(qrValue)}" />` : ``;
  const logoUrl = (state?.branding?.companyLogo) ? state.branding.companyLogo : "";

  const e = t.elements || {};
  const hole = e.hole || {x:3.5,y:0.85,d:1.2};
  const logo = e.logo || {x:5.55,y:0.45,w:1.05,h:1.05};
  const title = e.title || {x:0.35,y:2.05,w:6.3,h:1.4,fontPx:32,weight:900,align:'center'};
  const qr = e.qr || {x:1.05,y:3.55,w:4.9,h:4.9,pad:0.18,radius:0.25};
  const areaLine = e.areaLine || {x:0.55,y:8.40,w:5.9,h:0.65,fontPx:22,align:'left',keyText:'Area',keyWeight:700,valWeight:700};
  const equLine  = e.equLine  || {x:0.55,y:9.10,w:5.9,h:0.65,fontPx:22,align:'left',keyText:'Line',keyWeight:700,valWeight:700};
  const idLine   = e.idLine   || {x:0.55,y:9.80,w:5.9,h:0.65,fontPx:22,align:'left',keyText:'ID',keyWeight:900,valWeight:900};
  const date = e.date || {x:0.25,y:10.55,w:2.4,h:0.50,fontPx:11,weight:700,align:'left'};

  const tagW = (t.sizeCm && t.sizeCm.w) ? t.sizeCm.w : 7;
  const tagH = (t.sizeCm && t.sizeCm.h) ? t.sizeCm.h : 11;

  return `
    <div class="sbts-tag-preview">
      <div class="sbts-tag-card sbts-tag-card-tpl" style="background:${bg};width:${cm(tagW)};height:${cm(tagH)};font-family:${safeCssFont(t.fontFamily||'Arial')};">
        <div class="sbts-tag-hole-center" style="left:${cm(hole.x)};top:${cm(hole.y)};width:${cm(hole.d)};height:${cm(hole.d)};"></div>

        <div class="sbts-tag-logo-corner" style="${styleBox(logo)}${logoUrl ? `background-image:url('${logoUrl}');` : ''}"></div>

        <div class="sbts-tag-title" style="${styleBox(title)}font-size:${title.fontPx||32}px;font-weight:${title.weight||900};text-align:${title.align||'center'};">
          Smart Blind Tag
        </div>

        <div class="sbts-tag-qr" style="${styleBox(qr)}border-radius:${cm(qr.radius||0.25)};padding:${cm(qr.pad||0.18)};">
          <div class="sbts-tag-qr-inner">${qrImg}</div>
        </div>

        <div class="sbts-tag-row sbts-tag-area" style="${styleBox(areaLine)}font-size:${areaLine.fontPx||22}px;font-weight:400;text-align:${areaLine.align||'left'};">
          <span class="sbts-tag-row-key" style="font-weight:${Number(areaLine.keyWeight)||400};">${escapeHtml(areaLine.keyText||'Area')}:</span>
          <span class="sbts-tag-row-val" style="font-weight:${Number(areaLine.valWeight)||400};">${escapeHtml(data.area||'-')}</span>
        </div>

        <div class="sbts-tag-row sbts-tag-line" style="${styleBox(equLine)}font-size:${equLine.fontPx||22}px;font-weight:400;text-align:${equLine.align||'left'};">
          <span class="sbts-tag-row-key" style="font-weight:${Number(equLine.keyWeight)||400};">${escapeHtml(equLine.keyText||'Line')}:</span>
          <span class="sbts-tag-row-val" style="font-weight:${Number(equLine.valWeight)||400};">${escapeHtml(data.line||'-')}</span>
        </div>

        <div class="sbts-tag-row sbts-tag-id" style="${styleBox(idLine)}font-size:${idLine.fontPx||22}px;font-weight:400;text-align:${idLine.align||'left'};">
          <span class="sbts-tag-row-key" style="font-weight:${Number(idLine.keyWeight)||400};">${escapeHtml(idLine.keyText||'ID')}:</span>
          <span class="sbts-tag-row-val" style="font-weight:${Number(idLine.valWeight)||400};">${escapeHtml(data.blindId||'-')}</span>
        </div>

        <div class="sbts-tag-date-corner" style="${styleBox(date)}font-size:${date.fontPx||11}px;font-weight:${date.weight||700};text-align:${date.align||'left'};">
          ${escapeHtml(monthYear||'')}
        </div>
      </div>
    </div>
  `;
}


// ==========================
// TAG DESIGNER UI (Drag/Resize)
// ==========================
let TD_RUNTIME = { inited:false, selected:null, dragging:null, resizing:null, snap:true, gridMm: 1, scalePxPerCm: 52 };

function hydrateTagDesignerUI(){
  ensureTagThemeDefaults();
  ensureTagTemplateDefaults();

  if (!canUser('manageTagSettings')){ toast('No permission'); openPage('dashboardPage'); return; }

  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;

  // init controls once
  if (!TD_RUNTIME.inited){
    TD_RUNTIME.inited = true;
    const bg = document.getElementById('tdBgColor');
    if (bg){
      bg.addEventListener('input', () => tdUpdateBgColor(bg.value));
    }
    const font = document.getElementById('tdFontFamily');
    if (font){
      font.addEventListener('change', () => tdSetFontFamily(font.value));
    }
    const snap = document.getElementById('tdSnap');
    if (snap){
      snap.addEventListener('change', () => tdToggleSnap(snap.checked));
      TD_RUNTIME.snap = snap.checked;
    }

    // global mouse handlers
    window.addEventListener('mousemove', tdOnMouseMove);
    window.addEventListener('mouseup', tdOnMouseUp);
  }

  // sync controls
  // templates list (3 editable + 1 locked default)
  try {
    const list = document.getElementById('tdTemplateList');
    if (list){
      const activeId = getActiveTemplateId();
      const mkItem = (id, name, badgeText, badgeClass) => {
        const d = document.createElement('div');
        d.className = 'td-template-item' + (activeId===id ? ' active' : '');
        d.innerHTML = `
          <div class="td-template-name">${escapeHtml(name)}</div>
          <div class="td-template-badge ${badgeClass||''}">${escapeHtml(badgeText||'')}</div>
        `;
        d.onclick = () => tdSelectTemplate(id);
        return d;
      };
      list.innerHTML = '';
      list.appendChild(mkItem(getDefaultTemplateId(), 'Default template', 'Locked', 'locked'));
      list.appendChild(mkItem('tpl1', 'Template 1', activeId==='tpl1'?'Selected':'Editable', ''));
      list.appendChild(mkItem('tpl2', 'Template 2', activeId==='tpl2'?'Selected':'Editable', ''));
      list.appendChild(mkItem('tpl3', 'Template 3', activeId==='tpl3'?'Selected':'Editable', ''));
    }
  } catch(e){}
  const bg = document.getElementById('tdBgColor');
  const bgHex = document.getElementById('tdBgColorHex');
  if (bg) bg.value = state.tagTheme.color || '#0F6D8C';
  if (bgHex) bgHex.textContent = (state.tagTheme.color || '#0F6D8C');

  const font = document.getElementById('tdFontFamily');
  if (font) font.value = (state.tagTemplate.fontFamily || 'Arial');

  const t = state.tagTemplate || getDefaultTagTemplate();

  // sync tag size (mm)
  try {
    const wMm = Math.round((t.sizeCm?.w || 7) * 10);
    const hMm = Math.round((t.sizeCm?.h || 11) * 10);
    const wEl = document.getElementById('tdTagWmm');
    const hEl = document.getElementById('tdTagHmm');
    if (wEl) wEl.value = String(wMm);
    if (hEl) hEl.value = String(hMm);
  } catch(e){}

  // sync row label/style controls
  try {
    const a=(id,val)=>{const el=document.getElementById(id); if(el) el.value=String(val??'');};
    const s=(id,val)=>{const el=document.getElementById(id); if(el) el.value=String(val??'');};
    const area = t.elements?.areaLine || {};
    const line = t.elements?.equLine || {};
    const _id  = t.elements?.idLine  || {};

    a('tdLblArea', area.keyText || 'Area');
    a('tdLblLine', line.keyText || 'Line');
    a('tdLblId',   _id.keyText  || 'ID');

    s('tdLblAreaW', area.keyWeight || 700);
    s('tdValAreaW', area.valWeight || 700);
    s('tdLblLineW', line.keyWeight || 700);
    s('tdValLineW', line.valWeight || 700);
    s('tdLblIdW',   _id.keyWeight  || 900);
    s('tdValIdW',   _id.valWeight  || 900);

    s('tdAreaAlign', area.align || 'left');
    s('tdLineAlign', line.align || 'left');
    s('tdIdAlign',   _id.align  || 'left');

    const gridSel = document.getElementById('tdGridMm');
    if (gridSel) gridSel.value = String(TD_RUNTIME.gridMm || 1);
  } catch(e){}
  // set sliders
  try {
    const titlePx = t.elements?.title?.fontPx || 32;
    const areaPx = t.elements?.areaLine?.fontPx || 22;
    const equPx  = t.elements?.equLine?.fontPx  || 22;
    const idPx   = t.elements?.idLine?.fontPx   || 22;
    const datePx = t.elements?.date?.fontPx || 11;
    const a=(id,val)=>{const el=document.getElementById(id); if(el) el.value=val;};
    const s=(id,val)=>{const el=document.getElementById(id); if(el) el.textContent=String(val);};
    a('tdTitleSize', titlePx); s('tdTitleSizeVal', titlePx);
    a('tdAreaSize', areaPx); s('tdAreaSizeVal', areaPx);
    a('tdEquSize',  equPx);  s('tdEquSizeVal',  equPx);
    a('tdIdSize',   idPx);   s('tdIdSizeVal',   idPx);
    a('tdDateSize', datePx); s('tdDateSizeVal', datePx);
  } catch(e){}

  renderTDTagAudit();
  tdRenderStage();
  tdSelect(null);
}

function tdPx(cmVal){
  return (Number(cmVal)||0) * TD_RUNTIME.scalePxPerCm;
}
function tdMmToPx(mmVal){
  // 10mm = 1cm
  return tdPx((Number(mmVal)||0) / 10);
}
function tdCm(pxVal){
  return (Number(pxVal)||0) / TD_RUNTIME.scalePxPerCm;
}
function tdSnapPx(px){
  if (!TD_RUNTIME.snap) return px;
  const step = Math.max(1, tdMmToPx(TD_RUNTIME.gridMm || 1));
  return Math.round(px/step)*step;
}

function tdRenderStage(){
  ensureTagTemplateDefaults();
  ensureTagThemeDefaults();
  const t = state.tagTemplate || getDefaultTagTemplate();
  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;
  stage.innerHTML = '';

  const w = t.sizeCm?.w || 7;
  const h = t.sizeCm?.h || 11;

  // scale based on available width (keep readable on small screens)
  const wrap = stage.parentElement;
  const maxW = (wrap && wrap.clientWidth) ? wrap.clientWidth : 520;
  const targetPx = Math.min(maxW - 40, 420);
  TD_RUNTIME.scalePxPerCm = Math.max(34, Math.min(70, targetPx / w));

  stage.style.width = tdPx(w) + 'px';
  stage.style.height = tdPx(h) + 'px';
  stage.style.background = state.tagTheme.color || '#0F6D8C';

  // grid (visual)
  stage.style.backgroundImage = 'linear-gradient(to right, rgba(255,255,255,.10) 1px, transparent 1px), linear-gradient(to bottom, rgba(255,255,255,.10) 1px, transparent 1px)';
  const gridPx = tdMmToPx(TD_RUNTIME.gridMm || 1);
  stage.style.backgroundSize = (TD_RUNTIME.snap ? `${gridPx}px ${gridPx}px` : '0 0');

  const sample = { area:'SRU-3', line:'D-301', blindId:'BL-001', qrValue:'SBTS', publicLink:'SBTS' };
  const monthYear = getTagMonthYear();
  const logoUrl = (state?.branding?.companyLogo) ? state.branding.companyLogo : '';

  const el = t.elements || {};

  // helper to make element
  function mk(key, className, style, inner, resizable){
    const d = document.createElement('div');
    d.className = 'td-el ' + className;
    d.dataset.key = key;
    d.style.cssText = style;
    d.innerHTML = inner || '';
    d.addEventListener('mousedown', (ev)=>tdOnMouseDown(ev, key));
    d.addEventListener('click', (ev)=>{ev.stopPropagation(); tdSelect(key);});
    if (resizable){
      const h = document.createElement('div');
      h.className = 'td-handle br';
      h.addEventListener('mousedown', (ev)=>tdOnResizeDown(ev, key));
      d.appendChild(h);
    }
    return d;
  }

  // stage click to clear selection
  stage.onclick = ()=>tdSelect(null);

  // Hole (center-based)
  const def = getDefaultTagTemplate().elements;
  const hole = el.hole || def.hole;
  const holeLeft = tdPx(hole.x) - tdPx(hole.d)/2;
  const holeTop = tdPx(hole.y) - tdPx(hole.d)/2;
  stage.appendChild(mk('hole','td-hole', `left:${holeLeft}px;top:${holeTop}px;width:${tdPx(hole.d)}px;height:${tdPx(hole.d)}px;`, '', true));

  // Logo
  const logo = el.logo || def.logo;
  const logoBg = logoUrl ? `background-image:url('${logoUrl}');` : '';
  stage.appendChild(mk('logo','td-logo', `left:${tdPx(logo.x)}px;top:${tdPx(logo.y)}px;width:${tdPx(logo.w)}px;height:${tdPx(logo.h)}px;${logoBg}`, '', true));

  // Title
  const title = el.title || def.title;
  stage.appendChild(mk('title','td-title', `left:${tdPx(title.x)}px;top:${tdPx(title.y)}px;width:${tdPx(title.w)}px;height:${tdPx(title.h)}px;font-family:${t.fontFamily||'Arial'};font-size:${title.fontPx||32}px;font-weight:${title.weight||900};text-align:${title.align||'center'};`, 'Smart Blind Tag', false));

  // QR
  const qr = el.qr || def.qr;
  const qrUrl = buildQrImageUrl(sample.qrValue);
  stage.appendChild(mk('qr','td-qr', `left:${tdPx(qr.x)}px;top:${tdPx(qr.y)}px;width:${tdPx(qr.w)}px;height:${tdPx(qr.h)}px;padding:${tdPx(qr.pad)}px;border-radius:${tdPx(qr.radius||0.25)}px;`, `<div class="td-qr-inner"><img alt="QR" src="${qrUrl}"/></div>`, true));

  // Rows (each in its own box)
  const areaLine = el.areaLine || def.areaLine;
  stage.appendChild(mk('areaLine','td-row', `left:${tdPx(areaLine.x)}px;top:${tdPx(areaLine.y)}px;width:${tdPx(areaLine.w)}px;height:${tdPx(areaLine.h)}px;font-family:${t.fontFamily||'Arial'};font-size:${areaLine.fontPx||22}px;font-weight:${areaLine.valWeight||800};text-align:${areaLine.align||'left'};`,
    `<div>${escapeHtml(areaLine.keyText||'Area')}: ${escapeHtml(sample.area)}</div>`, false));

  const equLine = el.equLine || def.equLine;
  stage.appendChild(mk('equLine','td-row', `left:${tdPx(equLine.x)}px;top:${tdPx(equLine.y)}px;width:${tdPx(equLine.w)}px;height:${tdPx(equLine.h)}px;font-family:${t.fontFamily||'Arial'};font-size:${equLine.fontPx||22}px;font-weight:${equLine.valWeight||800};text-align:${equLine.align||'left'};`,
    `<div>${escapeHtml(equLine.keyText||'Line')}: ${escapeHtml(sample.line)}</div>`, false));

  const idLine = el.idLine || def.idLine;
  stage.appendChild(mk('idLine','td-row', `left:${tdPx(idLine.x)}px;top:${tdPx(idLine.y)}px;width:${tdPx(idLine.w)}px;height:${tdPx(idLine.h)}px;font-family:${t.fontFamily||'Arial'};font-size:${idLine.fontPx||22}px;font-weight:${idLine.valWeight||900};text-align:${idLine.align||'left'};`,
    `<div>${escapeHtml(idLine.keyText||'ID')}: ${escapeHtml(sample.blindId)}</div>`, false));

  // Date
  const date = el.date || def.date;
  stage.appendChild(mk('date','td-date', `left:${tdPx(date.x)}px;top:${tdPx(date.y)}px;width:${tdPx(date.w)}px;height:${tdPx(date.h)}px;font-family:${t.fontFamily||'Arial'};font-size:${date.fontPx||11}px;font-weight:${date.weight||700};text-align:${date.align||'left'};`, escapeHtml(monthYear), false));

  // apply selected outline
  tdApplySelectedOutline();
}

function tdApplySelectedOutline(){
  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;
  stage.querySelectorAll('.td-el').forEach(el=>{
    if (TD_RUNTIME.selected && el.dataset.key === TD_RUNTIME.selected) el.classList.add('selected');
    else el.classList.remove('selected');
  });
}

function tdSelect(key){
  TD_RUNTIME.selected = key;
  const lab = document.getElementById('tdSelectedName');
  if (lab) lab.textContent = key ? key.toUpperCase() : 'None';
  tdApplySelectedOutline();
}

function tdUpdateBgColor(color){
  ensureTagThemeDefaults();
  state.tagTheme.color = color;
  saveState();
  const hex = document.getElementById('tdBgColorHex');
  if (hex) hex.textContent = color;
  tdRenderStage();
}

function tdSetFontFamily(font){
  ensureTagTemplateDefaults();
  state.tagTemplate.fontFamily = font;
  saveTagTemplate(state.tagTemplate);
  tdRenderStage();
}

function tdSetFontSize(which, px){
  ensureTagTemplateDefaults();
  const t = state.tagTemplate;
  const v = parseInt(px,10) || 12;
  if (which === 'title') t.elements.title.fontPx = v;
  if (which === 'area') t.elements.areaLine.fontPx = v;
  if (which === 'equ')  t.elements.equLine.fontPx  = v;
  if (which === 'id')   t.elements.idLine.fontPx   = v;
  
  if (which === 'date') t.elements.date.fontPx = v;
  saveTagTemplate(t);
  const map = {title:['tdTitleSizeVal'], area:['tdAreaSizeVal'], equ:['tdEquSizeVal'], id:['tdIdSizeVal'], date:['tdDateSizeVal']};
  (map[which]||[]).forEach(id=>{const el=document.getElementById(id); if(el) el.textContent=String(v);});
  tdRenderStage();
}

function tdToggleSnap(on){
  TD_RUNTIME.snap = !!on;
  tdRenderStage();
}

function tdSetGridMm(mm){
  TD_RUNTIME.gridMm = Math.max(1, parseInt(mm,10) || 1);
  tdRenderStage();
}

function tdSaveThemeColor(){
  if (!canUser('manageTagSettings')) return toast('No permission');
  ensureTagThemeDefaults();
  const picker = document.getElementById('tdBgColor');
  const newColor = (picker && picker.value) ? picker.value : (state.tagTheme.color||'#0F6D8C');
  const oldColor = state.tagTheme.color || '#0F6D8C';
  state.tagTheme.color = newColor;
  const by = state.currentUser?.fullName || state.currentUser?.username || 'Admin';
  state.tagTheme.audit.unshift({ ts: new Date().toISOString(), by, old: oldColor, new: newColor, note: 'Tag Designer' });
  state.tagTheme.audit = state.tagTheme.audit.slice(0, 200);
  saveState();
  const hex = document.getElementById('tdBgColorHex');
  if (hex) hex.textContent = newColor;
  renderTDTagAudit();
  tdRenderStage();
  toast('Saved');
}

function tdApplyTagSizeMm(){
  if (!canUser('manageTagSettings')) return toast('No permission');
  ensureTagTemplateDefaults();
  const wEl = document.getElementById('tdTagWmm');
  const hEl = document.getElementById('tdTagHmm');
  const wMm = Math.max(40, Math.min(120, parseInt(wEl?.value||'70',10) || 70));
  const hMm = Math.max(40, Math.min(160, parseInt(hEl?.value||'110',10) || 110));
  state.tagTemplate.sizeCm = { w: wMm/10, h: hMm/10 };
  saveTagTemplate(state.tagTemplate);
  renderTDTagAudit();
  tdRenderStage();
  toast('Tag size updated');
}

function tdSetRowLabel(rowKey, text){
  ensureTagTemplateDefaults();
  const el = state.tagTemplate.elements[rowKey];
  if (!el) return;
  el.keyText = String(text||'').trim() || el.keyText || 'Label';
  saveTagTemplate(state.tagTemplate);
  tdRenderStage();
}

function tdSetRowWeights(rowKey, keyWeight, valWeight){
  ensureTagTemplateDefaults();
  const el = state.tagTemplate.elements[rowKey];
  if (!el) return;
  el.keyWeight = parseInt(keyWeight,10) || 400;
  el.valWeight = parseInt(valWeight,10) || 400;
  saveTagTemplate(state.tagTemplate);
  tdRenderStage();
}


function tdUiSetWeight(rowKey, part, weight){
  ensureTagTemplateDefaults();
  const el = state.tagTemplate.elements[rowKey];
  if (!el) return;
  const w = parseInt(weight,10) || 400;
  if (part === 'key') el.keyWeight = w;
  if (part === 'val') el.valWeight = w;
  saveTagTemplate(state.tagTemplate);
  tdRenderStage();
  tdUiRefreshTagRowButtons();
}

function tdUiSetAlign(rowKey, align){
  // keep hidden select in sync for hydration
  try{
    if (rowKey==='areaLine'){ const s=document.getElementById('tdAreaAlign'); if(s) s.value=align; }
    if (rowKey==='equLine'){ const s=document.getElementById('tdLineAlign'); if(s) s.value=align; }
    if (rowKey==='idLine'){ const s=document.getElementById('tdIdAlign'); if(s) s.value=align; }
  }catch(e){}
  tdSetRowAlign(rowKey, align);
  tdUiRefreshTagRowButtons();
}

function tdUiRefreshTagRowButtons(){
  try{
    ensureTagTemplateDefaults();
    const t = state.tagTemplate;
    const cfg = [
      {row:'areaLine', keySel:'tdLblAreaW', valSel:'tdValAreaW', align:'tdAreaAlign'},
      {row:'equLine',  keySel:'tdLblLineW', valSel:'tdValLineW', align:'tdLineAlign'},
      {row:'idLine',   keySel:'tdLblIdW',   valSel:'tdValIdW',   align:'tdIdAlign'},
    ];
    cfg.forEach(c=>{
      const el = t.elements[c.row];
      if (!el) return;
      // sync hidden selects
      const sk=document.getElementById(c.keySel); if(sk) sk.value=String(el.keyWeight||400);
      const sv=document.getElementById(c.valSel); if(sv) sv.value=String(el.valWeight||400);
      const sa=document.getElementById(c.align);  if(sa) sa.value=String(el.align||'left');

      // update button groups inside the same cell
      const rowNodes = Array.from(document.querySelectorAll('[onclick^="tdUiSetWeight(\''+c.row+'\'"]')).map(n=>n);
    });
    // active classes (weight)
    const setActive=(btns, predicate)=>{
      btns.forEach(b=>{ b.classList.toggle('active', predicate(b)); });
    };
    // weight buttons: parse weight from onclick
    ['areaLine','equLine','idLine'].forEach(r=>{
      const el = t.elements[r]; if(!el) return;
      const keyBtns = Array.from(document.querySelectorAll(`button[onclick="tdUiSetWeight('${r}','key',400)"],button[onclick="tdUiSetWeight('${r}','key',700)"],button[onclick="tdUiSetWeight('${r}','key',900)"]`));
      keyBtns.forEach(b=>{
        const w = b.textContent==='R'?400:(b.textContent==='B'?700:900);
        b.classList.toggle('active', (el.keyWeight||400)===w);
      });
      const valBtns = Array.from(document.querySelectorAll(`button[onclick="tdUiSetWeight('${r}','val',400)"],button[onclick="tdUiSetWeight('${r}','val',700)"],button[onclick="tdUiSetWeight('${r}','val',900)"]`));
      valBtns.forEach(b=>{
        const w = b.textContent==='R'?400:(b.textContent==='B'?700:900);
        b.classList.toggle('active', (el.valWeight||400)===w);
      });
      const alignBtns = Array.from(document.querySelectorAll(`button[onclick="tdUiSetAlign('${r}','left')"],button[onclick="tdUiSetAlign('${r}','center')"],button[onclick="tdUiSetAlign('${r}','right')"]`));
      alignBtns.forEach(b=>{
        const a = b.textContent==='L'?'left':(b.textContent==='C'?'center':'right');
        b.classList.toggle('active', (el.align||'left')===a);
      });
    });
  }catch(e){}
}

function tdSetRowAlign(rowKey, align){
  ensureTagTemplateDefaults();
  const el = state.tagTemplate.elements[rowKey];
  if (!el) return;
  const a = (align==='center'||align==='right'||align==='left') ? align : 'left';
  el.align = a;
  saveTagTemplate(state.tagTemplate);
  tdRenderStage();
}

function renderTDTagAudit(){
  const box = document.getElementById('tdTagAuditTable');
  if (!box) return;
  ensureTagThemeDefaults();
  ensureTagTemplateDefaults();
  const t = state.tagTemplate;
  const sizeMm = { w: Math.round((t.sizeCm?.w||7)*10), h: Math.round((t.sizeCm?.h||11)*10) };
  const rows = (state.tagTheme.audit||[]).map(a => {
    const dt = new Date(a.ts);
    const dstr = isNaN(dt.getTime()) ? (a.ts||'') : dt.toLocaleString();
    return `<tr>
      <td>${escapeHtml(dstr)}</td>
      <td>${escapeHtml(a.by||'')}</td>
      <td><span style="display:inline-block;width:14px;height:14px;border:1px solid #cbd5e1;border-radius:3px;background:${a.old||''};"></span> ${escapeHtml(a.old||'')}</td>
      <td><span style="display:inline-block;width:14px;height:14px;border:1px solid #cbd5e1;border-radius:3px;background:${a.new||''};"></span> ${escapeHtml(a.new||'')}</td>
    </tr>`;
  }).join('');
  box.innerHTML = `
    <div class="tiny" style="margin-bottom:8px;">Current size: <b>${sizeMm.w}×${sizeMm.h}mm</b> — Color: <b>${escapeHtml(state.tagTheme.color||'#0F6D8C')}</b></div>
    <table class="main-table">
      <thead><tr><th>Date</th><th>By</th><th>Old</th><th>New</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="4">No changes yet.</td></tr>'}</tbody>
    </table>
  `;
}

function tdResetTemplate(){
  resetTagTemplate();
  toast('Template reset');
  hydrateTagDesignerUI();
}

function tdSaveTemplate(){
  ensureTagTemplateDefaults();
  // Save snapshot as "last 3" templates: Template 1 newest, then 2, then 3
  // Default is locked and never overwritten.
  try {
    const obj = loadTagTemplates();
    obj.slots = obj.slots || {};
    // shift older
    obj.slots['tpl3'] = obj.slots['tpl2'] || getDefaultTagTemplate();
    obj.slots['tpl2'] = obj.slots['tpl1'] || getDefaultTagTemplate();
    obj.slots['tpl1'] = JSON.parse(JSON.stringify(state.tagTemplate || getDefaultTagTemplate()));
    obj.updatedAt = new Date().toISOString();
    saveTagTemplatesObj(obj);
    // make newest active
    setActiveTemplateId('tpl1');
    state.tagTemplate = obj.slots['tpl1'];
    // keep legacy key updated
    try { SBTS_UTILS.LS.set(TAG_TEMPLATE_KEY, JSON.stringify(state.tagTemplate)); } catch(e){}
    saveState();
    toast('Saved to Template 1 (latest).');
    hydrateTagDesignerUI();
  } catch (e) {
    // fallback: save into active slot
    saveTagTemplate(state.tagTemplate);
    toast('Template saved');
  }
}

function tdTestPrint(){
  ensureTagThemeDefaults();
  ensureTagTemplateDefaults();
  const cards=[{area:'SRU-3', line:'10\"-P-1001', blindId:'BL-001', qrValue:'SBTS', publicLink:'SBTS'}];
  openPrintWindow(cards);
}

function tdOnMouseDown(ev, key){
  // ignore handle drags
  if (ev.target && ev.target.classList && ev.target.classList.contains('td-handle')) return;
  ev.preventDefault();
  ev.stopPropagation();
  tdSelect(key);
  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;
  const rect = stage.getBoundingClientRect();
  const el = stage.querySelector(`.td-el[data-key="${key}"]`);
  if (!el) return;
  const elRect = el.getBoundingClientRect();
  TD_RUNTIME.dragging = {
    key,
    startX: ev.clientX,
    startY: ev.clientY,
    baseLeft: elRect.left - rect.left,
    baseTop: elRect.top - rect.top,
    stageW: rect.width,
    stageH: rect.height
  };
}

function tdOnResizeDown(ev, key){
  ev.preventDefault();
  ev.stopPropagation();
  tdSelect(key);
  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;
  const rect = stage.getBoundingClientRect();
  const el = stage.querySelector(`.td-el[data-key="${key}"]`);
  if (!el) return;
  const elRect = el.getBoundingClientRect();
  TD_RUNTIME.resizing = {
    key,
    startX: ev.clientX,
    startY: ev.clientY,
    baseLeft: elRect.left - rect.left,
    baseTop: elRect.top - rect.top,
    baseW: elRect.width,
    baseH: elRect.height,
    stageW: rect.width,
    stageH: rect.height
  };
}

function tdOnMouseMove(ev){
  if (!TD_RUNTIME.dragging && !TD_RUNTIME.resizing) return;
  const stage = document.getElementById('tagDesignerStage');
  if (!stage) return;

  ensureTagTemplateDefaults();
  const t = state.tagTemplate;

  if (TD_RUNTIME.dragging){
    const d = TD_RUNTIME.dragging;
    const dx = ev.clientX - d.startX;
    const dy = ev.clientY - d.startY;
    let left = d.baseLeft + dx;
    let top = d.baseTop + dy;
    left = tdSnapPx(left);
    top = tdSnapPx(top);

    // bounds
    left = Math.max(0, Math.min(left, d.stageW - 10));
    top = Math.max(0, Math.min(top, d.stageH - 10));

    // update template in cm
    const key = d.key;
    if (key === 'hole'){
      const e = t.elements.hole;
      const dcm = tdPx(e.d); // diameter in px
      e.x = tdCm(left + dcm/2);
      e.y = tdCm(top + dcm/2);
    } else {
      const e = t.elements[key];
      if (e){
        e.x = tdCm(left);
        e.y = tdCm(top);
      }
    }
    saveTagTemplate(t);
    tdRenderStage();
    return;
  }

  if (TD_RUNTIME.resizing){
    const r = TD_RUNTIME.resizing;
    const dx = ev.clientX - r.startX;
    const dy = ev.clientY - r.startY;
    let w = Math.max(12, r.baseW + dx);
    let h = Math.max(12, r.baseH + dy);
    w = tdSnapPx(w); h = tdSnapPx(h);

    const key = r.key;
    if (key === 'hole'){
      const s = Math.max(w, h);
      t.elements.hole.d = tdCm(s);
    } else {
      const e = t.elements[key];
      if (e){
        e.w = tdCm(w);
        e.h = tdCm(h);
      }
    }
    saveTagTemplate(t);
    tdRenderStage();
    return;
  }
}

function tdOnMouseUp(){
  TD_RUNTIME.dragging = null;
  TD_RUNTIME.resizing = null;
}
function buildTagCardHTML(data, color, monthYear){
  const bg = color || '#0F6D8C';
  const qrValue = data.qrValue || data.publicLink || '';
  const qrImg = qrValue ? `<img alt="QR" src="${buildQrImageUrl(qrValue)}" />` : ``;
  const logoUrl = (state?.branding?.companyLogo) ? state.branding.companyLogo : "";

  // Match physical tag layout: center hole, logo top-right, date bottom-left (MM/YYYY)
  return `
    <div class="sbts-tag-preview">
      <div class="sbts-tag-card" style="background:${bg};">
        <div class="sbts-tag-hole-center"></div>
        <div class="sbts-tag-logo-corner" style="${logoUrl ? `background-image:url('${logoUrl}');` : ''}"></div>

        <div class="sbts-tag-title">Smart Blind Tag</div>

        <div class="sbts-tag-qr">${qrImg}</div>

        <div class="sbts-tag-info">
          <div>Area: ${escapeHtml(data.area||'-')}</div>
          <div>Line: ${escapeHtml(data.line||data.equipment||'-')}</div>
          <div>ID: ${escapeHtml(data.blindId||'-')}</div>
        </div>

        <div class="sbts-tag-date-corner">${escapeHtml(monthYear||'')}</div>
      </div>
    </div>
  `;
}

function printBlindTagFromDetails(){
  const blindId = state.currentBlindId;
  const blind = state.blinds.find(b => b.id === blindId);
  if (!blind) return toast('Blind not found');
  const area = state.areas.find(a => a.id === blind.areaId);
  ensureTagThemeDefaults();
  const publicLink = buildPublicBlindUrl(blind.id);
  const cardData = {
    area: area?.name || '-',
    line: blind.line || '-',
    blindId: blind.name || '-',
    publicLink,
    qrValue: publicLink
  };
  openPrintWindow([cardData]);
}

function slipEnsureSelection(){
  if (!state.slip) state.slip = { areaId:'', projectId:'', selectedIds: [] };
  if (!Array.isArray(state.slip.selectedIds)) state.slip.selectedIds = [];
}

function slipToggleSelect(blindId, checked){
  slipEnsureSelection();
  const set = new Set(state.slip.selectedIds);
  if (checked) set.add(blindId); else set.delete(blindId);
  state.slip.selectedIds = Array.from(set);
  const c = document.getElementById('slipSelectedCount');
  if (c) c.textContent = String(state.slip.selectedIds.length);
}

function slipToggleSelectAll(checked){
  slipEnsureSelection();
  const list = state.blinds.filter(b => b.type === 'slip' && (!state.slip.projectId || b.projectId === state.slip.projectId));
  if (checked){
    state.slip.selectedIds = list.map(b => b.id);
  } else {
    state.slip.selectedIds = [];
  }
  renderSlipBlindPage();
}

function slipClearTagSelection(){
  slipEnsureSelection();
  state.slip.selectedIds = [];
  renderSlipBlindPage();
}

function slipPrintAllTags(){
  if (!state.slip?.projectId) return toast('Select a project first');
  ensureTagThemeDefaults();
  const list = state.blinds.filter(b => b.type === 'slip' && b.projectId === state.slip.projectId);
  const cards = list.map(b => {
    const area = state.areas.find(a => a.id === b.areaId);
    const publicLink = buildPublicBlindUrl(b.id);
    return { area: area?.name||'-', line: b.line||'-', blindId: b.name||'-', publicLink, qrValue: publicLink };
  });
  if (cards.length===0) return toast('No blinds');
  openPrintWindow(cards);
}

function slipPrintSelectedTags(){
  if (!state.slip?.projectId) return toast('Select a project first');
  slipEnsureSelection();
  const ids = state.slip.selectedIds || [];
  if (ids.length===0) return toast('No selection');
  const cards = ids.map(id => state.blinds.find(b=>b.id===id)).filter(Boolean).map(b=>{
    const area = state.areas.find(a => a.id === b.areaId);
    const publicLink = buildPublicBlindUrl(b.id);
    return { area: area?.name||'-', line: b.line||'-', blindId: b.name||'-', publicLink, qrValue: publicLink };
  });
  openPrintWindow(cards);
}

function openPrintWindow(cards){
  // Patch39.1: Label-printer print engine (one label per page, exact size)
  ensureTagThemeDefaults();
  ensureTagTemplateDefaults();
  const color = state.tagTheme.color || '#0F6D8C';
  const monthYear = getTagMonthYear();

  const t = state.tagTemplate || getDefaultTagTemplate();
  const wMm = Math.round((t.sizeCm?.w || 7) * 10);
  const hMm = Math.round((t.sizeCm?.h || 11) * 10);
  const cornerMm = 4; // Rounded corners for safe handling

  const w = window.open('', '_blank');
  if (!w) return alert('Popup blocked. Please allow popups for printing.');

  // Build pages: each tag is its own page (template size in mm)
  const pages = (cards || []).map(d => {
    return `<div class="page">${buildTagCardHTMLFromTemplate(d, color, monthYear)}</div>`;
  }).join('');

  w.document.open();
  w.document.write(`<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>SBTS Tags</title>
  <style>
  :root{ --w:${wMm}mm; --h:${hMm}mm; --r:${cornerMm}mm; }
  @page { size: var(--w) var(--h); margin: 0; }
  html,body{ width:var(--w); height:var(--h); margin:0; padding:0; }
  body{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }

  .page{ width:var(--w); height:var(--h); page-break-after: always; break-after: page; margin:0; padding:0; }
  .sbts-tag-preview{ width:var(--w); height:var(--h); margin:0; }
  .sbts-tag-card{ position:relative; box-sizing:border-box; overflow:hidden; color:#fff; width:var(--w) !important; height:var(--h) !important; border-radius:var(--r) !important; }

  .sbts-tag-hole-center{position:absolute;border-radius:999px;background:rgba(255,255,255,.96);box-shadow: inset 0 0 0 2.8mm rgba(0,0,0,.14);transform:translate(-50%,-50%);} 
  .sbts-tag-logo-corner{position:absolute;border-radius:3mm;background:rgba(255,255,255,.92);background-size:cover;background-position:center;}
  .sbts-tag-title{position:absolute;display:flex;align-items:center;justify-content:center;}

  .sbts-tag-qr{position:absolute;background:#fff;box-sizing:border-box;display:flex;align-items:center;justify-content:center;}
  .sbts-tag-qr-inner{width:100%;height:100%;display:flex;align-items:center;justify-content:center;}
  .sbts-tag-qr-inner img{max-width:100%;max-height:100%;object-fit:contain;}

  .sbts-tag-row{position:absolute;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .sbts-tag-row-key{opacity:.9;margin-right:2mm;}
  .sbts-tag-date-corner{position:absolute;opacity:.85;}

  @media print{ body{margin:0;padding:0;} }
</style>
</head>
<body>
${pages}
<script>window.onload=()=>{setTimeout(()=>{window.print();}, 200);};</script>
</body>
</html>`);
  w.document.close();
}


/* ==========================
   CERTIFICATE SETTINGS PAGE
========================== */
function hydrateCertificateSettingsUI() {
  const cfg = state.certificate;

  // templates select
  const sel = document.getElementById("certTemplateSelect");
  if (sel) {
    sel.innerHTML = "";
    cfg.templates.forEach(t => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = t.name;
      sel.appendChild(opt);
    });
    sel.value = cfg.activeTemplate || "default";
  }

  const setVal = (id, value) => { const el = document.getElementById(id); if (el) el.value = value ?? ""; };
  const setCheck = (id, value) => { const el = document.getElementById(id); if (el) el.checked = !!value; };

  setVal("certTitleInput", cfg.title);
  const bg = document.getElementById("certHeaderBg");
  if (bg) bg.value = cfg.headerBg || "#ffffff";

  const style = document.getElementById("certStatusStyle");
  if (style) style.value = cfg.statusStyle || "big";

  setCheck("certShowWorkflow", cfg.showWorkflow);
  setCheck("certShowApprovals", cfg.showApprovals);
  setVal("certFooterText", cfg.footerText);
}

function updateCertConfig(key, value) {
  state.certificate[key] = value;
  saveState();
}

function setCertTemplate(templateId) {
  state.certificate.activeTemplate = templateId;
  saveState();
}

function previewCertificateTemplate() {
  
  // Quick preview: if no current blind, create dummy page
  alert("Preview will open when you open any Blind -> Certificate. (This button keeps your current data safe)");
}

function setShowTrainingPage(val) {
  if (state.currentUser?.role !== "admin") { alert("Admin only."); return; }
  state.ui = state.ui || {};
  state.ui.showTrainingPage = !!val;
  saveState();
  applyNavVisibility();
}


/* ==========================
   QR + PUBLIC (Visitor) VIEW
========================== */
function buildBaseHref() {
  // Works for file:// and http(s)://
  return window.location.href.split("#")[0];
}
function buildPublicBlindUrl(blindId) {
  return buildBaseHref() + "#public/blind/" + encodeURIComponent(blindId);
}
function buildQrImageUrl(data) {
  // Online QR renderer (works when internet available). The system still works offline via the URL.
  return "https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=" + encodeURIComponent(data);
}
function copyToClipboard(text) {
  if (!text) return;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => toast("Copied")).catch(() => fallbackCopy(text));
  } else fallbackCopy(text);
}
function fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  document.body.appendChild(ta);
  ta.select();
  document.execCommand("copy");
  document.body.removeChild(ta);
  toast("Copied");
}
function copyBlindQrLink() {
  const el = document.getElementById("blindQrLink");
  copyToClipboard(el ? el.textContent : "");
}
function openBlindPublicView() {
  if (!state.currentBlindId) return;
  window.location.hash = "public/blind/" + encodeURIComponent(state.currentBlindId);
}
function copyPublicLink() {
  const el = document.getElementById("pub_link");
  copyToClipboard(el ? el.textContent : "");
}
function publicLogin() {
  // Save target blind to open after login
  const id = getPublicBlindIdFromHash();
  runtime.afterLoginBlindId = id;
  window.location.hash = ""; // back to main
  showAuthScreen();
  openLogin();
}
function goHomeFromPublic() {
  window.location.hash = "";
  if (state.currentUser) showAppScreen();
  else showAuthScreen();
}
function getPublicBlindIdFromHash() {
  const h = (window.location.hash || "").replace("#", "");
  const m = h.match(/^public\/blind\/(.+)$/);
  return m ? decodeURIComponent(m[1]) : null;
}
function renderPublicBlind(blindId) {
  const blind = state.blinds.find((b) => b.id === blindId);
  const badge = document.getElementById("publicBadge");
  if (badge) badge.textContent = (state.branding?.programLeftShort || "SB").slice(0,2).toUpperCase();

  const siteEl = document.getElementById("publicSiteName");
  if (siteEl) siteEl.textContent = (state.branding?.siteTitle || "Visitor view");

  const titleEl = document.getElementById("publicProgramTitle");
  if (titleEl) titleEl.textContent = state.branding?.programLeft || "Smart Blind Tag System";

  if (!blind) {
    document.getElementById("pub_blindName").textContent = "Not found";
    document.getElementById("pub_area").textContent = "-";
    document.getElementById("pub_project").textContent = "-";
    document.getElementById("pub_line").textContent = "-";
    document.getElementById("pub_type").textContent = "-";
    document.getElementById("pub_size").textContent = "-";
    document.getElementById("pub_rate").textContent = "-";
    document.getElementById("pub_phase").textContent = "-";
    document.getElementById("pub_approvals").textContent = "Blind not found in this browser storage.";
    document.getElementById("pub_history").innerHTML = "";
    return;
  }

  const area = state.areas.find((a) => a.id === blind.areaId);
  const project = state.projects.find((p) => p.id === blind.projectId);

  document.getElementById("pub_blindName").textContent = blind.name || blind.id;
  document.getElementById("pub_area").textContent = area ? area.name : "-";
  document.getElementById("pub_project").textContent = project ? project.name : "-";
  document.getElementById("pub_line").textContent = blind.line || "-";
  document.getElementById("pub_type").textContent = blind.type || "-";
  document.getElementById("pub_size").textContent = blind.size || "-";
  document.getElementById("pub_rate").textContent = blind.rate || "-";
  document.getElementById("pub_phase").textContent = phaseLabel(blind.phase || "broken");

  const link = buildPublicBlindUrl(blind.id);
  document.getElementById("pub_link").textContent = link;
  const img = document.getElementById("pub_qrImg");
  if (img) img.src = buildQrImageUrl(link);

  // approvals summary
  try {
    const fa = blind.finalApprovals || {};
    const lines = FINAL_APPROVALS.map((a) => {
      const ok = fa[a.key]?.status === "approved";
      const by = fa[a.key]?.name || "-";
      const lbl = faLookupName((a.key||"").toUpperCase(), a.label);
      return `${lbl}: ${ok ? "✅" : "⏳"} (By: ${by})`;
    }).join("\n");
    document.getElementById("pub_approvals").textContent = lines || "-";
  } catch (e) {
    document.getElementById("pub_approvals").textContent = "-";
  }

  // history table
  const body = document.getElementById("pub_history");
  body.innerHTML = "";
  const hist = (blind.history || []).slice().sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  hist.forEach((h) => {
    const tr = document.createElement("tr");
    const d = document.createElement("td");
    d.textContent = h.date ? new Date(h.date).toLocaleString() : "-";
    const f = document.createElement("td"); f.textContent = phaseLabel(h.fromPhase || "-");
    const t = document.createElement("td"); t.textContent = phaseLabel(h.toPhase || "-");
    const by = document.createElement("td"); by.textContent = h.byName || h.by || "-";
    tr.appendChild(d); tr.appendChild(f); tr.appendChild(t); tr.appendChild(by);
    body.appendChild(tr);
  });
}
function showPublicScreen(blindId) {
  document.getElementById("authContainer").classList.add("hidden");
  document.getElementById("appContainer").classList.add("hidden");
  document.getElementById("publicContainer").classList.remove("hidden");
  renderPublicBlind(blindId);
}
function handleRouting() {
  const id = getPublicBlindIdFromHash();
  if (id) {
    showPublicScreen(id);
    return;
  }
  // default routing
  document.getElementById("publicContainer").classList.add("hidden");
  if (state.currentUser) showAppScreen();
  else showAuthScreen();
}


/* ==========================
   INIT
========================== */

function initWorkflowSandboxToggle() {
  const el = document.getElementById("wfSandboxEnforcementToggle");
  if (!el) return;
  // default true
  if (state.ui?.workflowSandboxEnforcement === undefined) state.ui.workflowSandboxEnforcement = true;
  el.checked = !!state.ui.workflowSandboxEnforcement;
  el.onchange = () => {
    state.ui.workflowSandboxEnforcement = !!el.checked;
    saveState();
    toast("Workflow sandbox enforcement " + (el.checked ? "enabled" : "disabled") + ".");
  };
}


function sbtsShowFatalError(err) {
  try {
    const msg = (err && (err.message || err.reason && err.reason.message)) ? (err.message || err.reason.message) : String(err || "Unknown error");
    console.error("SBTS fatal error:", err);

    // Ensure something is visible for the user (avoid a blank index page)
    const app = document.getElementById("appContainer");
    const auth = document.getElementById("publicContainer");
    if (app) app.classList.remove("hidden");
    if (auth) auth.classList.add("hidden");

    let box = document.getElementById("sbtsFatalErrorBox");
    if (!box) {
      box = document.createElement("div");
      box.id = "sbtsFatalErrorBox";
      box.style.position = "fixed";
      box.style.left = "16px";
      box.style.right = "16px";
      box.style.bottom = "16px";
      box.style.zIndex = "99999";
      box.style.padding = "14px 14px";
      box.style.borderRadius = "12px";
      box.style.background = "#1f2937";
      box.style.color = "white";
      box.style.boxShadow = "0 10px 30px rgba(0,0,0,.35)";
      box.style.fontSize = "13px";
      box.innerHTML = `
        <div style="display:flex;align-items:flex-start;gap:12px;">
          <div style="font-size:18px;line-height:18px;">⚠️</div>
          <div style="flex:1;">
            <div style="font-weight:700;margin-bottom:6px;">SBTS Error</div>
            <div style="opacity:.9;margin-bottom:8px;" id="sbtsFatalErrorMsg"></div>
            <div style="opacity:.8;">Tip: Refresh the page. If it repeats, send me the error message shown here.</div>
          </div>
          <button id="sbtsFatalErrorClose" style="background:rgba(255,255,255,.15);color:white;border:none;border-radius:10px;padding:6px 10px;cursor:pointer;">Close</button>
        </div>
      `;
      document.body.appendChild(box);
      const btn = document.getElementById("sbtsFatalErrorClose");
      if (btn) btn.onclick = () => box.remove();
    }
    const msgEl = document.getElementById("sbtsFatalErrorMsg");
    if (msgEl) msgEl.textContent = msg;
  } catch (_) {
    // last resort: do nothing
  }
}

// Catch unexpected errors so the UI doesn't become a blank page
window.addEventListener("error", (e) => sbtsShowFatalError(e && e.error ? e.error : e));
window.addEventListener("unhandledrejection", (e) => sbtsShowFatalError(e));

/* ==========================
   CLEANUP5: BOOT REFACTOR (NO BEHAVIOR CHANGE)
   - Split sbtsBoot() into small init steps for easier maintenance
   - Keep call order exactly the same
========================== */

function sbtsInitStateAndConfig() {
  loadState();
  // Weekly local backup (silent) – Admin toggle controlled
  try { sbtsRunWeeklyAutoBackup(); } catch (e) {}
  state.rolesCatalog = loadRolesCatalog();
  populateRoleSelects();
  ensureReportsCardsConfig();
}

function sbtsInitThemeAndFont() {
  // apply theme + font
  const themePicker = document.getElementById("themePicker");
  if (themePicker) themePicker.value = state.themeColor || "#174c7e";
  applyFontSize(state.fontSize || 14);
}

function sbtsInitSidebarUserNotificationsLink() {
  // Open Notifications Inbox when clicking user profile box in sidebar
  const su = document.querySelector(".sidebar-user");

  const subox = document.getElementById("sidebarUserBox");
  if (subox) {
    subox.style.cursor = "pointer";
    subox.addEventListener("click", () => {
      try { openPage("notificationsPage"); } catch (_) {}
    });
  }

  if (su) {
    su.style.cursor = "pointer";
    su.addEventListener("click", () => {
      try { openPage("notificationsPage"); } catch (_) {}
    });
  }
}

function sbtsInitRoutingHandlers() {
  handleRouting();
  window.addEventListener("hashchange", handleRouting);
}

function sbtsInitBreadcrumbs() {
  // Init breadcrumbs / back stack
  setTimeout(() => {
    try {
      const s = navLoadStack();
      if (!s || s.length === 0) {
        const p = state.ui?.lastPage || SBTS_UTILS.LS.get("sbts_last_page") || "dashboardPage";
        navPush(navEntryForCurrent(p));
      } else {
        updateNavUI();
      }
    } catch (e) {}
  }, 0);
}

function sbtsInitResponsiveHandlers() {
  // Responsive re-render (switch table/cards on resize)
  let __sbtsResizeT;
  window.addEventListener("resize", () => {
    clearTimeout(__sbtsResizeT);
    __sbtsResizeT = setTimeout(() => {
      try { if (document.getElementById("projectBlindsTableBody") || document.getElementById("projectBlindsCards")) renderProjectBlindsTable(); } catch (e) {}
      // Close mobile sidebar if moved to desktop
      try { if (!isMobileView()) closeMobileSidebar(true); } catch (e) {}
    }, 120);
  });
}



// ===========================
// Patch 47.0 – Notification/Inbox Foundation Tools
// ===========================
function sbtsSetupGlobalErrorHandlers() {
  if (window.__sbtsErrHandlersInstalled) return;
  window.__sbtsErrHandlersInstalled = true;

  function capture(context, err) {
    try {
      const msg = (err && (err.message || err.reason?.message)) || String(err || "Unknown error");
      const stack = (err && (err.stack || err.reason?.stack)) || "";
      const item = {
        id: "E-" + Date.now(),
        at: new Date().toISOString(),
        context,
        message: msg,
        stack
      };
      state.debug = state.debug || { errors: [] };
      state.debug.errors = Array.isArray(state.debug.errors) ? state.debug.errors : [];
      state.debug.errors.unshift(item);
      state.debug.errors = state.debug.errors.slice(0, 50);
      saveState();
      // Show a helpful toast (no more "Script error" mystery)
      showToast("SBTS Error: " + msg, "error", 8000);
      console.error("[SBTS Error]", context, err);
    } catch (e) {
      console.error("[SBTS Error Handler Failed]", e);
    }
  }

  window.addEventListener("error", (e) => capture("window.error", e.error || e.message));
  window.addEventListener("unhandledrejection", (e) => capture("unhandledrejection", e.reason || e));
}

function sbtsResetNotifSystem() {
  try {
    // Reset only notification/inbox related local data
    ensureNotificationsState();
    state.notifications.byUser = {};
    state.notifications.settings = state.notifications.settings || {};
    state.notifications.lastSeenAt = {};

    // Legacy / request engine leftovers (safe reset)
    state.inboxItems = [];
    state.requests = [];
    state.messages = [];

    saveState();
    updateNotificationsBadge();
    showToast("Notification system reset (local).", "success");
  } catch (e) {
    showToast("Reset failed: " + (e?.message || e), "error");
  }
}

function sbtsSeedNotifDemo() {
  try {
    const u = state.currentUser?.id || "system-admin";
    const proj = state.currentProjectId || (state.projects?.[0]?.id) || null;

    addNotification({
      recipients: [u],
      kind: "action",
      title: "User registration pending approval",
      message: 'New user "fahad" requested access (needs admin approval).',
      projectId: proj,
      meta: { requestType: "user_registration", username: "fahad" }
    });

    addNotification({
      recipients: [u],
      kind: "update",
      title: "Visa transfer requested",
      message: 'Transfer SB-025 to user "Ahmed" (requires review).',
      projectId: proj,
      meta: { requestType: "visa_transfer", blindDisplay: "SB-025 | Area 2 | D-111 | 10\"" }
    });

    addNotification({
      recipients: [u],
      kind: "system",
      title: "Admin announcement",
      message: "Safety note: follow PTW / isolation procedure before any operation.",
      projectId: proj,
      meta: { requestType: "announcement" }
    });

    saveState();
    updateNotificationsBadge();
    showToast("Demo notifications created.", "success");
  } catch (e) {
    showToast("Seed failed: " + (e?.message || e), "error");
  }
}

function sbtsBindNotifInspectorButtons() {
  if (window.__sbtsNotifInspectorBound) return;
  window.__sbtsNotifInspectorBound = true;

  const sec = document.getElementById("notifInspectorSection");
  if (!sec) return;

  const btnOpen = document.getElementById("btnOpenNotifInspector");
  const btnClose = document.getElementById("btnCloseNotifInspector");
  const btnReset = document.getElementById("btnResetNotifSystem");
  const btnSeed = document.getElementById("btnSeedNotifDemo");
  const panel = document.getElementById("notifInspectorPanel");
  const body = document.getElementById("notifInspectorBody");

  function renderPanel() {
    ensureNotificationsState();
    const uid = state.currentUser?.id;
    const items = (uid && state.notifications.byUser?.[uid]) ? state.notifications.byUser[uid] : [];
    const unread = items.filter(x => !x.readAt && !x.doneAt && !x.archivedAt).length;
    const action = items.filter(x => (x.kind === "action") && !x.doneAt && !x.archivedAt).length;
    const updates = items.filter(x => (x.kind === "update" || x.kind === "system") && !x.archivedAt).length;

    const recent = items.slice(0, 8).map(n => {
      const t = escapeHtml(n.title || "Untitled");
      const p = escapeHtml(n.projectName || "");
      const k = escapeHtml((n.kind || "").toUpperCase());
      return `<div style="padding:8px 0; border-bottom:1px solid rgba(0,0,0,.06);">
        <div style="font-weight:700;">${t} <span class="muted" style="font-weight:600;">${k}</span></div>
        <div class="muted" style="font-size:12px;">${p} • ${formatRelativeTime(n.createdAt || Date.now())}</div>
      </div>`;
    }).join("");

    const errs = (state.debug?.errors || []).slice(0, 5).map(e => {
      const msg = escapeHtml(e.message || "");
      const ctx = escapeHtml(e.context || "");
      return `<div style="padding:8px 0; border-bottom:1px solid rgba(0,0,0,.06);">
        <div style="font-weight:700;">${ctx}</div>
        <div class="muted" style="font-size:12px;">${msg}</div>
      </div>`;
    }).join("");

    body.innerHTML = `
      <div class="grid" style="gap:10px;">
        <div class="card" style="padding:12px;">
          <div style="font-weight:800; font-size:14px;">Counts</div>
          <div class="muted" style="margin-top:6px;">Unread: <b>${unread}</b> • Action: <b>${action}</b> • Updates: <b>${updates}</b></div>
        </div>
        <div class="card" style="padding:12px;">
          <div style="font-weight:800; font-size:14px;">Recent notifications</div>
          <div style="margin-top:6px;">${recent || '<div class="muted">No items.</div>'}</div>
        </div>
        <div class="card" style="padding:12px;">
          <div style="font-weight:800; font-size:14px;">Recent errors</div>
          <div style="margin-top:6px;">${errs || '<div class="muted">No errors logged.</div>'}</div>
        </div>
      </div>
    `;
  }

  btnOpen?.addEventListener("click", () => {
    panel.style.display = "";
    renderPanel();
  });

  btnClose?.addEventListener("click", () => {
    panel.style.display = "none";
  });

  btnReset?.addEventListener("click", () => {
    if (!confirm("Reset notifications/inbox (local)?")) return;
    sbtsResetNotifSystem();
    renderPanel();
  });

  btnSeed?.addEventListener("click", () => {
    // Seed practical demo scenarios (Project > Blind threads + action workflows)
    try{
      seedDemoInbox();
      // Button UX
      if (btnSeed){
        const count = parseInt(localStorage.getItem("sbts_seed_live_scenarios_v2_count")||"0",10)||0;
        btnSeed.textContent = (count>0) ? "Seed ✓" : "Seed";
        btnSeed.classList.add("btn-secondary");
      }
    }catch(e){
      console.error(e);
    }
    renderPanel();
  });
}


/* ==========================
   PATCH47.31 - DEMO MULTI-USER MODE (uses Roles & Specialties Manager data)
   - Builds demo users from Roles Catalog if missing
   - Adds "Acting as" switcher on Inbox
   - Persists selection via localStorage (sbts_demo_acting_user_v1)
========================== */

function sbtsGetDemoPinnedUserId(){
  try { return SBTS_UTILS.LS.get("sbts_demo_acting_user_v1"); } catch(e){ return null; }
}
function sbtsSetDemoPinnedUserId(uid){
  try { SBTS_UTILS.LS.set("sbts_demo_acting_user_v1", uid || ""); } catch(e){}
}
function sbtsEnsureDemoUsersFromRolesCatalog(){
  // Use Roles & Specialties Manager (roles catalog) as source of truth
  try{
    const roles = getActiveRoleOptions ? getActiveRoleOptions() : (state.rolesCatalog||[]).filter(r=>r && r.active!==false);
    if(!Array.isArray(roles) || roles.length===0) return;

    state.users = Array.isArray(state.users) ? state.users : [];
    const existingByRole = new Set(state.users.map(u => (u && u.role ? normalizeRoleId(u.role) : "")).filter(Boolean));

    roles.forEach(r=>{
      const roleId = normalizeRoleId(r.id || r.roleId || r.value || r.name || "");
      if(!roleId) return;

      // if no real user with this role exists, create a demo one
      if(existingByRole.has(roleId)) return;

      const uid = "demo-role-" + roleId;
      const label = (r.label || r.name || roleId);
      if(state.users.some(u=>u && u.id===uid)) return;

      state.users.push({
        id: uid,
        fullName: label + " (Demo)",
        username: "demo_" + roleId,
        password: "demo",
        role: roleId,
        phone: "",
        email: "",
        status: "active",
        jobTitle: label,
        profileImage: null,
        createdAt: new Date().toISOString(),
        __demo: true
      });

      // baseline permissions entry to avoid undefined access
      state.permissions = state.permissions || {};
      if(!state.permissions[uid]) state.permissions[uid] = {};
    });
    saveState();
  }catch(e){
    console.warn("[demoUsersFromCatalog] skipped", e);
  }
}

function sbtsIsDemoSwitcherAvailable(){
  // Admin-only demo mode. Hide for regular users.
  try{
    if(!sbtsIsAdminLike()) return false;
    const users = Array.isArray(state.users) ? state.users.filter(u=>u && u.status!=="inactive") : [];
    if(users.length < 2) return false;
    const roles = new Set(users.map(u => normalizeRoleId(u.role||"")).filter(Boolean));
    return roles.size >= 2 || users.length >= 2;
  }catch(e){ return false; }
}

function sbtsDemoGetCurrentActingUserId(){
  // If pinned, use it; else use currentUser id
  const pinned = sbtsGetDemoPinnedUserId();
  if(pinned && pinned !== "") return pinned;
  return (state.currentUser && state.currentUser.id) ? state.currentUser.id : null;
}

function sbtsDemoSwitchUser(userId){
  try{
    if(!userId) return;
    const u = (Array.isArray(state.users)?state.users:[]).find(x=>x && x.id===userId);
    if(!u) return;
    state.currentUser = u;
    saveState();
    try{ updateSidebarUserCard(); }catch(_){}
    try{ updateNotificationsBadge(); }catch(_){}
    try{ renderNotificationsInbox(); }catch(_){}
    try{ renderNotificationsDrawer(); }catch(_){}
    try{ renderNotificationsPage(); }catch(_){}
    showToast(`Acting as: ${u.fullName || u.username || u.role}`, "success");
  }catch(e){
    console.error("[sbtsDemoSwitchUser] failed", e);
    showToast("Demo switch failed", "error");
  }
}

function sbtsDemoTogglePinUser(){
  try{
    const sel = document.getElementById("demoUserSelect");
    if(!sel) return;
    const uid = sel.value;
    const cur = sbtsGetDemoPinnedUserId();
    if(cur && cur !== ""){
      sbtsSetDemoPinnedUserId("");
      showToast("Demo user unpinned.", "success");
    }else{
      sbtsSetDemoPinnedUserId(uid);
      showToast("Demo user pinned for this browser.", "success");
    }
    sbtsDemoRefreshSwitcherUI();
  }catch(e){ console.warn(e); }
}

function sbtsDemoRefreshSwitcherUI(){
  const wrap = document.getElementById("demoUserSwitch");
  const sel  = document.getElementById("demoUserSelect");
  const pin  = document.getElementById("demoPinUserBtn");
  if(!wrap || !sel) return;

  if(!sbtsIsDemoSwitcherAvailable()){
    wrap.style.display = "none";
    return;
  }
  wrap.style.display = "flex";

  // Populate options grouped by role from catalog labels (realistic)
  const users = (Array.isArray(state.users)?state.users:[])
    .filter(u=>u && u.status!=="inactive")
    .slice()
    .sort((a,b)=>{
      const ra = getRoleLabelById ? (getRoleLabelById(a.role)||a.role||"") : (a.role||"");
      const rb = getRoleLabelById ? (getRoleLabelById(b.role)||b.role||"") : (b.role||"");
      return (ra+""+ (a.fullName||a.username||"")).localeCompare(rb+""+(b.fullName||b.username||""));
    });

  const grouped = {};
  users.forEach(u=>{
    const roleId = normalizeRoleId(u.role||"") || "other";
    const label = (getRoleLabelById ? (getRoleLabelById(roleId)||roleId) : roleId);
    const g = label || roleId;
    grouped[g] = grouped[g] || [];
    grouped[g].push(u);
  });

  const acting = sbtsDemoGetCurrentActingUserId();
  sel.innerHTML = Object.keys(grouped).sort().map(g=>{
    const opts = grouped[g].map(u=>{
      const name = (u.fullName || u.username || u.id);
      const demoTag = u.__demo ? " • demo" : "";
      return `<option value="${u.id}">${name}${demoTag}</option>`;
    }).join("");
    return `<optgroup label="${escapeHtml(g)}">${opts}</optgroup>`;
  }).join("");

  if(acting && users.some(u=>u.id===acting)){
    sel.value = acting;
  }else if(users[0]){
    sel.value = users[0].id;
  }

  const pinned = sbtsGetDemoPinnedUserId();
  if(pin){
    const pinnedOn = pinned && pinned !== "";
    pin.textContent = pinnedOn ? "📌" : "📍";
    pin.title = pinnedOn ? "Unpin demo user" : "Pin demo user for this browser";
  }
}


function sbtsApplyAdminOnlyVisibility(){
  try{
    const isAdmin = sbtsIsAdminLike();
    // Hide any element marked as admin-only
    document.querySelectorAll('[data-admin-only="1"]').forEach(el=>{
      if(!isAdmin){
        el.style.display = "none";
      }else{
        // restore default for switcher; sbtsDemoRefreshSwitcherUI will control it
        if(el.id === "demoUserSwitch") el.style.display = "none";
        else el.style.display = "";
      }
    });
  }catch(e){ /* ignore */ }
}

function sbtsInitDemoUserSwitcher(){
  try{
    sbtsEnsureDemoUsersFromRolesCatalog();
    sbtsDemoRefreshSwitcherUI();

    // Apply pinned user on load (realistic role-based demo)
    const pinned = sbtsGetDemoPinnedUserId();
    if(pinned && pinned !== "" && state.currentUser && state.currentUser.id !== pinned){
      sbtsDemoSwitchUser(pinned);
    }
  }catch(e){
    console.warn("[sbtsInitDemoUserSwitcher] skipped", e);
  }
}


function sbtsBoot() {
  window.__SBTS_BOOTED__ = true;

  try{ initSmartLinkDelegation(); }catch(_){ }

  
  sbtsSetupGlobalErrorHandlers();
sbtsInitStateAndConfig();
  try{ sbtsEnsureRealUser(); }catch(_){ }
  try{ sbtsApplyAdminOnlyVisibility(); }catch(_){ }
  try{ if(sbtsIsAdminLike()) sbtsInitDemoUserSwitcher(); }catch(_){ }
  sbtsInitThemeAndFont();

  initWorkflowSandboxToggle();

  // Routing (supports public visitor view via QR)
  sbtsInitSidebarUserNotificationsLink();
  sbtsInitRoutingHandlers();
  sbtsInitBreadcrumbs();
  sbtsInitResponsiveHandlers();
  try{ sbtsInitInboxFiltersRailHover(); }catch(_){ }
  try{ sbtsInitInboxDetailsResizer(); }catch(_){ }
}

document.addEventListener("DOMContentLoaded", () => {
  try {
    sbtsBoot();
  } catch (e) {
    sbtsShowFatalError(e);
  }
});



/* ==========================
   WORKFLOW CONTROL PAGE (D1)
   Config only. Saved in localStorage key: sbts_workflow_config_v1
========================== */

const WF_STORAGE_KEY = "sbts_workflow_config_v1";
const WF_SANDBOX_PROJECT_NAME = "TEST-WORKFLOW";

// Default phase colors (admin can override per phase in Workflow Control)
const WF_DEFAULT_PHASE_COLORS = {
  broken: "#ef4444",
  assembly: "#f59e0b",
  tightTorque: "#eab308",
  finalTight: "#22c55e",
  inspectionReady: "#3b82f6",
};

// Palette for new/custom phases
const WF_COLOR_PALETTE = [
  "#06b6d4", // cyan
  "#a855f7", // purple
  "#f97316", // orange
  "#22c55e", // green
  "#3b82f6", // blue
  "#ef4444", // red
  "#eab308", // yellow
  "#14b8a6", // teal
  "#8b5cf6", // violet
  "#f43f5e", // rose
];

function getWorkflowConfig() {
  const raw = SBTS_UTILS.LS.get(WF_STORAGE_KEY);
  if (!raw) return wfDefaultConfig();
  try {
    const cfg = JSON.parse(raw);
    if (!cfg || !Array.isArray(cfg.phases)) return wfDefaultConfig();
    wfEnsurePhaseColors(cfg);
    return cfg;
  } catch {
    return wfDefaultConfig();
  }
}

function wfEnsurePhaseColors(cfg) {
  if (!cfg || !Array.isArray(cfg.phases)) return;
  // Assign missing colors deterministically (stable ordering)
  const phasesSorted = [...cfg.phases].sort((a, b) => (a.order || 0) - (b.order || 0));
  let paletteIdx = 0;

  const used = new Set(phasesSorted.map(p => (p && p.color) ? String(p.color).toLowerCase() : null).filter(Boolean));
  const nextPalette = () => {
    // find next unused palette color
    for (let i = 0; i < WF_COLOR_PALETTE.length; i++) {
      const c = WF_COLOR_PALETTE[(paletteIdx + i) % WF_COLOR_PALETTE.length].toLowerCase();
      if (!used.has(c)) {
        paletteIdx = (paletteIdx + i + 1) % WF_COLOR_PALETTE.length;
        used.add(c);
        return c;
      }
    }
    // fallback
    const c = WF_COLOR_PALETTE[paletteIdx % WF_COLOR_PALETTE.length].toLowerCase();
    paletteIdx = (paletteIdx + 1) % WF_COLOR_PALETTE.length;
    return c;
  };

  phasesSorted.forEach((p) => {
    if (!p) return;
    if (!p.color) {
      const def = WF_DEFAULT_PHASE_COLORS[p.id];
      p.color = def ? def : nextPalette();
    }
  });
}

function wfPhaseColor(phaseId) {
  const cfg = getWorkflowConfig();
  const p = (cfg.phases || []).find(x => x.id === phaseId);
  return p?.color || WF_DEFAULT_PHASE_COLORS[phaseId] || "#2563eb";
}

function wfPhaseIds({ includeInactive = false } = {}) {
  const cfg = getWorkflowConfig();
  const phases = (cfg.phases || []).slice().sort((a,b)=> (a.order||0)-(b.order||0));
  return phases.filter(p => includeInactive ? true : p.active !== false).map(p => p.id);
}

function getProjectPhaseEnabledMap(projectId) {
  const p = getProjectByIdSafe(projectId);
  const map = p?.phaseOwners?.phaseEnabled;
  if (map && typeof map === "object") return map;
  return {};
}
function isPhaseEnabledForProject(projectId, phaseId) {
  const map = getProjectPhaseEnabledMap(projectId);
  const key = normalizePhaseId(phaseId);
  return map[key] !== false;
}
function setPhaseEnabledForProject(projectId, phaseId, enabled) {
  const p = getProjectByIdSafe(projectId);
  if (!p) return;
  if (!p.phaseOwners) p.phaseOwners = defaultProjectPhaseOwners();
  if (!p.phaseOwners.phaseEnabled) p.phaseOwners.phaseEnabled = {};
  p.phaseOwners.phaseEnabled[normalizePhaseId(phaseId)] = !!enabled;
  saveState();
}
function wfProjectPhaseIds(projectId, { includeInactive = false } = {}) {
  const ids = wfPhaseIds({ includeInactive });
  return ids.filter(ph => isPhaseEnabledForProject(projectId, ph));
}


function wfPhaseIndex(phaseId) {
  const ids = wfPhaseIds({ includeInactive: true });
  return ids.indexOf(phaseId);
}

function wfDefaultConfig() {
  const roles = Object.keys(ROLE_LABELS).map((id) => ({ id, label: ROLE_LABELS[id], active: true }));
  const phases = PHASES.map((id, idx) => ({
    id,
    label: phaseLabels[id] || id,
    color: WF_DEFAULT_PHASE_COLORS[id] || "",
    active: true,
    order: idx + 1,
    canUpdate: [],
    approval: { enabled: false, roles: [], mode: "parallel" },
    extra: { enabled: false, items: [] } // per-phase extra approvals
  }));

  // Global / conditional final approvals (not phases)
  const rules = {
    // B: Conditional final confirmation for Slip Blind only.
    // This is NOT a phase. It is a final approval requirement.
    slipBlindMetalForeman: {
      enabled: true,
      role: "metal_foreman",
      name: "Metal Foreman – Slip Blind Confirmation",
      type: "confirmation",
      options: ["Reinstalled", "Demolished"]
    }
  };

  // Return a complete default config object.
  const cfg = { version: 1, roles, phases, ui: { layout: "horizontal" }, rules };
  wfEnsurePhaseColors(cfg);
  return cfg;
}

// ==========================
// Global navigation stack (Smart Back + Breadcrumbs)
// - Stored in sessionStorage (per-tab) so it survives refresh but stays per session.
// ==========================
const NAV_STACK_KEY = "sbts_nav_stack_v1";
let NAV_SUPPRESS_PUSH = false;

function navLoadStack() {
  try {
    const raw = sessionStorage.getItem(NAV_STACK_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (_) { return []; }
}

function navSaveStack(stack) {
  try { sessionStorage.setItem(NAV_STACK_KEY, JSON.stringify(stack || [])); } catch (_) {}
}

function navCurrent() {
  const stack = navLoadStack();
  return stack.length ? stack[stack.length - 1] : null;
}

function navEntryForCurrent(pageId) {
  const entry = { pageId };
  // Attach context when relevant
  if (pageId === "projectDetailsPage") {
    if (state.currentProjectId) entry.projectId = state.currentProjectId;
    entry.subTab = state.ui?.projectSubTab || "blinds";
  }
  if (pageId === "blindDetailsPage") {
    if (state.currentBlindId) entry.blindId = state.currentBlindId;
    // Try to keep project context too (helps breadcrumbs)
    const b = (state.currentBlindId && state.blinds) ? state.blinds.find(x => x.id === state.currentBlindId) : null;
    const pid = b?.projectId || state.currentProjectId;
    if (pid) entry.projectId = pid;
  }
  return entry;
}

function navPush(entry) {
  if (!entry || !entry.pageId) return;
  const stack = navLoadStack();
  const last = stack[stack.length - 1];
  const same = last && last.pageId === entry.pageId && last.projectId === entry.projectId && last.blindId === entry.blindId && (last.subTab || "") === (entry.subTab || "");
  if (!same) {
    stack.push({ ...entry, ts: Date.now() });
    // keep last 50 to avoid unlimited growth
    while (stack.length > 50) stack.shift();
    navSaveStack(stack);
  }
  updateNavUI();
}

function navReplace(entry) {
  if (!entry || !entry.pageId) return;
  const stack = navLoadStack();
  if (stack.length === 0) {
    stack.push({ ...entry, ts: Date.now() });
  } else {
    stack[stack.length - 1] = { ...stack[stack.length - 1], ...entry, ts: Date.now() };
  }
  navSaveStack(stack);
  updateNavUI();
}

function navBack(evt) {
  if (evt) evt.stopPropagation();
  const stack = navLoadStack();
  // Pop current
  if (stack.length > 0) stack.pop();
  navSaveStack(stack);
  const prev = stack.length ? stack[stack.length - 1] : null;
  if (!prev) {
    // fallback
    openPage("dashboardPage");
    navPush({ pageId: "dashboardPage" });
    return;
  }
  navGo(prev, { replace: true });
}

function navGo(entry, opts = {}) {
  if (!entry || !entry.pageId) return;

  NAV_SUPPRESS_PUSH = true;

  // Restore context first
  if (entry.projectId) state.currentProjectId = entry.projectId;
  if (entry.blindId) state.currentBlindId = entry.blindId;

  // Navigate
  if (entry.pageId === "projectDetailsPage" && entry.projectId) {
    openProjectDetails(entry.projectId);
    if (entry.subTab) {
      try { openProjectSubTab(entry.subTab); } catch (_) {}
    }
  } else if (entry.pageId === "blindDetailsPage" && entry.blindId) {
    openBlindDetails(entry.blindId);
  } else {
    openPage(entry.pageId);
  }

  NAV_SUPPRESS_PUSH = false;

  if (opts.replace) navReplace(entry);
  else navPush(entry);
}

function getPageTitle(pageId) {
  const map = {
    dashboardPage: "Dashboard",
    areasPage: "Areas",
    projectsPage: "Projects",
    projectDetailsPage: "Project",
    blindDetailsPage: "Blind",
    notificationsPage: "Notifications",
    slipBlindPage: "Slip Blind",
    reportsPage: "Reports",
    reportsCardsPage: "Reports Cards",
    workflowControlPage: "Workflow Control",
    permissionsPage: "Users",
    certificateSettingsPage: "Certificate Settings",
    trainingPage: "Training",
    settingsPage: "Settings",
  };
  return map[pageId] || "";
}

function buildCrumbsForCurrentView() {
  // Use active page if possible
  const visible = Array.from(document.querySelectorAll('.page')).find(p => !p.classList.contains('hidden'));
  const activePage = visible?.id || state.ui?.lastPage || "dashboardPage";
  const crumbs = [];

  const pid = state.currentProjectId;
  const projectName = pid ? (state.projects.find(p => p.id === pid)?.name || "Project") : null;
  const bid = state.currentBlindId;
  const blindName = bid ? (state.blinds.find(b => b.id === bid)?.name || "Blind") : null;
  const subTab = state.ui?.projectSubTab || "blinds";

  if (activePage === "projectDetailsPage" && pid) {
    crumbs.push({ label: "Projects", go: () => navGo({ pageId: "projectsPage" }) });
    crumbs.push({ label: projectName, go: () => navGo({ pageId: "projectDetailsPage", projectId: pid, subTab }) });
    crumbs.push({ label: subTab === "settings" ? "Project Settings" : "Blinds" });
  } else if (activePage === "blindDetailsPage" && pid && bid) {
    crumbs.push({ label: "Projects", go: () => navGo({ pageId: "projectsPage" }) });
    crumbs.push({ label: projectName, go: () => navGo({ pageId: "projectDetailsPage", projectId: pid, subTab: "blinds" }) });
    crumbs.push({ label: "Blinds", go: () => navGo({ pageId: "projectDetailsPage", projectId: pid, subTab: "blinds" }) });
    crumbs.push({ label: blindName });
  } else if (activePage === "projectsPage") {
    crumbs.push({ label: "Projects" });
  } else if (activePage === "areasPage") {
    crumbs.push({ label: "Areas" });
  } else if (activePage === "notificationsPage") {
    crumbs.push({ label: "Notifications" });
  } else {
    const title = getPageTitle(activePage) || "Dashboard";
    crumbs.push({ label: title });
  }

  return crumbs;
}

function updateNavUI() {
  // Determine active page
  const visible = Array.from(document.querySelectorAll('.page')).find(p => !p.classList.contains('hidden'));
  const activePage = visible?.id || state.ui?.lastPage || "dashboardPage";

  // Root pages (opened directly from sidebar) should not show Back/Breadcrumbs.
  // On mobile we still keep a minimal bar to host the menu button.
  const rootPages = new Set([
    "dashboardPage",
    "areasPage",
    "projectsPage",
    "slipBlindPage",
    "reportsPage",
    "reportsCardsPage",
    "workflowControlPage",
    "permissionsPage",
    "certificateSettingsPage",
    "trainingPage",
    "settingsPage",
    "notificationsPage",
  ]);

  const topbar = document.getElementById("pageTopbar");
  if (topbar) {
    if (rootPages.has(activePage)) {
      topbar.classList.add("topbar-root");
      topbar.classList.remove("hidden");
    } else {
      topbar.classList.remove("topbar-root");
      topbar.classList.remove("hidden");
    }
  }

  // Back button enable/disable
  const stack = navLoadStack();
  const backBtn = document.getElementById("pageBackBtn") || document.getElementById("navBackBtn");
  if (backBtn) {
    // For root pages, hide back entirely via CSS (topbar-root). Still keep logic safe.
    if (stack.length <= 1) backBtn.classList.add("disabled");
    else backBtn.classList.remove("disabled");
  }

  // Breadcrumbs
  const trail = document.getElementById("navTrail");
  if (!trail) return;

  const crumbs = buildCrumbsForCurrentView();
  if (!crumbs || crumbs.length === 0) {
    trail.classList.add("hidden");
    trail.innerHTML = "";
    return;
  }
  // IMPORTANT: Do not serialize functions into onclick strings.
  // `c.go` closures rely on runtime variables (projectId/blindId), and `toString()` loses that context.
  // Instead, render links with data attributes and dispatch clicks to the real handler functions.
  window.__sbtsCrumbHandlers = crumbs.map(c => (typeof c.go === "function" ? c.go : null));

  trail.classList.remove("hidden");
  trail.innerHTML = `<div class="crumbs">` + crumbs.map((c, idx) => {
    const sep = idx === 0 ? "" : `<span class="crumb-sep">›</span>`;
    const isLink = typeof c.go === "function";
    const inner = isLink
      ? `<a href="#" data-crumb-idx="${idx}">${escapeHtml(c.label)}</a>`
      : `<span>${escapeHtml(c.label)}</span>`;
    return `${sep}<div class="crumb">${inner}</div>`;
  }).join("") + `</div>`;

  // One-time delegated click handler
  if (!window.__sbtsCrumbClickBound) {
    window.__sbtsCrumbClickBound = true;
    document.addEventListener("click", (e) => {
      const a = e.target && e.target.closest ? e.target.closest("a[data-crumb-idx]") : null;
      if (!a) return;
      e.preventDefault();
      const idx = Number(a.getAttribute("data-crumb-idx"));
      const handlers = window.__sbtsCrumbHandlers || [];
      const fn = handlers[idx];
      if (typeof fn === "function") {
        try { fn(); } catch (err) { console.warn("Breadcrumb handler failed", err); }
      }
    }, true);
  }
}
function wfLoadConfig() {
  const raw = SBTS_UTILS.LS.get(WF_STORAGE_KEY);
  if (!raw) return wfDefaultConfig();
  try {
    const cfg = JSON.parse(raw);

    // Minimal validation/migration
    if (!cfg || !Array.isArray(cfg.phases)) return wfDefaultConfig();
    // Sync roles from Roles & Specialties Catalog (single source of truth)
    cfg.roles = getRolesCatalog().map(r => ({
      id: r.id,
      label: r.label,
      type: r.type || "role",
      active: r.active !== false
    }));

    // Migrate phase canUpdate entries (labels -> ids) and ensure array
    cfg.phases.forEach((p) => {
      if (!Array.isArray(p.canUpdate)) p.canUpdate = [];
      p.canUpdate = p.canUpdate.map((x) => normalizeRoleId(x)).filter(Boolean);
    });

    if (!cfg.ui) cfg.ui = { layout: "horizontal" };
    if (!cfg.ui.layout) cfg.ui.layout = "horizontal";

    if (!cfg.rules) cfg.rules = {};
    if (!cfg.rules.slipBlindMetalForeman) {
      cfg.rules.slipBlindMetalForeman = {
        enabled: false,
        role: "metal_foreman",
        name: "Metal Foreman – Slip Blind Confirmation",
        type: "confirmation",
        options: ["Reinstalled", "Demolished"]
      };
    } else {
      // Ensure required fields exist
      if (typeof cfg.rules.slipBlindMetalForeman.enabled !== "boolean") cfg.rules.slipBlindMetalForeman.enabled = !!cfg.rules.slipBlindMetalForeman.enabled;
      if (!cfg.rules.slipBlindMetalForeman.role) cfg.rules.slipBlindMetalForeman.role = "metal_foreman";
      if (!cfg.rules.slipBlindMetalForeman.name) cfg.rules.slipBlindMetalForeman.name = "Metal Foreman – Slip Blind Confirmation";
      if (!cfg.rules.slipBlindMetalForeman.type) cfg.rules.slipBlindMetalForeman.type = "confirmation";
      if (!Array.isArray(cfg.rules.slipBlindMetalForeman.options)) cfg.rules.slipBlindMetalForeman.options = ["Reinstalled", "Demolished"];
    }

    wfEnsurePhaseColors(cfg);
    return cfg;
  } catch {
    return wfDefaultConfig();
  }
}


function wfSaveConfig(cfg) {
  SBTS_UTILS.LS.set(WF_STORAGE_KEY, JSON.stringify(cfg));
}

function wfRoleLabel(roleId) {
  return ROLE_LABELS[roleId] || roleId;
}

function wfApplyLayout() {
  const grid = document.querySelector("#workflowControlPage .wf-grid");
  if (!grid) return;
  const layout = (wfState.cfg?.ui?.layout || "horizontal");
  grid.classList.toggle("vertical", layout === "vertical");

  const hb = document.getElementById("wfLayoutHorizontalBtn");
  const vb = document.getElementById("wfLayoutVerticalBtn");
  if (hb) hb.classList.toggle("active", layout === "horizontal");
  if (vb) vb.classList.toggle("active", layout === "vertical");
}

function wfBindLayoutControls() {
  const hb = document.getElementById("wfLayoutHorizontalBtn");
  const vb = document.getElementById("wfLayoutVerticalBtn");
  if (hb) hb.onclick = () => { wfState.cfg.ui.layout = "horizontal"; wfSaveConfig(wfState.cfg); wfApplyLayout(); };
  if (vb) vb.onclick = () => { wfState.cfg.ui.layout = "vertical"; wfSaveConfig(wfState.cfg); wfApplyLayout(); };
}

function wfBindSlipRuleToggle() {
  const t = document.getElementById("wfSlipRuleToggle");
  if (!t) return;
  const rule = wfState.cfg?.rules?.slipBlindMetalForeman || { enabled: true, role: "metal_foreman", name: "Metal Foreman – Slip Blind Confirmation" };
  t.checked = !!rule.enabled;
  t.onchange = () => {
    if (!wfState.cfg.rules) wfState.cfg.rules = {};
    if (!wfState.cfg.rules.slipBlindMetalForeman) wfState.cfg.rules.slipBlindMetalForeman = rule;
    wfState.cfg.rules.slipBlindMetalForeman.enabled = !!t.checked;
    wfSaveConfig(wfState.cfg);
  };
}

let wfState = { cfg: null, selectedPhaseId: null, dragId: null };

function renderWorkflowControlPage() {
  wfState.cfg = wfLoadConfig();
  if (!wfState.selectedPhaseId) wfState.selectedPhaseId = wfState.cfg.phases[0]?.id || "broken";
  renderWfPhaseList();
  renderWfSelected();
  if (!wfState.cfg.ui) wfState.cfg.ui = { layout: "horizontal" };
  wfBindLayoutControls();

  // Link: manage roles (opens Settings -> Roles & Specialties Manager)
  const mrl = document.getElementById("wfManageRolesLink");
  if (mrl) {
    mrl.onclick = () => {
      try { openPage("settingsPage"); } catch (_) {}
      setTimeout(() => {
        const el = document.getElementById("rolesCatalogCard");
        if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 60);
    };
  }

  wfApplyLayout();
}

function getWfPhase(id) {
  return wfState.cfg.phases.find((p) => p.id === id) || null;
}

function renderWfPhaseList() {
  const list = document.getElementById("wfPhaseList");
  if (!list) return;
  list.innerHTML = "";
  const phasesSorted = [...wfState.cfg.phases].sort((a,b)=> (a.order||0)-(b.order||0));
  phasesSorted.forEach((p) => {
    const item = document.createElement("div");
    item.className = "wf-item" + (p.id === wfState.selectedPhaseId ? " active" : "");
    item.setAttribute("draggable", "true");
    item.dataset.id = p.id;

    item.ondragstart = (e) => { wfState.dragId = p.id; e.dataTransfer.effectAllowed="move"; };
    item.ondragover = (e) => { e.preventDefault(); e.dataTransfer.dropEffect="move"; };
    item.ondrop = (e) => {
      e.preventDefault();
      const fromId = wfState.dragId;
      const toId = p.id;
      if (!fromId || fromId === toId) return;
      wfReorderPhase(fromId, toId);
    };

    const left = document.createElement("div");
    left.className="wf-left";
    const dot = `<span style="width:10px;height:10px;border-radius:999px;background:${escapeHTML(p.color||wfPhaseColor(p.id))};display:inline-block;box-shadow:0 0 0 1px rgba(0,0,0,0.08)"></span>`;
    left.innerHTML = `<span class="wf-handle">☰</span>${dot}<div><div style="font-weight:700">${escapeHTML(p.label)}</div><div class="wf-muted">${p.id}${p.active? "" : " • disabled"}</div></div>`;

    const right = document.createElement("div");
    right.className="wf-muted";
    right.textContent = `#${p.order||""}`;

    item.appendChild(left);
    item.appendChild(right);

    item.onclick = () => { wfState.selectedPhaseId = p.id; renderWfPhaseList(); renderWfSelected(); };

    list.appendChild(item);
  });
  wfSaveConfig(wfState.cfg);
}

function wfReorderPhase(fromId, toId) {
  const phases = [...wfState.cfg.phases].sort((a,b)=>(a.order||0)-(b.order||0));
  const fromIndex = phases.findIndex(p=>p.id===fromId);
  const toIndex = phases.findIndex(p=>p.id===toId);
  if (fromIndex<0 || toIndex<0) return;
  const [moved] = phases.splice(fromIndex,1);
  phases.splice(toIndex,0,moved);
  phases.forEach((p,idx)=> p.order = idx+1);
  // write back
  wfState.cfg.phases.forEach((p)=>{ p.order = phases.find(x=>x.id===p.id).order; });
  renderWfPhaseList();
  renderWfSelected();
}

function renderWfSelected() {
  const p = getWfPhase(wfState.selectedPhaseId);
  if (!p) return;
  document.getElementById("wfSelectedPhaseBadge").textContent = p.label;
  document.getElementById("wfPhaseNameInput").value = p.label;
  document.getElementById("wfPhaseActiveChk").checked = !!p.active;
  document.getElementById("wfPhaseKeyText").textContent = `Key: ${p.id}`;

  // color
  const colorInput = document.getElementById("wfPhaseColorInput");
  const colorHex = document.getElementById("wfPhaseColorHex");
  const colorPrev = document.getElementById("wfPhaseColorPreview");
  if (colorInput) {
    if (!p.color) p.color = wfPhaseColor(p.id);
    colorInput.value = p.color;
    if (colorHex) colorHex.textContent = p.color;
    if (colorPrev) colorPrev.style.background = p.color;
    colorInput.oninput = (e) => {
      p.color = e.target.value;
      if (colorHex) colorHex.textContent = p.color;
      if (colorPrev) colorPrev.style.background = p.color;
      renderWfPhaseList();
      wfSaveConfig(wfState.cfg);
      renderDashboard();
      // refresh workflow steps if on blind details
      const blind = state.blinds.find(b => b.id === state.currentBlindId);
      if (blind && document.getElementById("blindDetailsPage") && !document.getElementById("blindDetailsPage").classList.contains("hidden")) {
        renderWorkflowSteps(blind);
      }
    };
  }

  // chips: canUpdate
  const canWrap = document.getElementById("wfCanUpdateChips");
  canWrap.innerHTML = "";
  wfState.cfg.roles.filter(r=>r.active).forEach((r)=>{
    const chip = document.createElement("div");
    chip.className = "wf-chip" + (p.canUpdate.includes(r.id) ? " on" : "");
    chip.textContent = r.label;
    chip.onclick = ()=>{ toggleInArray(p.canUpdate, r.id); chip.classList.toggle("on"); wfSaveConfig(wfState.cfg); };
    canWrap.appendChild(chip);
  });

  // approval
  document.getElementById("wfApprovalEnabledChk").checked = !!p.approval?.enabled;
  const appWrap = document.getElementById("wfApprovalRolesChips");
  appWrap.innerHTML="";
  wfState.cfg.roles.filter(r=>r.active).forEach((r)=>{
    const chip=document.createElement("div");
    chip.className="wf-chip"+(p.approval?.roles?.includes(r.id)?" on":"");
    chip.textContent=r.label;
    chip.onclick=()=>{
      if (!p.approval) p.approval={enabled:false,roles:[],mode:"parallel"};
      toggleInArray(p.approval.roles, r.id);
      chip.classList.toggle("on");
      wfSaveConfig(wfState.cfg);
    };
    appWrap.appendChild(chip);
  });

  document.getElementById("wfApprovalEnabledChk").onchange = (e)=>{
    if (!p.approval) p.approval={enabled:false,roles:[],mode:"parallel"};
    p.approval.enabled = e.target.checked;
    wfSaveConfig(wfState.cfg);
  };

  // extra approvals
  document.getElementById("wfExtraEnabledChk").checked = !!p.extra?.enabled;
  document.getElementById("wfExtraEnabledChk").onchange=(e)=>{
    if (!p.extra) p.extra={enabled:false,items:[]};
    p.extra.enabled = e.target.checked;
    if (!p.extra.enabled) p.extra.items=[];
    renderWfExtraList();
    wfSaveConfig(wfState.cfg);
  };
  renderWfExtraList();

  // name & active handlers
  document.getElementById("wfPhaseNameInput").oninput = (e)=>{ p.label = e.target.value; renderWfPhaseList(); document.getElementById("wfSelectedPhaseBadge").textContent=p.label; wfSaveConfig(wfState.cfg); };
  document.getElementById("wfPhaseActiveChk").onchange = (e)=>{ p.active = e.target.checked; renderWfPhaseList(); wfSaveConfig(wfState.cfg); };
}

function renderWfExtraList() {
  const p = getWfPhase(wfState.selectedPhaseId);
  const list = document.getElementById("wfExtraList");
  list.innerHTML="";
  if (!p.extra?.enabled) {
    list.innerHTML = `<div class="wf-muted">No extra approvals required for this phase.</div>`;
    return;
  }
  if (!p.extra.items?.length) {
    list.innerHTML = `<div class="wf-muted">Add extra approvals if needed (optional).</div>`;
    return;
  }
  p.extra.items.forEach((it,idx)=>{
    const row=document.createElement("div");
    row.className="wf-item";
    row.innerHTML = `<div class="wf-left"><div><div style="font-weight:700">${escapeHTML(it.name)}</div><div class="wf-muted">${(it.roles||[]).map(wfRoleLabel).join(", ") || "No roles"}${it.required? " • required": ""}</div></div></div><div class="wf-muted">✎</div>`;
    row.onclick=()=> wfEditExtra(idx);
    list.appendChild(row);
  });
}

function wfAddPhase() {
  const name = prompt("New phase name:");
  if (!name) return;
  const key = "custom_" + Math.random().toString(16).slice(2,8);
  const maxOrder = Math.max(...wfState.cfg.phases.map(p=>p.order||0), 0);
  // assign an auto color (can be changed by admin)
  const used = wfState.cfg.phases.map(p => p.color).filter(Boolean);
  const tmp = { id: key, order: maxOrder + 1 };
  const autoColor = (WF_DEFAULT_PHASE_COLORS[key] || "") || (wfPhaseColor(tmp.id) || "#2563eb");
  // ensure uniqueness by letting ensureColors re-pick if needed
  wfState.cfg.phases.push({ id:key, label:name, color: autoColor, active:true, order:maxOrder+1, canUpdate:[], approval:{enabled:false,roles:[],mode:"parallel"}, extra:{enabled:false,items:[]} });
  wfEnsurePhaseColors(wfState.cfg);
  wfState.selectedPhaseId = key;
  renderWfPhaseList();
  renderWfSelected();
}

function wfSaveSelectedPhase() {
  wfSaveConfig(wfState.cfg);
  toast("Saved.");
}

function wfDisableSelectedPhase() {
  const p = getWfPhase(wfState.selectedPhaseId);
  if (!p) return;
  if (!confirm("Disable this phase? (Recommended over delete)")) return;
  p.active = false;
  renderWfPhaseList();
  renderWfSelected();
  wfSaveConfig(wfState.cfg);
}

function wfAddExtra() {
  const p = getWfPhase(wfState.selectedPhaseId);
  if (!p) return;
  if (!p.extra) p.extra={enabled:true,items:[]};
  p.extra.enabled = true;
  const name = prompt("Extra approval name (e.g. Metal Foreman – Demolish):");
  if (!name) return;
  const roles = prompt("Roles (comma-separated role IDs) e.g. metal_foreman") || "";
  const roleIds = roles.split(",").map(s=>s.trim()).filter(Boolean);
  p.extra.items.push({ name, roles: roleIds, required: true });
  renderWfExtraList();
  wfSaveConfig(wfState.cfg);
}

function wfEditExtra(idx) {
  const p = getWfPhase(wfState.selectedPhaseId);
  const it = p?.extra?.items?.[idx];
  if (!it) return;
  const name = prompt("Edit name:", it.name);
  if (name === null) return;
  const roles = prompt("Edit roles (comma-separated role IDs):", (it.roles||[]).join(","));
  if (roles === null) return;
  const required = confirm("Required? OK=yes, Cancel=no");
  it.name = name.trim() || it.name;
  it.roles = roles.split(",").map(s=>s.trim()).filter(Boolean);
  it.required = required;
  wfSaveConfig(wfState.cfg);
  renderWfExtraList();
}

function wfResetDefaults() {
  if (!confirm("Reset workflow config to defaults?")) return;
  sbtsAutoBackupBefore("Reset Workflow Config");
  wfState.cfg = wfDefaultConfig();
  wfState.selectedPhaseId = wfState.cfg.phases[0]?.id || "broken";
  wfSaveConfig(wfState.cfg);
  renderWorkflowControlPage();
}

function wfExport() {
  const cfg = wfLoadConfig();
  const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "sbts_workflow_config_v1.json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function wfImport() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "application/json";
  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const cfg = JSON.parse(reader.result);
        if (!cfg || !cfg.phases || !cfg.roles) throw new Error("Bad schema");
        wfSaveConfig(cfg);
        wfState.cfg = cfg;
        wfState.selectedPhaseId = cfg.phases[0]?.id || "broken";
        renderWorkflowControlPage();
        toast("Imported.");
      } catch (e) {
        alert("Import failed: " + e.message);
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

function toggleInArray(arr, value) {
  if (!Array.isArray(arr)) return;
  const i = arr.indexOf(value);
  if (i >= 0) arr.splice(i, 1);
  else arr.push(value);
}

function escapeHTML(str) {
  return String(str || "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

/* ==========================
   D2: Sandbox Move Permission enforcement (TEST-WORKFLOW only)
========================== */
function wfIsSandboxBlind(blind) {
  const proj = state.projects.find((p) => p.id === blind?.projectId);
  return !!(proj && (proj.name || "").trim() === WF_SANDBOX_PROJECT_NAME);
}

function wfAllowedRolesForPhaseMove(toPhaseId) {
  const cfg = wfLoadConfig();
  const phase = cfg.phases.find((p) => p.id === toPhaseId);
  if (phase && Array.isArray(phase.canUpdate) && phase.canUpdate.length) return phase.canUpdate;
  // fallback to legacy required role mapping
  const req = WORKFLOW_REQUIRED_ROLE[toPhaseId];
  return req ? [req] : [];
}

// Patch canTransition to enforce move-permissions in sandbox when enabled
const __orig_canTransition = canTransition;
canTransition = function(fromPhase, toPhase, user, blind = null) {
  const base = __orig_canTransition(fromPhase, toPhase, user, blind);
  if (!base) return false;
  if (!state.ui?.workflowSandboxEnforcement) return base;

  const activeBlind = blind || state.blinds.find((b) => b.id === state.currentBlindId) || null;
  if (!activeBlind) return base;
  if (!wfIsSandboxBlind(activeBlind)) return base;
  if (user?.role === "admin") return true;

  const allowed = wfAllowedRolesForPhaseMove(toPhase);
  if (!allowed.includes(user?.role)) return false;
  return true;
};


/* ================================
   Final Approvals Manager (Safe Mode)
   UI + display-lookup only
================================ */
const SBTS_FA_DIR_KEY = "sbts_finalApprovalsDirectory_v1";
const SBTS_FA_SAFE_MODE_KEY = "sbts_finalApprovalsSafeMode_v1";
const SBTS_FA_REDIRECTS_KEY = "sbts_finalApprovalsRedirects_v1";

function faLoadRedirects(){
  try{ return JSON.parse(SBTS_UTILS.LS.get(SBTS_FA_REDIRECTS_KEY) || "{}") || {}; }catch(e){ return {}; }
}
function faSaveRedirects(map){
  SBTS_UTILS.LS.set(SBTS_FA_REDIRECTS_KEY, JSON.stringify(map||{}));
}
function faResolveId(id){
  const map = faLoadRedirects();
  let cur = (id||"").trim();
  const seen = new Set();
  while(map[cur] && !seen.has(cur)){
    seen.add(cur);
    cur = String(map[cur]).trim();
  }
  return cur;
}
function faIsIdUsedInBlinds(id){
  try{
    const key = (id||"").trim();
    const blinds = (state && Array.isArray(state.blinds)) ? state.blinds : [];
    return blinds.some(b => b && b.finalApprovals && b.finalApprovals[key]);
  }catch(e){ return false; }
}
function faIsIdApprovedForKey(faObj, requiredKey){
  const k = (requiredKey||"").trim();
  if(faObj && faObj[k] && faObj[k].status==="approved") return true;
  const redirects = faLoadRedirects();
  // any old keys that resolve to k
  for(const oldKey in redirects){
    if(!oldKey) continue;
    if(faResolveId(oldKey) === k){
      if(faObj && faObj[oldKey] && faObj[oldKey].status==="approved") return true;
    }
  }
  return false;
}


function faDefaultDirectory(){
  return [
    { id:"METAL_FOREMAN", name:"Foreman Metal", applies:"all", status:"active" },
    { id:"QAQC", name:"QA/QC", applies:"all", status:"active" },
    { id:"SAFETY", name:"Safety Officer", applies:"isolation", status:"active" },
  ];
}

function faNormalizeDirectory(dir){
  // Ensure each item has roles[], required, order
  let changed = false;

  dir.forEach((x, idx) => {
    if(!x.status){ x.status = "active"; changed = true; }
    if(!Array.isArray(x.roles)){
      x.roles = x.roles ? (Array.isArray(x.roles) ? x.roles : [x.roles]) : [];
      changed = true;
    }
    if(typeof x.required !== "boolean"){
      x.required = true; // default required ON
      changed = true;
    }
    if(typeof x.order !== "number" || isNaN(x.order) || x.order === 0){
      x.order = (idx+1) * 10; // default 10,20,30...
      changed = true;
    }
  });

  return { dir, changed };
}

function faLoadDirectory(){
  try{
    const raw = SBTS_UTILS.LS.get(SBTS_FA_DIR_KEY);
    if(!raw) return faDefaultDirectory();
    const parsed = JSON.parse(raw);
    if(!Array.isArray(parsed)) return faDefaultDirectory();
    const norm = faNormalizeDirectory(parsed);
    if(norm.changed){ faSaveDirectory(norm.dir); }
    return norm.dir;
  }catch(e){
    return faDefaultDirectory();
  }
}

function faSaveDirectory(dir){
  if (!requirePerm("manageFinalApprovals")) return;
  SBTS_UTILS.LS.set(SBTS_FA_DIR_KEY, JSON.stringify(dir));
}

function faIsSafeModeOn(){
  return SBTS_UTILS.LS.get(SBTS_FA_SAFE_MODE_KEY) === "1";
}

function faSetSafeMode(on){
  if (!requirePerm("manageFinalApprovals")) return;
  SBTS_UTILS.LS.set(SBTS_FA_SAFE_MODE_KEY, on ? "1" : "0");
}

function faLookupName(approvalId, fallbackName){
  if(!faIsSafeModeOn()) return fallbackName || approvalId || "";
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (approvalId||"").toUpperCase());
  return item?.name || fallbackName || approvalId || "";
}

function faLookupMeta(approvalId){
  const dir = faLoadDirectory();
  return dir.find(x => (x.id||"").toUpperCase() === (approvalId||"").toUpperCase()) || null;
}

function faResetDirectory(){
  if (!requirePerm("manageFinalApprovals")) return;
  if(!confirm("Reset Final Approvals Directory to defaults?\n\nThis will restore default items (recommended). History remains safe.")) return;
  sbtsAutoBackupBefore("Reset Final Approvals Directory");
  const def = faDefaultDirectory();
  faSaveDirectory(def);
  faRenderTable();
  toast("Defaults restored");
}

function faRenderTable(){
  if (!requirePerm("manageFinalApprovals")) return;
  const tbody = document.getElementById("famTbody");
  if (!tbody) return;
  if (!canUser("manageFinalApprovals")) {
    tbody.innerHTML = `<tr><td colspan="8" class="muted">🔒 Access denied.</td></tr>`;
    return;
  }
  if(!tbody) return;

  const q = (document.getElementById("famSearch")?.value || "").trim().toLowerCase();
  const status = document.getElementById("famFilterStatus")?.value || "all";
  const applies = document.getElementById("famFilterApplies")?.value || "all";

  let dir = faLoadDirectory();

  // default: sort by order then name (professional)
  dir = dir.slice().sort((a,b)=>{
    const ao = (typeof a.order === "number" ? a.order : 9999);
    const bo = (typeof b.order === "number" ? b.order : 9999);
    if(ao !== bo) return ao - bo;
    return String(a.name||"").localeCompare(String(b.name||""));
  });

  if(q){
    dir = dir.filter(x => (x.name||"").toLowerCase().includes(q) || (x.id||"").toLowerCase().includes(q));
  }
  if(status !== "all"){
    dir = dir.filter(x => (x.status||"active") === status);
  }
  if(applies !== "all"){
    dir = dir.filter(x => (x.applies||"all") === applies);
  }

  tbody.innerHTML = dir.map((x,i)=>{
    const appliesBadge = x.applies === "isolation" ? "badge-blue" : (x.applies === "slip" ? "badge-gray" : "badge-gray");
    const statusBadge = x.status === "disabled" ? "badge-red" : (x.status === "archived" ? "badge-gray" : "badge-green");
    const statusLabel = x.status === "disabled" ? "Disabled" : (x.status === "archived" ? "Archived" : "Active");
    const appliesLabel = x.applies === "isolation" ? "Isolation" : (x.applies === "slip" ? "Slip" : "All");

    const rolesHtml = (x.roles && x.roles.length)
      ? x.roles.map(r=>'<span class="chip">'+escapeHtml(r)+'</span>').join(" ")
      : '<span class="tiny muted">—</span>';

    const toggleLabel = (x.status === "disabled") ? "Enable" : "Disable";
    const archiveLabel = (x.status === "archived") ? "Restore" : "Archive";

    return `
      <tr>
        <td>${i+1}</td>
        <td><b>${escapeHtml(x.name||"")}</b><div class="tiny muted">ID: ${escapeHtml(x.id||"")}</div></td>
        <td><span class="badge ${appliesBadge}">${appliesLabel}</span></td>
        <td><div class="tiny">${rolesHtml}</div></td>
        <td>
          <label class="switch fam-required-mini" title="Required = must be completed">
            <input type="checkbox" ${(x.required!==false)?"checked":""} onchange="faQuickSetRequired('${escapeHtml(x.id)}', this.checked)">
            <span class="slider"></span>
          </label>
        </td>
        <td>
          <input class="input fam-quick-order" type="number" value="${(typeof x.order==='number')?x.order:10}" onchange="faQuickSetOrder('${escapeHtml(x.id)}', this.value)" title="Display order (10,20,30...)">
        </td>
        <td><span class="badge ${statusBadge}">${statusLabel}</span></td>
        <td class="table-actions">
          <button class="secondary-btn" onclick="faOpenEdit('${escapeHtml(x.id)}')">Edit</button>
          <button class="secondary-btn" onclick="faToggleStatus('${escapeHtml(x.id)}')">${toggleLabel}</button>
          <button class="secondary-btn" onclick="faArchiveOrRestore('${escapeHtml(x.id)}')">${archiveLabel}</button>
          <button class="secondary-btn" onclick="faRedirectPrompt('${escapeHtml(x.id)}')" title="Redirect old ID to a new ID (keeps history safe)">Redirect</button>
          <button class="danger-btn" onclick="faHardDelete('${escapeHtml(x.id)}')">Delete</button>
        </td>
      </tr>
    `;
  }).join("");
}
function faQuickSetRequired(id, checked){
  if (!requirePerm("manageFinalApprovals")) return;
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  item.required = !!checked;
  faSaveDirectory(dir);
  toast("Updated: Required = " + (item.required ? "ON" : "OFF"));
}

function faQuickSetOrder(id, val){
  if (!requirePerm("manageFinalApprovals")) return;
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  const num = parseInt(val,10);
  item.order = isNaN(num) ? (item.order||10) : num;
  faSaveDirectory(dir);
  faRenderTable();
  famWireEditModal();
  toast("Updated: Order = " + item.order);
}

let famEditCurrentId = null;
let famEditMode = "edit";

let famEditModalWired = false;


function famCloseOverlay(){
  const o = document.getElementById("famEditOverlay");
  if(o) o.style.display = "none";
  famEditCurrentId = null;
  famEditMode = "edit";
}

function faOpenEdit(id){
  if (!requirePerm("manageFinalApprovals")) return;
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  famEditCurrentId = item.id;
  famEditMode = "edit";
  famPopulateModalForItem(item, false);
  famOpenOverlay();
}

function getAllRolesForFAM(){
  const fallback = ["operation_foreman","supervisor","qaqc","safety","inspection","ti_engineer","metal_foreman"];
  try{
    if(typeof loadRolesCatalog === "function"){
      const cat = loadRolesCatalog();
      const roles = (cat||[])
        .filter(x => x && x.type === "role" && (x.status||"active") !== "hidden")
        .map(x => x.slug || x.id || x.name)
        .filter(Boolean)
        .map(r => String(r));
      const uniq = [...new Set(roles)];
      if(uniq.length) return uniq;
    }
  }catch(e){}
  return fallback;
}

function famWireEditModal(){
  if (!requirePerm("manageFinalApprovals")) return;
  if (!canUser("manageFinalApprovals")) {
    // section hidden, but guard in case of direct calls
    return;
  }
  if(famEditModalWired) return;
  const overlay = document.getElementById("famEditOverlay");
  if(!overlay) return;
  famEditModalWired = true;

  document.getElementById("famEditCloseBtn")?.addEventListener("click", famCloseOverlay);
  document.getElementById("famEditCancelBtn")?.addEventListener("click", famCloseOverlay);

  overlay.addEventListener("click", (e)=>{
    if(e.target === overlay) famCloseOverlay();
  });

  document.querySelectorAll(".fam-tab").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const tab = btn.getAttribute("data-tab");
      document.querySelectorAll(".fam-tab").forEach(b=>b.classList.remove("active"));
      document.querySelectorAll(".fam-panel").forEach(p=>p.classList.remove("active"));
      btn.classList.add("active");
      document.querySelector(`.fam-panel[data-panel="${tab}"]`)?.classList.add("active");
    });
  });

  document.getElementById("famEditSaveBtn")?.addEventListener("click", ()=>{
    const dir = faLoadDirectory();

    // collect fields
    const name = (document.getElementById("famEditName").value || "").trim();
    const applies = document.getElementById("famEditApplies").value || "all";
    const required = !!document.getElementById("famEditRequired").checked;
    const num = parseInt(document.getElementById("famEditOrder").value,10);
    const order = isNaN(num) ? faNextOrder() : num;

    const picked = [];
    const wrap = document.getElementById("famEditRolesWrap");
    if(wrap){
      wrap.querySelectorAll("input[type=checkbox][data-role]").forEach(cb=>{
        if(cb.checked) picked.push(cb.getAttribute("data-role"));
      });
    }

    if(famEditMode === "create"){
      let idVal = (document.getElementById("famEditId").value || "").trim().toUpperCase();
      if(!idVal){
        idVal = faGenerateIdFromName(name || "APPROVAL");
        document.getElementById("famEditId").value = idVal;
      }
      if(dir.some(x => (x.id||"").toUpperCase() === idVal)){
        alert("ID already exists. Choose a different ID.");
        return;
      }
      const item = { id: idVal, name: name || idVal, applies, roles: picked, required, order, status:"active" };
      dir.push(item);
      faSaveDirectory(dir);
      faRenderTable();
      toast("Approval added");
      famCloseOverlay();
      return;
    }

    // edit mode
    if(!famEditCurrentId) return;
    const item = dir.find(x => (x.id||"").toUpperCase() === (famEditCurrentId||"").toUpperCase());
    if(!item) return;

    item.name = name || item.name;
    item.applies = applies || item.applies || "all";
    item.required = required;
    item.order = order;
    item.roles = picked;

    faSaveDirectory(dir);
    faRenderTable();
    toast("Approval updated");
    famCloseOverlay();
  });

  // Auto ID button (create mode)
  document.getElementById("famIdAutoBtn")?.addEventListener("click", ()=>{
    if(famEditMode !== "create") return;
    const name = (document.getElementById("famEditName").value || "").trim();
    const idVal = faGenerateIdFromName(name || "APPROVAL");
    document.getElementById("famEditId").value = idVal;
  });

  // Redirect button (edit mode)
  document.getElementById("famRedirectBtn")?.addEventListener("click", ()=>{
    if(!famEditCurrentId) return;
    faRedirectPrompt(famEditCurrentId);
  });
}

function faEnsureInit(){
  // wire events
  const search = document.getElementById("famSearch");
  const fs = document.getElementById("famFilterStatus");
  const fa = document.getElementById("famFilterApplies");
  if(search) search.oninput = faRenderTable;
  if(fs) fs.onchange = faRenderTable;
  if(fa) fa.onchange = faRenderTable;

  const addBtn = document.getElementById("famAddBtn");
  if(addBtn) addBtn.onclick = faAdd;

  const toggle = document.getElementById("faSafeModeToggle");
  if(toggle){
    toggle.checked = faIsSafeModeOn();
    toggle.onchange = () => {
      faSetSafeMode(toggle.checked);
      toast(toggle.checked ? "Safe Mode ON: labels from Final Approvals Manager" : "Safe Mode OFF: classic labels");
      // refresh current views if needed
      try{ renderCurrentBlind?.(); }catch(e){}
      try{ renderCertificatePreview?.(); }catch(e){}
    };
  }

  faRenderTable();
  famWireEditModal();
}


function faGenerateIdFromName(name){
  const base = String(name||"APPROVAL").trim().toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  let id = base || "APPROVAL";
  const dir = faLoadDirectory();
  let i = 2;
  while(dir.some(x => (x.id||"").toUpperCase() === id)){
    id = base + "_" + i;
    i++;
  }
  return id;
}

function faNextOrder(){
  const existing = faLoadDirectory();
  const maxOrder = existing.reduce((m,x)=>Math.max(m, (typeof x.order==="number"?x.order:0)), 0);
  return (Math.floor(maxOrder/10)+1) * 10;
}

function famOpenOverlay(){
  const o = document.getElementById("famEditOverlay");
  if(o) o.style.display = "block";
}

function famPopulateModalForItem(item, isCreate){
  // tabs reset
  document.querySelectorAll(".fam-tab").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".fam-panel").forEach(p=>p.classList.remove("active"));
  document.querySelector('.fam-tab[data-tab="general"]')?.classList.add("active");
  document.querySelector('.fam-panel[data-panel="general"]')?.classList.add("active");

  document.getElementById("famEditName").value = item.name || "";
  document.getElementById("famEditApplies").value = item.applies || "all";
  document.getElementById("famEditRequired").checked = item.required !== false;
  document.getElementById("famEditOrder").value = (typeof item.order==="number" ? item.order : faNextOrder());

  // roles checkboxes
  const roles = Array.isArray(item.roles) ? item.roles : [];
  document.querySelectorAll("#famRolesBox input[type=checkbox][data-role]").forEach(cb=>{
    const r = cb.getAttribute("data-role");
    cb.checked = roles.includes(r);
  });

  const idInput = document.getElementById("famEditId");
  const idHint = document.getElementById("famIdHint");
  const idTag = document.getElementById("famIdLockTag");
  const autoBtn = document.getElementById("famIdAutoBtn");

  if(isCreate){
    idInput.disabled = false;
    idInput.placeholder = "AUTO";
    idInput.value = item.id || "";
    if(idHint) idHint.textContent = "Choose ID (AUTO recommended). ID becomes locked after first use.";
    if(idTag) idTag.textContent = "editable";
    if(autoBtn) autoBtn.style.display = "inline-flex";
  }else{
    idInput.disabled = true;
    idInput.value = item.id || "";
    if(idHint) idHint.textContent = "ID is used in history and must not change.";
    if(idTag) idTag.textContent = "locked";
    if(autoBtn) autoBtn.style.display = "none";
  }

  // show redirect button only in edit
  const redBtn = document.getElementById("famRedirectBtn");
  if(redBtn) redBtn.style.display = isCreate ? "none" : "inline";
}
function faAdd(){
  // open modal in CREATE mode (better than prompts)
  famEditCurrentId = null;
  famEditMode = "create";
  famPopulateModalForItem({ id:"", name:"", applies:"all", roles:[], required:true, order: faNextOrder(), status:"active" }, true);
  famOpenOverlay();
}

function faRename(id){
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  const newName = prompt("New display name:", item.name || "");
  if(!newName) return;
  item.name = newName.trim();
  faSaveDirectory(dir);
  faRenderTable();
  famWireEditModal();
}


function faArchiveOrRestore(id){
  if (!requirePerm("manageFinalApprovals")) return;
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  if(item.status === "archived"){
    item.status = "active";
    item.archivedAt = null;
    item.archivedReason = "";
    faSaveDirectory(dir);
    faRenderTable();
    famWireEditModal();
    toast("Restored");
    return;
  }
  const reason = prompt("Archive reason (optional):","") || "";
  item.status = "archived";
  item.archivedAt = new Date().toISOString();
  item.archivedReason = reason;
  faSaveDirectory(dir);
  faRenderTable();
  famWireEditModal();
  toast("Archived");
}

function faHardDelete(id){
  if (!requirePerm("manageFinalApprovals")) return;
  const key = (id||"").trim();
  if(!key) return;
  if(faIsIdUsedInBlinds(key)){
    alert("Can't delete: this approval ID is used in blinds history/certificates. Use Archive or Redirect instead.");
    return;
  }
  if(!confirm("Delete this approval permanently?\n\nThis is allowed only when it has never been used.")) return;
  sbtsAutoBackupBefore("Delete Final Approval (Permanent)");
  let dir = faLoadDirectory();
  dir = dir.filter(x => (x.id||"").toUpperCase() !== key.toUpperCase());
  // also remove redirects pointing from/to this id
  const redirects = faLoadRedirects();
  let changed = false;
  for(const k in redirects){
    if(k.toUpperCase()===key.toUpperCase() || String(redirects[k]||"").toUpperCase()===key.toUpperCase()){
      delete redirects[k]; changed = true;
    }
  }
  if(changed) faSaveRedirects(redirects);
  faSaveDirectory(dir);
  faRenderTable();
  famWireEditModal();
  toast("Deleted");
}

function faRedirectPrompt(fromId){
  if (!requirePerm("manageFinalApprovals")) return;
  const from = (fromId||"").trim();
  if(!from) return;
  const dir = faLoadDirectory();
  const ids = dir.map(x=>x.id).filter(Boolean).filter(x=>x.toUpperCase()!==from.toUpperCase());
  if(!ids.length){ alert("No target IDs available."); return; }
  const target = prompt("Redirect this ID to which existing ID?\n\nAvailable:\n- " + ids.join("\n- "), ids[0]) || "";
  const to = target.trim().toUpperCase();
  if(!to) return;
  if(!ids.some(x=>x.toUpperCase()===to.toUpperCase())){
    alert("Target ID not found.");
    return;
  }
  const redirects = faLoadRedirects();
  redirects[from] = to;
  faSaveRedirects(redirects);
  // archive the old entry (recommended)
  const item = dir.find(x => (x.id||"").toUpperCase() === from.toUpperCase());
  if(item){
    item.status = "archived";
    item.archivedAt = item.archivedAt || new Date().toISOString();
    item.archivedReason = item.archivedReason || "Redirected to " + to;
    faSaveDirectory(dir);
  }
  faRenderTable();
  famWireEditModal();
  toast("Redirected");
}
function faToggleStatus(id){
  if (!requirePerm("manageFinalApprovals")) return;
  const dir = faLoadDirectory();
  const item = dir.find(x => (x.id||"").toUpperCase() === (id||"").toUpperCase());
  if(!item) return;
  if(item.status==="archived"){ toast("Archived item. Restore first."); return; }
  item.status = item.status === "disabled" ? "active" : "disabled";
  faSaveDirectory(dir);
  faRenderTable();
  famWireEditModal();
}

function faPolicyModelForBlind(blind){
  const type = (blind?.type || "").toLowerCase();
  const isSlip = type.includes("slip");
  const isIsolation = type.includes("isolation");
  const dir = faLoadDirectory()
    .filter(x => (x.status||"active") === "active")
    .filter(x => {
      const a = (x.applies||"all");
      if(a === "all") return true;
      if(a === "slip") return isSlip;
      if(a === "isolation") return isIsolation;
      return true;
    })
    .slice()
    .sort((a,b)=> (a.order??50)-(b.order??50));

  const required = [];
  const optional = [];
  dir.forEach(x=>{
    const entry = {
      key: (x.id||"").trim(),
      label: x.name || x.id,
      role: (x.roles && x.roles.length===1) ? x.roles[0] : (x.role || null),
      roles: Array.isArray(x.roles) ? x.roles : (x.roles ? [x.roles] : []),
      required: x.required !== false
    };
    (entry.required ? required : optional).push(entry);
  });
  return { required, optional };
}




/* ==== SBTS Tooltip Portal (Hover/Focus only, no click) ==== */
(function(){
  if(window.__SBTS_TOOLTIP_PORTAL_HOVER__) return;
  window.__SBTS_TOOLTIP_PORTAL_HOVER__ = true;

  function ensurePortal(){
    let p = document.getElementById('sbts-tip-portal');
    if(!p){
      p = document.createElement('div');
      p.id = 'sbts-tip-portal';
      p.innerHTML = '<div class="box" role="tooltip"></div>';
      document.body.appendChild(p);
    }
    return p;
  }

  function getText(t){
    const dt = t.getAttribute('data-tip');
    if(dt && dt.trim()) return dt.trim();
    const legacy = t.querySelector('.sbts-tip');
    if(legacy) return (legacy.textContent||'').trim();
    return '';
  }

  function place(p, anchor){
    const box = p.querySelector('.box');
    const r = anchor.getBoundingClientRect();
    const margin = 10;

    // force measure
    const bw = box.offsetWidth || 320;
    const bh = box.offsetHeight || 60;

    // Try above-right, else above-left, else below
    let x = r.right + margin;
    let y = r.top - margin - bh;

    if(x + bw > window.innerWidth - margin){
      x = r.left - margin - bw;
    }
    if(x < margin){
      x = window.innerWidth - margin - bw;
      if(x < margin) x = margin;
    }
    if(y < margin){
      y = r.bottom + margin;
    }
    if(y + bh > window.innerHeight - margin){
      y = window.innerHeight - margin - bh;
    }
    if(y < margin) y = margin;

    box.style.left = x + 'px';
    box.style.top  = y + 'px';
  }

  let active = null;
  let raf = null;

  function show(t){
    const text = getText(t);
    if(!text) return;
    const p = ensurePortal();
    const box = p.querySelector('.box');
    box.textContent = text;
    p.classList.add('show');
    const anchor = t.querySelector('.sbts-i') || t;
    place(p, anchor);
    active = t;
  }

  function hide(){
    const p = document.getElementById('sbts-tip-portal');
    if(p) p.classList.remove('show');
    active = null;
  }

  function schedule(){
    if(raf) return;
    raf = requestAnimationFrame(()=>{
      raf = null;
      if(!active) return;
      const p = document.getElementById('sbts-tip-portal');
      if(!p) return;
      const anchor = active.querySelector('.sbts-i') || active;
      place(p, anchor);
    });
  }

  // Hover / focus show
  document.addEventListener('mouseover', (e)=>{
    const t = e.target && e.target.closest ? e.target.closest('.sbts-tooltip') : null;
    if(!t) return;
    show(t);
  }, true);

  document.addEventListener('mouseout', (e)=>{
    const t = e.target && e.target.closest ? e.target.closest('.sbts-tooltip') : null;
    if(!t) return;
    if(t.contains(e.relatedTarget)) return;
    hide();
  }, true);

  document.addEventListener('focusin', (e)=>{
    const t = e.target && e.target.closest ? e.target.closest('.sbts-tooltip') : null;
    if(!t) return;
    show(t);
  }, true);

  document.addEventListener('focusout', (e)=>{
    const t = e.target && e.target.closest ? e.target.closest('.sbts-tooltip') : null;
    if(!t) return;
    hide();
  }, true);

  window.addEventListener('scroll', schedule, true);
  window.addEventListener('resize', schedule);

  // Prevent drag ghosts (applies even if user clicks)
  document.addEventListener('dragstart', (e)=>{
    if(e.target && e.target.closest && e.target.closest('.sbts-tooltip')) e.preventDefault();
  }, true);
})();


/* ==== Tooltip icons tabindex ==== */
(function(){
  if(window.__SBTS_TIP_TABINDEX__) return;
  window.__SBTS_TIP_TABINDEX__ = true;
  function apply(){
    document.querySelectorAll('.sbts-tooltip .sbts-i').forEach(i=>{
      if(!i.hasAttribute('tabindex')) i.setAttribute('tabindex','0');
      i.setAttribute('draggable','false');
    });
  }
  document.addEventListener('DOMContentLoaded', apply);
  setInterval(apply, 1200);
})();


/* ==== SBTS Convert inline Tip/Note text to tooltip icon ==== */
(function(){
  if(window.__SBTS_TIP_TRANSFORM__) return;
  window.__SBTS_TIP_TRANSFORM__ = true;

  function isTipText(t){
    if(!t) return false;
    const s = t.trim();
    return /^((Tip|Note|Hint)\s*:)/i.test(s);
  }
  function cleanText(t){
    return t.replace(/^((Tip|Note|Hint)\s*:)\s*/i,'').trim();
  }

  function makeIcon(text){
    const wrap = document.createElement('span');
    wrap.className='sbts-tooltip';
    wrap.setAttribute('data-tip', text);

    const i = document.createElement('span');
    i.className='sbts-i';
    i.setAttribute('role','button');
    i.setAttribute('aria-label','Help');
    i.setAttribute('tabindex','0');
    i.setAttribute('draggable','false');
    i.textContent='i';
    wrap.appendChild(i);
    return wrap;
  }

  function transform(){
    // targets: any small help lines commonly used across pages
    const nodes = Array.from(document.querySelectorAll(
      '.tiny.muted, .muted.tiny, .help, .help-text, .hint, .note'
    ));

    nodes.forEach(el=>{
      if(el.dataset && el.dataset.tipConverted==='1') return;
      const txt = (el.textContent||'').trim();
      if(!isTipText(txt)) return;

      const tip = cleanText(txt);
      if(!tip) return;

      // Replace content with icon only (keeps spacing)
      el.textContent='';
      el.classList.add('sbts-tip-replaced');
      el.appendChild(makeIcon(tip));
      if(el.dataset) el.dataset.tipConverted='1';
    });
  }

  document.addEventListener('DOMContentLoaded', transform);
  // run again after page switches (SBTS app uses dynamic show/hide)
  setInterval(transform, 1200);
})();


/* ==========================
   PATCH43 - NOTIFICATIONS INBOX SYSTEM
========================== */


function setNotifSavedView(viewId){
  state.ui = state.ui || {};
  state.ui.notifInboxView = viewId || "default";
  saveState();
  renderNotificationsInbox();
}

function setNotifInboxTab(tab) {
  state.ui = state.ui || {};
  state.ui.notifInboxTab = tab;
  state.ui.notifInboxSelected = state.ui.notifInboxSelected || {};
  state.ui.notifInboxSelectedIds = {};
  // update active filter UI
  const box = document.getElementById("notifFilters");
  if (box) {
    Array.from(box.querySelectorAll(".filter-item")).forEach(el => {
      el.classList.toggle("active", el.getAttribute("data-filter") === tab);
    });
  }
  renderNotificationsInbox();
}
function setNotifInboxTab(tab) {
  state.ui = state.ui || {};
  state.ui.notifInboxTab = tab;
  state.ui.notifInboxSelected = state.ui.notifInboxSelected || {};
  state.ui.notifInboxSelectedIds = {};
  // update active filter UI
  const box = document.getElementById("notifFilters");
  if (box) {
    Array.from(box.querySelectorAll(".filter-item")).forEach(el => {
      el.classList.toggle("active", el.getAttribute("data-filter") === tab);
    });
  }
  renderNotificationsInbox();
}

function setNotifInboxUnreadOnly(v){
  state.ui = state.ui || {};
  state.ui.notifInboxUnreadOnly = !!v;
  renderNotificationsInbox();
}
function setNotifInboxProjectFilter(v){
  state.ui = state.ui || {};
  state.ui.notifInboxProjectFilter = v || "all";
  renderNotificationsInbox();
}
function setNotifInboxTypeFilter(v){
  state.ui = state.ui || {};
  state.ui.notifInboxTypeFilter = v || "all";
  // Keep left tabs working; this is an extra layer on top.
  renderNotificationsInbox();
}

function getNotifContextPartsForList(n){
  const parts = [];
  const projects = Array.isArray(state.projects) ? state.projects : [];
  const blinds = Array.isArray(state.blinds) ? state.blinds : [];

  // Project
  const projectName = n.projectId ? (projects.find(p=>p.id===n.projectId)?.name || "") : (n.projectName || "");
  if (projectName) parts.push(projectName);

  // Blind tag / name
  let blindTag = n.blindTag || "";
  if (!blindTag && n.blindId) {
    const b = blinds.find(x=>x.id===n.blindId);
    blindTag = b?.tagNo || b?.tag_no || b?.name || "";
  }
  if (blindTag) parts.push(blindTag);

  // Area
  let area = n.area || "";
  if (!area && n.blindId) {
    const b = blinds.find(x=>x.id===n.blindId);
    area = b?.area || "";
  }
  if (area) parts.push(area);

  // Phase
  const ph = n.phaseId || (n.blindId ? (blinds.find(x=>x.id===n.blindId)?.phase || "") : "");
  if (ph) parts.push(phaseLabel(ph));

  return parts;
}

function buildInboxPrimaryTitle(n){
  // Keep it short and stable (Gmail-like). Avoid mixing BL/Project into title.
  if (!n) return "Notification";
  const kind = notifKind(n);

  // Special-case a few common actions for clarity
  if (n.actionKey && String(n.actionKey).startsWith("user_approval:")) return "New user request";
  if (n.actionKey && String(n.actionKey).startsWith("request:")) return "Request update";
  if (kind === "warning" && n.title) return n.title;

  return n.title || "Notification";
}

function toggleNotificationSelected(notifId, checked) {
  state.ui = state.ui || {};
  state.ui.notifInboxSelectedIds = state.ui.notifInboxSelectedIds || {};
  if (checked) state.ui.notifInboxSelectedIds[notifId] = true;
  else delete state.ui.notifInboxSelectedIds[notifId];
  updateNotifSelectionHint();
}

function toggleSelectAllNotifications(checked) {
  state.ui = state.ui || {};
  state.ui.notifInboxSelectedIds = state.ui.notifInboxSelectedIds || {};
  const listBox = document.getElementById("notificationsInboxList");
  if (!listBox) return;
  const ids = Array.from(listBox.querySelectorAll("input[data-notif-id]")).map(x => x.getAttribute("data-notif-id"));
  if (checked) ids.forEach(id => state.ui.notifInboxSelectedIds[id] = true);
  else ids.forEach(id => delete state.ui.notifInboxSelectedIds[id]);
  // reflect on UI
  Array.from(listBox.querySelectorAll("input[data-notif-id]")).forEach(x => x.checked = checked);
  updateNotifSelectionHint();
}

function updateNotifSelectionHint() {
  const hint = document.getElementById("notifListHint");
  const selectAll = document.getElementById("notifSelectAll");
  state.ui = state.ui || {};
  state.ui.notifInboxSelectedIds = state.ui.notifInboxSelectedIds || {};
  const ids = Object.keys(state.ui.notifInboxSelectedIds || {});
  if (hint) hint.textContent = `${ids.length} selected`;
  if (selectAll) {
    const listBox = document.getElementById("notificationsInboxList");
    const visible = listBox ? Array.from(listBox.querySelectorAll("input[data-notif-id]")).length : 0;
    selectAll.checked = (visible > 0 && ids.length === visible);
  }
}


/* ==========================
   PATCH47.18 - QUICK ACTIONS (DONE)
========================== */
function quickDoneNotification(id){
  if (!id) return;
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  const list1 = (uid && state.notifications.byUser[uid]) ? state.notifications.byUser[uid] : [];
  const list2 = Array.isArray(state.notifications.global) ? state.notifications.global : [];
  const all = list1.concat(list2);
  const n = all.find(x=>x.id===id);
  if (!n) return;
  n.read = true;
  n.resolved = true;
  saveState();
  try { renderNotificationsInbox(); } catch(e){ console.warn(e); }
  try { renderNotificationsDrawer(); } catch(e){ /* ignore */ }
  try { updateNotificationsBadge(); } catch(e){ /* ignore */ }
}
function archiveNotification(notifId) {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;

  const u = (state.users || []).find(x => x.id === uid);
  const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";

  n.archived = true;
  n.archivedTs = Date.now();
  n.read = true;
  if (!n.readTs) n.readTs = Date.now();

  n.activity = Array.isArray(n.activity) ? n.activity : [];
  n.activity.push({ ts: Date.now(), action: "Archived", by: byName });

  saveState();
  renderNotificationsInbox();
  updateNotificationsBadge();
}

function restoreNotification(notifId) {
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;

  const u = (state.users || []).find(x => x.id === uid);
  const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";

  n.archived = false;
  n.restoredTs = Date.now();

  n.activity = Array.isArray(n.activity) ? n.activity : [];
  n.activity.push({ ts: Date.now(), action: "Restored to Inbox", by: byName });

  saveState();
  renderNotificationsInbox();
  updateNotificationsBadge();
}

function archiveSelectedNotifications() {
  state.ui = state.ui || {};
  const ids = Object.keys(state.ui.notifInboxSelectedIds || {});
  if (ids.length === 0) return;
  ids.forEach(id => archiveNotification(id));
  state.ui.notifInboxSelectedIds = {};
  const selectAll = document.getElementById("notifSelectAll");
  if (selectAll) selectAll.checked = false;
  updateNotifSelectionHint();
  renderNotificationsInbox();
}

function deleteSelectedNotifications() {
  state.ui = state.ui || {};
  const ids = Object.keys(state.ui.notifInboxSelectedIds || {});
  if (ids.length === 0) return;
  if (!confirm(`Delete ${ids.length} notification(s)? This removes them only for you.`)) return;
  ids.forEach(id => deleteNotification(id));
  state.ui.notifInboxSelectedIds = {};
  const selectAll = document.getElementById("notifSelectAll");
  if (selectAll) selectAll.checked = false;
  updateNotifSelectionHint();
  renderNotificationsInbox();
}



// ================================
// PATCH 47.28 - Inline Approvals (Workflow Engine v1)
// - Action Panel inside Reader/Modal/Details for Action Required notifications
// - Approve / Reject with comments + Activity logging
// - Optional request/user hooks for demo scenarios
// ================================
function __safeDomId(s){
  try{
    return String(s||"").replace(/[^a-zA-Z0-9_-]/g,"_");
  }catch(e){ return "x"; }
}

function __getNotifByIdForWrite(notifId){
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if(!uid) return null;
  const list = state.notifications.byUser[uid] || [];
  return list.find(x=>x.id===notifId) || null;
}

function __pushNotifActivity(n, action, byName){
  if(!n) return;
  n.activity = Array.isArray(n.activity) ? n.activity : [];
  n.activity.push({ ts: Date.now(), action, by: byName || "You" });
}

function __resolveActionNotification(notifId, outcome, comment){
  const n = __getNotifByIdForWrite(notifId);
  if(!n) return false;

  const uid = getCurrentUserIdStable();
  const u = (state.users || []).find(x => x && x.id === uid);
  const byName = (u && (u.name || u.fullName || u.username)) ? (u.name || u.fullName || u.username) : "You";

  n.read = true;
  if(!n.readTs) n.readTs = Date.now();

  n.resolved = true;
  n.resolvedTs = Date.now();
  n.resolution = outcome; // "approved" | "rejected" | "info_requested"
  if (comment) n.resolutionComment = String(comment).slice(0, 800);

  const label = (outcome === "approved") ? "Approved"
    : (outcome === "rejected") ? "Rejected"
    : (outcome === "info_requested") ? "Info requested"
    : "Action completed";

  __pushNotifActivity(n, label + (comment ? `: ${comment}` : ""), byName);

  // Optional: keep related request object in sync when present
  try{
    const k = String(n.actionKey||"");
    // New user approval demo hook
    if(k.startsWith("user_approval:")){
      const userId = (n.meta && n.meta.userId) ? n.meta.userId : k.split(":")[1];
      const req = (state.requests||[]).find(r => r && r.type === "NEW_USER" && r.meta && r.meta.userId === userId);
      if(req){
        req.status = (outcome === "approved") ? "approved" : (outcome === "rejected" ? "rejected" : req.status);
        req.updatedAt = new Date().toISOString();
        req.comments = Array.isArray(req.comments) ? req.comments : [];
        req.comments.push({
          ts: new Date().toISOString(),
          by: byName,
          message: `${label}${comment ? ` — ${comment}` : ""}`
        });
      }
      // If approved, create/activate user (demo-safe: only if missing)
      if(outcome === "approved" && userId){
        const exists = (state.users||[]).some(x=>x && x.id === userId);
        if(!exists){
          const meta = (n.meta||{});
          (state.users||[]).push({
            id: userId,
            username: meta.username || ("user_"+String(userId).slice(-4)),
            fullName: meta.fullName || meta.name || "New User",
            role: "viewer",
            status: "active"
          });
        } else {
          const uu = (state.users||[]).find(x=>x && x.id===userId);
          if(uu) uu.status = "active";
        }
      }
    }
  }catch(e){ /* keep approval robust */ }

  saveState();
  try{ renderNotificationsInbox(); }catch(_){}
  try{ renderNotificationsDrawer(); }catch(_){}
  try{ updateNotificationsBadge(); }catch(_){}
  // If reader is open, refresh it
  try{
    const rid = state.ui?.notifInboxSelected;
    if(rid === notifId){
      openInboxReader(notifId);
    }
  }catch(_){}
  return true;
}


/* ==========================
   PATCH 47.29 - Smart Inline Approvals (New User Approval)
   - If notification is linked to a NEW_USER request, show role selector
   - Approve will activate user + assign selected role
========================== */

function __getRequestFromNotification(n){
  try{
    if(!n) return null;
    const ak = String(n.actionKey||"");
    if(ak.startsWith("request:")){
      const rid = ak.split(":")[1];
      return (state.requests||[]).find(r=>r && r.id===rid) || null;
    }
  }catch(e){}
  return null;
}

function __upsertUserFromNewUserRequest(req, roleId){
  if(!req) return null;
  const meta = req.meta || {};
  const uid = meta.userId || uid("u");
  let u = (state.users||[]).find(x=>x && x.id===uid);
  const role = roleId || "user";

  if(!u){
    u = {
      id: uid,
      username: meta.username || ("user"+String(uid).slice(-4)),
      fullName: meta.fullName || meta.username || "New User",
      password: "",
      phone: "",
      email: meta.email || "",
      role: role,
      status: "active",
      jobTitle: meta.jobTitle || "",
      profileImage: null,
      createdAt: new Date().toISOString()
    };
    state.users = Array.isArray(state.users) ? state.users : [];
    state.users.unshift(u);
  } else {
    u.status = "active";
    u.role = role;
    if(meta.fullName) u.fullName = meta.fullName;
    if(meta.username) u.username = meta.username;
  }

  // Permissions baseline
  state.permissions = state.permissions || {};
  if(!state.permissions[uid]) state.permissions[uid] = {};
  const isAdmin = String(role).toLowerCase() === "admin";
  const base = {
    manageAreas: isAdmin,
    manageProjects: isAdmin,
    manageBlinds: isAdmin,
    changePhases: true,
    viewReports: isAdmin,
    manageReportsCards: isAdmin,
    manageCertificateSettings: isAdmin,
    manageWorkflowControl: isAdmin,
    editBranding: isAdmin,
    manageTrainingVisibility: isAdmin,
    manageRolesCatalog: isAdmin,
    manageUsers: isAdmin,
    manageRequests: isAdmin,
    manageFinalApprovals: isAdmin,
    managePhaseOwnership: isAdmin
  };
  state.permissions[uid] = Object.assign(base, state.permissions[uid]);

  return u;
}

function approveNewUserInline(notifId){
  const n = __getNotifByIdForWrite(notifId);
  if(!n) return;
  const req = __getRequestFromNotification(n);
  if(!req) return approveNotificationInline(notifId);

  const sid = __safeDomId(notifId);
  const roleSel = document.getElementById(`newUserRole_${sid}`);
  const roleId = roleSel ? roleSel.value : "user";

  const commentEl = document.getElementById(`notifActionComment_${sid}`);
  const comment = commentEl ? String(commentEl.value||"").trim() : "";

  // Activate user
  const user = __upsertUserFromNewUserRequest(req, roleId);

  // Update request status
  req.status = "approved";
  req.updatedAt = new Date().toISOString();
  req.comments = Array.isArray(req.comments) ? req.comments : [];
  if(comment) req.comments.push({ ts: Date.now(), by: (state.currentUser?.fullName||"Admin"), message: comment });

  // Resolve notification + activity
  __resolveActionNotification(notifId, "approved", comment ? comment : `Approved (role: ${getRoleLabelById(roleId)})`);
  __pushNotifActivity(n, `User activated (${user?.username||""})`, state.currentUser?.fullName||"You");

  saveState();
  try { renderUsersTable(); } catch(e) {}
  try { renderNotificationsInbox(); } catch(e) {}
  try { renderNotificationsDrawer(); } catch(e) {}
  showToast("User approved and activated.", "success");
}

function approveNotificationInline(notifId){
  const sid = __safeDomId(notifId);
  const ta = document.getElementById(`notifActionComment_${sid}`);
  const comment = ta ? String(ta.value||"").trim() : "";
  // Approve: comment optional
  __resolveActionNotification(notifId, "approved", comment);
}

function rejectNotificationInline(notifId){
  const sid = __safeDomId(notifId);
  const ta = document.getElementById(`notifActionComment_${sid}`);
  const comment = ta ? String(ta.value||"").trim() : "";
  if(!comment){
    showToast("Comment is required for Reject.", "warning");
    if(ta) ta.focus();
    return;
  }
  __resolveActionNotification(notifId, "rejected", comment);
}

function requestInfoNotificationInline(notifId){
  const sid = __safeDomId(notifId);
  const ta = document.getElementById(`notifActionComment_${sid}`);
  const comment = ta ? String(ta.value||"").trim() : "";
  if(!comment){
    showToast("Please add what info you need.", "warning");
    if(ta) ta.focus();
    return;
  }
  __resolveActionNotification(notifId, "info_requested", comment);
}

function buildInlineApprovalPanelHTML(n){
  if(!n || !(n.requiresAction && !n.resolved)) return "";
  const id = n.id;
  const sid = __safeDomId(id);
  const actionLabel = getNotifPrimaryAction(n)?.label || "Open";
  const kind = notifKind(n);

  // Strong hint for user: read here, act here, open target if needed
  const hint = (kind === "action") ? "Review and take action. Use Open to go to the related item." : "Take action.";

// Smart approvals: detect NEW_USER request to offer Role selection on Approve
const req = __getRequestFromNotification(n);
const isNewUser = !!(req && String(req.type||"").toUpperCase() === "NEW_USER");
const roleOpts = getActiveRoleOptions("role");
const roleOptionsHTML = roleOpts.map(r => `<option value="${esc(r.id)}">${esc(r.label)}</option>`).join("") || `<option value="user">User</option>`;
const approveHandler = isNewUser ? "approveNewUserInline" : "approveNotificationInline";
const roleSelectHTML = isNewUser ? `
      <div class="nap-row">
        <label class="nap-label" for="newUserRole_${sid}">Assign role</label>
        <select id="newUserRole_${sid}" class="input" style="width:100%;max-width:320px;">
          ${roleOptionsHTML}
        </select>
        <div class="tiny muted" style="margin-top:6px;">This will activate the user and assign the selected role.</div>
      </div>
` : ``;

  return `
    <div class="notif-action-panel">
      <div class="nap-head">
        <div>
          <div class="nap-title">Action Panel</div>
          <div class="nap-hint">${esc(hint)}</div>
        </div>
        <div class="nap-primary">
          ${getNotifPrimaryAction(n) ? `<button class="btn sm" onclick="event.stopPropagation(); openNotificationTarget('${esc(id)}')">${esc(actionLabel)}</button>` : ``}
        </div>
      </div>

      <div class="nap-body">
        ${roleSelectHTML}
        <label class="nap-label" for="notifActionComment_${sid}">Comment</label>
        <textarea id="notifActionComment_${sid}" class="nap-textarea" rows="2" placeholder="Add a note (required for Reject)..."></textarea>

        <div class="nap-actions">
          <button class="btn primary" onclick="event.stopPropagation(); ${approveHandler}('${esc(id)}')">Approve</button>
          <button class="btn danger" onclick="event.stopPropagation(); rejectNotificationInline('${esc(id)}')">Reject</button>
          <button class="btn" onclick="event.stopPropagation(); requestInfoNotificationInline('${esc(id)}')">Request Info</button>
          <button class="btn ghost" onclick="event.stopPropagation(); quickDoneNotification('${esc(id)}')">Mark done</button>
        </div>
      </div>
    </div>
  `;
}

function renderNotifDetails(n, panelId = "notifDetails", mode = "panel") {
  const panel = document.getElementById(panelId);
  if (!panel) return;

  if (!n) {
    panel.innerHTML = `<div class="notif-details-empty"><div class="muted">Select a notification to see details.</div></div>`;
    return;
  }

  const proj = n.projectId ? (state.projects || []).find(p => p.id === n.projectId) : null;
  const blind = n.blindId ? (state.blinds || []).find(b => b.id === n.blindId) : null;

  const title = esc(n.title || "Notification");
  const msg = esc(n.message || "");

  const kind = notifKind(n);
  const typeLabel = (n.requiresAction && !n.resolved) ? "Action Required" : (kind === "warning" ? "Warning" : (kind === "admin" ? "Admin" : "System"));
  const badgeClass = (n.requiresAction && !n.resolved) ? "action" : kind;

  // Context (compact)
  const ctxRows = [];
  if (proj) ctxRows.push(`<div class="kv"><b>Project</b><br>${esc(proj.name)}</div>`);
  if (blind) ctxRows.push(`<div class="kv"><b>Blind</b><br>${esc(blind.tag_no || blind.id)}</div>`);
  if (blind && blind.area) ctxRows.push(`<div class="kv"><b>Area</b><br>${esc(blind.area)}</div>`);
  if (blind && blind.line_no) ctxRows.push(`<div class="kv"><b>Line</b><br>${esc(blind.line_no)}</div>`);
  if (blind && blind.size) ctxRows.push(`<div class="kv"><b>Size</b><br>${esc(blind.size)}</div>`);
  if (blind && blind.rating) ctxRows.push(`<div class="kv"><b>Rating</b><br>${esc(blind.rating)}</div>`);
  if (blind && blind.phase) ctxRows.push(`<div class="kv"><b>Phase</b><br>${esc(blind.phase)}</div>`);

  // History (optional, tidy)
  const events = [];
  const pushEv = (label, ts, by) => { if (ts) events.push({ label, ts, by }); };
  pushEv("Created", n.ts, n.createdBy || "System");
  pushEv("Read", n.readTs, "You");
  pushEv("Unread", n.unreadTs, "You");
  pushEv("Action done", n.resolvedTs, "You");
  pushEv("Archived", n.archivedTs, "You");
  pushEv("Restored", n.restoredTs, "You");
  if (Array.isArray(n.activity)) n.activity.forEach(a => pushEv(a.action, a.ts, a.by));
  events.sort((a,b)=> new Date(a.ts).getTime() - new Date(b.ts).getTime());
  const seen = new Set();
  const uniq = events.filter(e => { const k = `${e.label}|${new Date(e.ts).getTime()}`; if (seen.has(k)) return false; seen.add(k); return true; });

  const actionPanel = buildInlineApprovalPanelHTML(n);

  const history = uniq.length ? `
    <details class="notif-history">
      <summary>History</summary>
      <div class="notif-activity">
        ${uniq.slice().reverse().slice(0,12).map(e=>`
          <div class="a-row">
            <div class="a-action">${esc(e.label)}</div>
            <div class="a-meta">${esc(fmtNotifTime(e.ts))}${e.by ? ` • ${esc(e.by)}` : ""}</div>
          </div>
        `).join("")}
      </div>
    </details>
  ` : "";

  // Reader/Modal: toolbar handled outside for reader; inside content just header+body
  const showInnerToolbar = (mode === "modal");
  const toolbar = showInnerToolbar ? `<div class="notif-toolbar">${buildNotifToolbarHTML(n, "modal")}</div>` : "";

  panel.innerHTML = `
    <div class="inbox-details-card inbox-details-rich">
      ${toolbar}
      <div class="detail-row">
        <span class="badge ${badgeClass}">${esc(typeLabel)}</span>
        ${n.read ? '<span class="badge">Read</span>' : '<span class="badge">Unread</span>'}
        ${n.archived ? '<span class="badge">Archived</span>' : ''}
        ${(n.requiresAction && !n.resolved) ? '<span class="badge action">Pending</span>' : ''}
      </div>

      <h2 style="margin-top:10px;">${title}</h2>
      <div class="tiny" style="margin:0 0 14px 0;">${fmtNotifTime(n.ts)} • ${esc(n.createdBy || "System")}</div>

      <div class="notif-message">
        ${msg ? `<div style="white-space:pre-wrap;line-height:1.55;">${msg}</div>` : `<div class="muted">No message details.</div>`}
      </div>

      ${ctxRows.length ? `
        <div class="notif-section" style="margin-top:14px;">
          <div class="notif-section-title">Context</div>
          <div class="notif-ctx-grid">${ctxRows.join("")}</div>
        </div>
      ` : ""}

      ${actionPanel}

      ${history}
    </div>
  `;
}


function openBlindDetailsFromNotif(blindId) {
  // navigate to the blind details page with context preserved
  state.currentBlindId = blindId;
  const b = (state.blinds || []).find(x => x.id === blindId);
  if (b && b.projectId) state.currentProjectId = b.projectId;
  saveState();
  showPage("blindDetailsPage");
  renderBlindDetails();
}

function openProjectFromNotif(projectId) {
  state.currentProjectId = projectId;
  saveState();
  showPage("projectsPage");
  renderProjects();
}

function fmtNotifTime(iso) {
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso || "";
    return d.toLocaleString();
  } catch(e) {
    return iso || "";
  }
}


function notifIsSnoozed(n){
  if (!n) return false;
  if (!n.snoozedUntil) return false;
  const t = (typeof n.snoozedUntil === "number") ? n.snoozedUntil : Date.parse(n.snoozedUntil);
  if (!t || isNaN(t)) return false;
  return Date.now() < t;
}

function snoozeNotification(notifId, minutes){
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;
  const ms = Math.max(1, Number(minutes||0)) * 60 * 1000;
  n.snoozedUntil = Date.now() + ms;
  n.read = true;
  saveState();
  renderNotificationsInbox();
  updateNotificationsBadge();
}

function clearNotificationSnooze(notifId){
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return;
  const list = state.notifications.byUser[uid] || [];
  const n = list.find(x => x.id === notifId);
  if (!n) return;
  n.snoozedUntil = null;
  saveState();
  renderNotificationsInbox();
  updateNotificationsBadge();
}


/* const NOTIF_CATEGORY_META moved above
const NOTIF_CATEGORY_META = {
  action: { label: "Action Required", color: "red", icon: "sbts-ico-action", badge: "Action" },
  system: { label: "System", color: "blue", icon: "sbts-ico-system", badge: "System" },
  warning:{ label: "Warnings", color: "amber", icon: "sbts-ico-warning", badge: "Warning" },
  admin:  { label: "Admin Notes", color: "gray", icon: "sbts-ico-admin", badge: "Admin" },
};
*/



/* ==========================
   Inbox UX helpers (Patch 47.20)
   - Double-click opens standalone page
   - Second click opens Quick View modal
========================== */
function openNotificationStandalone(notifId){
  // (Deprecated) kept for backward compatibility; use Quick View modal instead.
  try{ openNotifQuickModal(notifId); }catch(e){ console.warn(e); }
}

function openNotifQuickModal(notifId){
  const modal = document.getElementById("notifQuickModal");
  const body = document.getElementById("notifQuickModalBody");
  if (!modal || !body) return;

  const n = getNotificationByIdForCurrentUser(notifId);
  if (!n) return;

  // reuse the same renderer into modal body
  body.innerHTML = `<div id="__tmpModalDetails"></div>`;
  renderNotifDetails(n, "__tmpModalDetails", "modal");

  modal.classList.remove("hidden");
  document.addEventListener("keydown", __notifModalEscClose);
}

function closeNotifQuickModal(ev){
  // allow calling without event
  if (ev && ev.target && ev.target.closest && ev.target.closest(".notif-modal-card")) return;
  const modal = document.getElementById("notifQuickModal");
  if (modal) modal.classList.add("hidden");
  const body = document.getElementById("notifQuickModalBody");
  if (body) body.innerHTML = "";
  document.removeEventListener("keydown", __notifModalEscClose);
}
function __notifModalEscClose(e){
  if (e.key === "Escape") closeNotifQuickModal();
}
// PATCH 47.25+ (Gmail-like): open a Reader view inside Inbox (no external page, no always-on right pane)
function openInboxReader(notifId){
  const listView = document.getElementById("inboxListView");
  const reader = document.getElementById("inboxReaderView");
  const body = document.getElementById("inboxReaderBody");
  const tb = document.getElementById("inboxReaderToolbar");
  if (!listView || !reader || !body || !tb) return;

  const n = getNotificationByIdForCurrentUser(notifId);
  if (!n) return;

  // select + mark read (like Gmail)
  try{ state.ui = state.ui || {}; state.ui.notifInboxSelected = notifId; }catch(_){}
  try{ markNotificationRead(notifId); }catch(_){}

  // Toolbar (Gmail-like)
  tb.innerHTML = buildNotifToolbarHTML(n, "reader");

  // Conversation threading (Gmail-like)
const all = getNotificationsForCurrentUser().filter(x => !x.archived);
const tkey = (state.ui?.notifInboxThreadKey) || getNotifThreadKey(n);
const threadItems = all.filter(x => getNotifThreadKey(x) === tkey)
  .sort((a,b)=> new Date(b.ts).getTime() - new Date(a.ts).getTime());

const convoHTML = (threadItems.length > 1) ? `
  <div class="thread-box">
    <div class="thread-head">
      <div class="thread-title">Conversation</div>
      <div class="thread-count">${threadItems.length}</div>
    </div>
    <div class="thread-list">
      ${threadItems.map(x => `
        <button class="thread-row ${x.id===n.id?'active':''}" onclick="event.stopPropagation(); openInboxReader('${esc(x.id)}')">
          <div class="trow-main">
            <div class="trow-title">${esc(buildInboxPrimaryTitle(x) || x.title || 'Notification')}</div>
            <div class="trow-time">${esc(fmtNotifTime(x.ts))}</div>
          </div>
          <div class="trow-sub">${esc((x.message||'').replace(/<[^>]*>/g,'').slice(0,120))}</div>
        </button>
      `).join('')}
    </div>
  </div>` : ``;

// Render rich details into reader body
body.innerHTML = `
  <div id="__tmpReaderDetails"></div>
  ${convoHTML}
`;
renderNotifDetails(n, "__tmpReaderDetails", "reader");

  // Switch views
  listView.classList.add("hidden");
  reader.classList.remove("hidden");

  // Keep list up to date (badges/counters)
  try{ renderNotificationsInbox(); }catch(_){}
}

function closeInboxReader(){
  const listView = document.getElementById("inboxListView");
  const reader = document.getElementById("inboxReaderView");
  const body = document.getElementById("inboxReaderBody");
  const tb = document.getElementById("inboxReaderToolbar");
  if (body) body.innerHTML = "";
  if (tb) tb.innerHTML = "";
  if (reader) reader.classList.add("hidden");
  if (listView) listView.classList.remove("hidden");
}

// Build a Gmail-style toolbar (shared by modal/reader/panel)
function buildNotifToolbarHTML(n, mode){
  const canArchive = !n.archived;
  const btnClose = (mode === "reader")
    ? `<button class="icon-btn" title="Back" onclick="event.stopPropagation(); closeInboxReader()"><i class="ph ph-arrow-left"></i></button>`
    : `<button class="icon-btn" title="Close" onclick="event.stopPropagation(); closeNotifQuickModal()"><i class="ph ph-x"></i></button>`;

  const archiveBtn = canArchive
    ? `<button class="icon-btn" title="Archive" onclick="event.stopPropagation(); archiveNotification('${"${id}"}')"><i class="ph ph-archive"></i></button>`
    : `<button class="icon-btn" title="Restore" onclick="event.stopPropagation(); restoreNotification('${"${id}"}')"><i class="ph ph-arrow-u-up-left"></i></button>`;

  const readBtn = `<button class="icon-btn" title="${"${readTitle}"}" onclick="event.stopPropagation(); toggleNotificationRead('${"${id}"}')">
    <i class="ph ${"${readIcon}"}"></i>
  </button>`;

  const doneBtn = (n.requiresAction && !n.resolved && !n.archived)
    ? `<button class="icon-btn" title="Done" onclick="event.stopPropagation(); resolveNotification('${"${id}"}')"><i class="ph ph-check-circle"></i></button>`
    : ``;

  const delBtn = `<button class="icon-btn danger" title="Delete" onclick="event.stopPropagation(); deleteNotification('${"${id}"}')"><i class="ph ph-trash"></i></button>`;

  
// Primary "Open target" action (Project/Blind/User/Request) as a compact button near the toolbar (not at the bottom)
  let primary = "";
  const pa = getNotifPrimaryAction(n);
  if (pa) {
    primary = `<button class="btn-secondary btn-small" title="${"${hint}"}" onclick="event.stopPropagation(); openNotifTargetFromList('${"${id}"}')">${"${label}"}</button>`
      .replaceAll("${label}", pa.label)
      .replaceAll("${hint}", pa.hint || pa.label);
  }

const html = `
    <div style="display:flex;align-items:center;gap:8px;min-width:0;width:100%;">
      ${btnClose}
      <div class="spacer"></div>
      ${primary}
      ${archiveBtn}
      ${readBtn}
      ${doneBtn}
      ${delBtn}
    </div>
  `;

  return html
    .replaceAll("${id}", n.id)
    .replaceAll("${readTitle}", n.read ? "Mark unread" : "Mark read")
    .replaceAll("${readIcon}", n.read ? "ph-envelope-simple" : "ph-envelope-open")
    .replaceAll("${blindId}", n.blindId || "")
    .replaceAll("${projectId}", n.projectId || "");
}


// Gmail-style Views/Filters rail (hover to expand, leave to collapse)
function sbtsInitInboxFiltersRailHover(){
  const panel = document.getElementById("inboxFiltersPanel");
  const layout = document.getElementById("inboxLayout");
  if (!panel || !layout) return;

  let t = null;
  const collapse = () => {
    if (isInboxViewsPinned()) return;
    panel.classList.add("collapsed");
    layout.classList.add("filters-collapsed");
  };
  const expand = () => {
    panel.classList.remove("collapsed");
    layout.classList.remove("filters-collapsed");
  };

  // Default collapsed unless pinned
  if (isInboxViewsPinned()) expand();
  else collapse();

  try{ updateViewsPinBtn(); }catch(_){}

  panel.addEventListener("mouseenter", () => {
    if (t) clearTimeout(t);
    expand();
  });
  panel.addEventListener("mouseleave", () => {
    if (t) clearTimeout(t);
    t = setTimeout(collapse, 220);
  });
}

function isInboxViewsPinned(){
  try{
    state.ui = state.ui || {};
    return !!state.ui.inboxViewsPinned;
  }catch(_){ return false; }
}
function setInboxViewsPinned(v){
  state.ui = state.ui || {};
  state.ui.inboxViewsPinned = !!v;
  saveState();
  try{ updateViewsPinBtn(); }catch(_){}
}
function updateViewsPinBtn(){
  const btn = document.getElementById("viewsPinBtn");
  if (!btn) return;
  const pinned = isInboxViewsPinned();
  btn.classList.toggle("active", pinned);
  btn.title = pinned ? "Views pinned (click to unpin)" : "Pin Views (keep expanded)";
}
function toggleInboxViewsPinned(ev){
  if (ev && ev.preventDefault) ev.preventDefault();
  if (ev && ev.stopPropagation) ev.stopPropagation();
  const next = !isInboxViewsPinned();
  setInboxViewsPinned(next);
  // If pinned, expand immediately
  const panel = document.getElementById("inboxFiltersPanel");
  const layout = document.getElementById("inboxLayout");
  if (panel && layout && next){
    panel.classList.remove("collapsed");
    layout.classList.remove("filters-collapsed");
  }
}

/* Draggable resizer between Inbox list and Details */
function sbtsInitInboxDetailsResizer(){
  const layout = document.getElementById("inboxLayout");
  const grip = document.getElementById("inboxResizer");
  if (!layout || !grip) return;

  // apply saved width
  try{
    state.ui = state.ui || {};
    const w = Number(state.ui.inboxDetailsWidth || 0);
    if (w && isFinite(w)) layout.style.setProperty("--inbox-details-w", w + "px");
  }catch(_){}

  let dragging = false;
  let startX = 0;
  let startW = 0;

  const minW = 380;

  const onMove = (e) => {
    if (!dragging) return;
    const x = (e.touches && e.touches[0]) ? e.touches[0].clientX : e.clientX;
    const dx = startX - x; // drag left -> larger details
    let w = startW + dx;
    const maxW = Math.max(520, Math.min(820, window.innerWidth - 520));
    w = Math.max(minW, Math.min(maxW, w));
    layout.style.setProperty("--inbox-details-w", w + "px");
  };
  const onUp = () => {
    if (!dragging) return;
    dragging = false;
    grip.classList.remove("dragging");
    document.body.classList.remove("no-select");
    try{
      state.ui = state.ui || {};
      const cur = parseInt(getComputedStyle(layout).getPropertyValue("--inbox-details-w")) || 0;
      if (cur) state.ui.inboxDetailsWidth = cur;
      saveState();
    }catch(_){}
    window.removeEventListener("mousemove", onMove);
    window.removeEventListener("mouseup", onUp);
    window.removeEventListener("touchmove", onMove);
    window.removeEventListener("touchend", onUp);
  };
  const onDown = (e) => {
    e.preventDefault();
    dragging = true;
    grip.classList.add("dragging");
    document.body.classList.add("no-select");
    startX = (e.touches && e.touches[0]) ? e.touches[0].clientX : e.clientX;
    const cur = getComputedStyle(layout).getPropertyValue("--inbox-details-w").trim();
    startW = cur ? parseInt(cur) : 620;
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    window.addEventListener("touchmove", onMove, {passive:false});
    window.addEventListener("touchend", onUp);
  };

  grip.addEventListener("mousedown", onDown);
  grip.addEventListener("touchstart", onDown, {passive:false});
}

// Small helper to safely find a notification for the current user
function getNotificationByIdForCurrentUser(notifId){
  ensureNotificationsState();
  const uid = getCurrentUserIdStable();
  if (!uid) return null;
  const list = state.notifications?.byUser?.[uid] || [];
  return list.find(x => x.id === notifId) || null;
}

function notifKind(n){
  if (n.requiresAction && !n.resolved) return "action";
  const t = (n.type || "system").toLowerCase();
  if (t.includes("warn")) return "warning";
  if (t.includes("admin")) return "admin";
  if (t.includes("system") || t.includes("info")) return "system";
  return "system";
}

function notifIconSVG(kind){
  const k = (kind === "action" || kind === "warning" || kind === "admin" || kind === "system") ? kind : "system";
  const meta = (typeof NOTIF_CATEGORY_META !== "undefined" && NOTIF_CATEGORY_META[k]) ? NOTIF_CATEGORY_META[k] : { icon: "sbts-ico-system" };
  const iconId = meta.icon || "sbts-ico-system";
  return `<svg class="sbts-ico ${k}" aria-hidden="true"><use href="#${iconId}"></use></svg>`;
}

function notifIconLetter(kind){
  // Backward-compatible fallback for older widgets
  if (kind === "action") return "!";
  if (kind === "warning") return "⚠";
  if (kind === "admin") return "N";
  return "i";
}

function notifOpenIconSVG(){
  return `<svg class="sbts-ico system" aria-hidden="true"><use href="#ico-open"></use></svg>`;
}

function notifRestoreIconSVG(){
  return `<svg class="sbts-ico system" aria-hidden="true"><use href="#ico-restore"></use></svg>`;
}

function notifDoneIconSVG(){
  return `<svg class="sbts-ico action" aria-hidden="true"><use href="#ico-done"></use></svg>`;
}
function notifArchiveIconSVG(){
  return `<svg class="sbts-ico system" aria-hidden="true"><use href="#ico-archive"></use></svg>`;
}


// ================================
// 47.10 NAV ROUTER (Notification/Inbox -> Correct Destination)
// - Open buttons go to the right place (Project / Blind / User / Assign Projects / Request Thread)
// - Safe even if data is missing
// ================================
function focusUserRow(userId){
  if (!userId) return;
  state.ui = state.ui || {};
  state.ui.focusUserId = userId;
  saveState();
  // renderUsers() will apply highlight
  setTimeout(() => {
    try{
      const el = document.getElementById(`user_${userId}`);
      if (el && el.scrollIntoView) el.scrollIntoView({behavior:"smooth", block:"center"});
      // clear focus after a short time
      setTimeout(() => { 
        if (state.ui && state.ui.focusUserId === userId){
          state.ui.focusUserId = null;
          saveState();
          // re-render users to remove highlight if still on page
          if (document.getElementById("permissionsPage") && !document.getElementById("permissionsPage").classList.contains("hidden")){
            try{ renderUsers(); }catch(e){}
          }
        }
      }, 2500);
    }catch(e){}
  }, 200);
}

function openUserManagement(userId){
  // Users live in Permissions page
  openPage("permissionsPage");
  try{ renderUsers(); }catch(e){}
  focusUserRow(userId);
}

function openAssignProjectsForUser(userId){
  openUserManagement(userId);
  // open modal if available
  try{
    if (typeof openAssignProjectsModal === "function") openAssignProjectsModal(userId);
  }catch(e){}
}

function openRequestThread(requestId){
  // Use drawer request details (stable) for now
  const nd = ensureNotifDrawerUI();
  nd.selectedReqId = requestId;
  state.ui = state.ui || {};
  state.ui.notifDrawer = state.ui.notifDrawer || {};
  state.ui.notifDrawer.tab = "inbox";
  saveState();
  openNotificationsDrawer();
}

function routeByActionKey(actionKey, meta){
  const k = (actionKey || "").toString();
  if (!k) return false;
  // user approval
  if (k.startsWith("user_approval:")){
    const id = k.split(":")[1];
    openUserManagement(id);
    return true;
  }
  // assign projects
  if (k.startsWith("assign_projects:")){
    const id = k.split(":")[1];
    openAssignProjectsForUser(id);
    return true;
  }
  // open request thread
  if (k.startsWith("request:")){
    const id = k.split(":")[1];
    openRequestThread(id);
    return true;
  }
  // generic user target
  if (k.startsWith("user:")){
    const id = k.split(":")[1];
    openUserManagement(id);
    return true;
  }
  return false;

}



// ================================
// PATCH 47.27 - Inbox/Notifications action engine
// - Fix undefined click handlers (openNotificationInInbox)
// - Unify "Open" actions across Inbox list, Reader view, and Notification Center
// - Route to correct destination (Project / Blind / User / Request / Assignment)
// ================================
function openNotificationInInbox(notifId){
  // Row click: select/highlight only (no navigation)
  try{
    if(!notifId) return;
    state.ui = state.ui || {};
    state.ui.notifInboxSelectedId = notifId;
    saveState();
    // Track selected conversation thread for richer Reader
    try{
      const nn = getNotificationByIdForCurrentUser(notifId);
      state.ui.notifInboxThreadKey = getNotifThreadKey(nn);
      saveState();
    }catch(e){}
    // Re-render to apply selected style
    try{ renderNotificationsInbox(); }catch(e){}
  }catch(e){
    console.warn("[openNotificationInInbox] failed", e);
  }
}

function getNotifPrimaryAction(n){
  // Returns {label, hint} for primary target button, or null if none.
  if(!n) return null;
  const k = String(n.actionKey||"");
  if(k.startsWith("user_approval:") || k.startsWith("user:") || (n.meta && n.meta.userId)) return { label: "Open User", hint: "Go to user" };
  if(k.startsWith("assign_projects:")) return { label: "Open Assignments", hint: "Go to assignments" };
  if(k.startsWith("request:")) return { label: "Open Request", hint: "Go to request thread" };

  const bid = n.blindId || (n.meta && n.meta.blindId);
  if(bid) return { label: "Open Blind", hint: "Go to blind" };

  const pid = n.projectId || (n.meta && n.meta.projectId);
  if(pid) return { label: "Open Project", hint: "Go to project" };

  return null;
}

function openNotifTargetFromList(notifId){
  const n = getNotificationsForCurrentUser().find(x=>x.id===notifId);
  if (!n) return;
  // Mark read when opening
  try { markNotificationRead(notifId); } catch (e) {}
  openNotificationTarget(n);
}



/* ==========================
   PATCH 47.29 - Gmail-Level UX: Conversation Threading + Keyboard
   - Conversation view: group notifications into threads
   - Keyboard navigation (J/K + Enter)
   - Thread selection drives Reader content
========================== */

function getNotifThreadKey(n){
  // Threading policy: Project > Blind (primary grouping)
  // - If project exists: group by projectId, and within it by blindId if present
  // - If no project: group by request/actionKey for non-project items (e.g., new user)
  if(!n) return "misc";
  const pid = n.projectId ? String(n.projectId) : "system";
  const bid = n.blindId ? String(n.blindId) : "none";

  // Requests should stay grouped (one thread per request), but still respect project>blind when present
  const ak = String(n.actionKey || "").trim();
  if (ak && (ak.startsWith("request:") || ak.startsWith("req:"))) {
    // Attach to project/blind bucket if they exist, else by request id
    if (pid !== "system" || bid !== "none") return `pb:${pid}:${bid}:req:${ak.split(":")[1]||""}`;
    return `req:${ak.split(":")[1]||""}`;
  }

  // Final approvals and blind/project events belong to project>blind buckets
  if (pid !== "system" || bid !== "none") return `pb:${pid}:${bid}`;

  // User/admin/system items (no project)
  if (ak) return `ak:${ak}`;
  const t = (n.title||"").toLowerCase().replace(/\s+/g," ").trim().slice(0,40);
  return `misc:${t || "untitled"}`;
}

function buildThreadSummary(items){
  const arr = Array.isArray(items) ? items.slice() : [];
  arr.sort((a,b)=> new Date(b.ts).getTime() - new Date(a.ts).getTime());
  const top = arr[0] || null;
  const unread = arr.filter(x=>!x.read).length;
  const pending = arr.filter(x=>x.requiresAction && !x.resolved).length;
  return { top, count: arr.length, unread, pending };
}

window.__SBTS_NOTIF_VISIBLE_IDS = window.__SBTS_NOTIF_VISIBLE_IDS || [];

function __sbtsBindNotifKeyboardOnce(){
  if (window.__SBTS_NOTIF_KB_BOUND) return;
  window.__SBTS_NOTIF_KB_BOUND = true;

  document.addEventListener("keydown", (ev) => {
    try {
      const activePage = document.querySelector(".page.active");
      if (!activePage || activePage.id !== "notificationsPage") return;

      const tag = (ev.target && ev.target.tagName) ? ev.target.tagName.toLowerCase() : "";
      if (tag === "input" || tag === "textarea" || ev.target?.isContentEditable) return;

      const key = String(ev.key || "").toLowerCase();
      const ids = window.__SBTS_NOTIF_VISIBLE_IDS || [];
      if (!ids.length) return;

      const cur = state.ui?.notifInboxSelectedId || null;
      let idx = cur ? ids.indexOf(cur) : -1;

      if (key === "j") {
        ev.preventDefault();
        idx = Math.min(ids.length - 1, idx + 1);
        const id = ids[idx];
        if (id) { openNotificationInInbox(id); scrollNotifIntoView(id); }
      } else if (key === "k") {
        ev.preventDefault();
        if (idx === -1) idx = 0;
        idx = Math.max(0, idx - 1);
        const id = ids[idx];
        if (id) { openNotificationInInbox(id); scrollNotifIntoView(id); }
      } else if (key === "enter") {
        ev.preventDefault();
        const id = (idx >= 0 ? ids[idx] : ids[0]);
        if (id) openNotifPreviewFromTitle(id);
      }
    } catch(e){ }
  });
}

function scrollNotifIntoView(id){
  try{
    const el = document.querySelector(`.notif-item[data-id="${CSS.escape(id)}"]`);
    if (el) el.scrollIntoView({ block: "nearest" });
  }catch(e){}
}

__sbtsBindNotifKeyboardOnce();

function renderNotificationsInbox() {
  const listBox = document.getElementById("notificationsInboxList");
  const filtersBox = document.getElementById("notifFilters");
  if (!listBox || !filtersBox) return;

  // Admin-only: allow sending admin notes
  const adminBtn = document.getElementById("notifAdminNoteBtn");
  if (adminBtn) adminBtn.style.display = (state.currentUser?.role === "admin") ? "" : "none";

  state.ui = state.ui || {};
  state.ui.notifInboxTab = state.ui.notifInboxTab || "all";
  state.ui.notifInboxSelectedIds = state.ui.notifInboxSelectedIds || {};
  // PATCH47.18 - Saved Views (smart views)
  state.ui.notifInboxView = state.ui.notifInboxView || "default";


  const tab = state.ui.notifInboxTab;
  const q = (document.getElementById("notifInboxSearch")?.value || "").trim().toLowerCase();
  const sort = (document.getElementById("notifSort")?.value || "newest");

  const unreadOnly = !!state.ui.notifInboxUnreadOnly;
  const projectFilter = state.ui.notifInboxProjectFilter || "all";
  const typeFilter = state.ui.notifInboxTypeFilter || "all";

  // Sync filter controls UI
  const unreadEl = document.getElementById("notifUnreadOnly");
  if (unreadEl) unreadEl.checked = unreadOnly;


  const allList = getNotificationsForCurrentUser();
  const visibleBase = allList.filter(n => !n.archived);


  // Apply Saved View filters (before tabs/filters)
  const view = state.ui.notifInboxView || "default";
  const now = Date.now();
  const dayStart = new Date(); dayStart.setHours(0,0,0,0);
  const weekAgo = now - 7*24*60*60*1000;
  const curPid = state.currentProjectId || null;

  let viewBase = visibleBase.slice();
  if (view === "my_actions") {
    viewBase = viewBase.filter(n => n.requiresAction && !n.resolved);
  } else if (view === "current_project") {
    viewBase = curPid ? viewBase.filter(n => (n.projectId || null) === curPid) : [];
  } else if (view === "today") {
    viewBase = viewBase.filter(n => ((n.ts ?? n.createdAt) || 0) >= dayStart.getTime());
  } else if (view === "week") {
    viewBase = viewBase.filter(n => ((n.ts ?? n.createdAt) || 0) >= weekAgo);
  }


  const counts = {
    all: visibleBase.length,
    action: visibleBase.filter(n => n.requiresAction && !n.resolved).length,
    system: visibleBase.filter(n => notifKind(n) === "system" && !(n.requiresAction && !n.resolved)).length,
    warning: visibleBase.filter(n => notifKind(n) === "warning").length,
    admin: visibleBase.filter(n => notifKind(n) === "admin").length,
    unread: visibleBase.filter(n => !n.read).length,
    archived: allList.filter(n => n.archived).length,
  };


  // Render Saved Views
  const viewsBox = document.getElementById("notifSavedViews");
  if (viewsBox) {
    const myActionsCount = visibleBase.filter(n => n.requiresAction && !n.resolved).length;
    const curCount = (curPid ? visibleBase.filter(n => (n.projectId||null) === curPid).length : 0);
    const todayCount = visibleBase.filter(n => (((n.ts ?? n.createdAt) || 0) >= dayStart.getTime())).length;
    const weekCount = visibleBase.filter(n => (((n.ts ?? n.createdAt) || 0) >= weekAgo)).length;

    const defs = [
      { id:"default", name:"All Inbox", hint:"Everything (not archived)", count: counts.all, icon:'📥' },
      { id:"my_actions", name:"My Actions", hint:(myActionsCount===0 ? "No actions assigned to the current user. Use the Acting as switcher (Demo) to simulate roles." : "Pending approvals & tasks"), count: myActionsCount, icon:'✅' },
      { id:"current_project", name:"Current Project", hint: curPid ? "Filtered by opened project" : "Open a project first", count: curCount, disabled: !curPid, icon:'📁' },
      { id:"today", name:"Today", hint:"Notifications created today", count: todayCount, icon:'🗓️' },
      { id:"week", name:"This Week", hint:"Last 7 days", count: weekCount, icon:'📆' },
    ];

    viewsBox.innerHTML = defs.map(v => {
      const active = (state.ui.notifInboxView || "default") === v.id;
      const dis = v.disabled ? 'opacity:.55;pointer-events:none;' : '';
      return `
        <div class="saved-view-btn ${active?'active':''}" style="${dis}" onclick="setNotifSavedView('${v.id}')">
          <div class="left">
            <div class="sv-ico" title="${esc(v.name)}">${esc(v.icon || "")}</div>
            <div>
              <div class="name">${esc(v.name)}</div>
              <div class="hint">${esc(v.hint || "")}</div>
            </div>
          </div>
          <div class="count">${v.count}</div>
        </div>
      `;
    }).join("");
  }


// Populate project dropdown (stable, local)
const projSel = document.getElementById("notifProjectFilter");
if (projSel) {
  const projects = Array.isArray(state.projects) ? state.projects : [];
  projSel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "all";
  optAll.textContent = "All projects";
  projSel.appendChild(optAll);
  projects.forEach(p=>{
    const o = document.createElement("option");
    o.value = p.id;
    o.textContent = p.name || p.id;
    projSel.appendChild(o);
  });
  projSel.value = projectFilter;
}
const typeSel = document.getElementById("notifTypeFilter");
if (typeSel) typeSel.value = typeFilter;

  const filterItems = [
    { key:"all", label:"All", pill: counts.all },
    { key:"action", label:"Action Required", pill: counts.action },
    { key:"system", label:"System", pill: counts.system },
    { key:"warning", label:"Warnings", pill: counts.warning },
    { key:"admin", label:"Admin Notes", pill: counts.admin },
    { key:"archived", label:"Archived", pill: counts.archived },
  ];

  filtersBox.innerHTML = filterItems.map(it => {
    const k = it.key;
    const kind = (k === "action" || k === "system" || k === "warning" || k === "admin") ? k : null;
    const dot = kind ? `<span class="filter-dot ${kind}"></span>` : "";
    const ico = kind ? `<svg class="sbts-ico ${kind}" aria-hidden="true"><use href="#${(NOTIF_CATEGORY_META[kind]||NOTIF_CATEGORY_META.system).icon}"></use></svg>` : "";
    return `
    <div class="filter-item ${tab===it.key?'active':''}" data-filter="${it.key}" onclick="setNotifInboxTab('${it.key}')">
      <div class="filter-left">
        ${dot}
        ${ico}
        <span class="filter-pill">${it.pill}</span>
        <span>${esc(it.label)}</span>
      </div>
    </div>
  `;
  }).join("");


  // Base set for this tab
  let list = [];
  if (tab === "archived") list = allList.filter(n => n.archived);
  else list = visibleBase.slice();

  // Apply filters
  if (tab === "action") list = list.filter(n => n.requiresAction && !n.resolved);
  else if (tab === "warning") list = list.filter(n => notifKind(n) === "warning");
  else if (tab === "system") list = list.filter(n => notifKind(n) === "system" && !(n.requiresAction && !n.resolved));
  else if (tab === "admin") list = list.filter(n => notifKind(n) === "admin");

  // Search
  if (q) {
    list = list.filter(n => {
      const ctxParts = getNotifContextPartsForList(n).join(" ");
      const hay = `${n.title||""} ${n.message||""} ${n.projectName||""} ${ctxParts}`.toLowerCase();
      return hay.includes(q);
    });
  }

  // Sort
  list.sort((a,b) => {
    const ta = new Date(a.ts).getTime();
    const tb = new Date(b.ts).getTime();
    return sort === "oldest" ? (ta - tb) : (tb - ta);
  });

  
const selectedId = state.ui.notifInboxSelectedId;

// Build conversation threads (Gmail-style)
const threadMap = new Map();
list.forEach(n=>{
  const k = getNotifThreadKey(n);
  if(!threadMap.has(k)) threadMap.set(k, []);
  threadMap.get(k).push(n);
});

const threadRows = Array.from(threadMap.values()).map(items => buildThreadSummary(items))
  .filter(t => t.top)
  .sort((a,b)=> new Date(b.top.ts).getTime() - new Date(a.top.ts).getTime());

// Expose visible ids for keyboard navigation (J/K)
window.__SBTS_NOTIF_VISIBLE_IDS = threadRows.map(t => t.top.id);

// Render list
if (threadRows.length === 0) {
  listBox.innerHTML = `<div class="muted" style="padding:12px;">No notifications.</div>`;
} else {
  listBox.innerHTML = threadRows.map(t => {
    const n = t.top;
    const kind = notifKind(n);
    const icon = notifIconSVG(kind);
    const badgeText = (kind === "action") ? "Action" : (kind === "warning" ? "Warning" : (kind === "admin" ? "Admin" : "System"));

    const ctxParts = getNotifContextPartsForList(n);
    const ctx = ctxParts.join(" • ");
    const primaryTitle = buildInboxPrimaryTitle(n);
    const isSelected = selectedId === n.id;
    const isChecked = !!state.ui.notifInboxSelectedIds[n.id];

    const threadPill = (t.count > 1) ? `<span class="thread-pill" title="${t.count} messages">${t.count}</span>` : ``;
    const unreadPill = (t.unread > 0) ? `<span class="thread-pill unread" title="${t.unread} unread">${t.unread}</span>` : ``;

    return `
      <div class="notif-item ${n.read?'':'unread'} ${isSelected?'selected':''}" data-id="${esc(n.id)}" onclick="openNotificationInInbox('${n.id}')" >
        <div>
          <label class="chk" onclick="event.stopPropagation();">
            <input type="checkbox" data-notif-id="${n.id}" ${isChecked?'checked':''} onchange="toggleNotificationSelected('${n.id}', this.checked)" />
            <span></span>
          </label>
        </div>
        <div class="notif-icon ${kind}">${icon}</div>
        <div class="notif-meta">
          <div class="notif-title-row">
            <div class="notif-title" onclick="event.stopPropagation(); openNotifPreviewFromTitle('${n.id}')">
              ${esc(primaryTitle || "Notification")}
              ${threadPill}${unreadPill}
            </div>
            <div class="notif-title-actions">
              <button class="icon-btn mini" title="Open target" onclick="event.stopPropagation(); openNotifTargetFromList('${n.id}')">${notifOpenIconSVG()}</button>
              ${(!n.archived && n.requiresAction && !n.resolved) ? `<button class="icon-btn mini" title="Done" onclick="event.stopPropagation(); quickDoneNotification('${n.id}')">${notifDoneIconSVG()}</button>` : ``}
              ${(!n.archived) ? `<button class="icon-btn mini" title="Archive" onclick="event.stopPropagation(); archiveNotification('${n.id}')">${notifArchiveIconSVG()}</button>` : ``}
              ${n.archived ? `<button class="icon-btn mini" title="Restore" onclick="event.stopPropagation(); restoreNotification('${n.id}')">${notifRestoreIconSVG()}</button>` : ``}
            </div>
          </div>
          <div class="notif-sub">${esc(ctx || "")}${(ctx && (n.message||"")) ? ` <span class="muted">—</span> ${esc((n.message||"").replace(/<[^>]*>/g,"").slice(0,140))}` : (ctx ? "" : esc((n.message||"").replace(/<[^>]*>/g,"").slice(0,180)))}</div>
          <div class="notif-foot">
            <span class="badge ${kind}">${badgeText}</span>
            ${(!n.read ? '<span class="badge">Unread</span>' : '')}
            ${(n.requiresAction && !n.resolved) ? '<span class="badge action">Pending</span>' : ''}
            <span class="badge">${fmtNotifTime(n.ts)}</span>
          </div>
        </div>
      </div>
    `;
  }).join("");
}

// Maintain checkbox state after render

  updateNotifSelectionHint();

  // Render details (Lazy preview like Gmail)
  state.ui.notifInboxPreviewId = state.ui.notifInboxPreviewId || null;
  const previewId = state.ui.notifInboxPreviewId;
  const selected = previewId ? (allList.find(x => x.id === previewId) || null) : null;
  renderNotifDetails(selected, "notifDetails", "panel");
}



function updateNotificationsBadge() {
  const badge = document.getElementById("notifBadge");
  const infoDot = document.getElementById("notifInfoDot");
  const btn = document.getElementById("notifBellBtn");
  if (!badge || !btn) return;

  const list = getNotificationsForCurrentUser().filter(n => !n.archived);
  const action = list.filter(n => n.requiresAction && !n.resolved).length;
  const info = list.filter(n => !n.requiresAction && !n.read).length;
  const attention = action + info;

  badge.textContent = attention;
  badge.title = `Action: ${action} • Info: ${info}`;
  btn.title = badge.title;

  badge.classList.toggle('notif-badge-action', action > 0);
  badge.classList.toggle('notif-badge-info', action === 0 && info > 0);

  if (attention > 0) badge.classList.remove("hidden");
  else badge.classList.add("hidden");

  if (infoDot) {
    infoDot.classList.toggle("hidden", !(info > 0));
  }
}

// Seed realistic demo items (admin can clear via Settings > Clear data, or...)

function seedDemoInbox() {
  try {
    const actingId = sbtsDemoGetCurrentActingUserId();
    const ok = (SBTS_ACTIVITY.seedLiveScenariosForActing ? SBTS_ACTIVITY.seedLiveScenariosForActing(actingId) : SBTS_ACTIVITY.seedLiveScenarios()); // role-aware seed
    let msg = "Demo scenarios added.";
    if(ok && typeof ok === "object" && ok.roleId){
      msg = `Demo added for role: ${ok.roleId} (run ${ok.run||""}).`;
    }
    showToast(msg, "success");

    try { updateNotificationsBadge(); } catch(e) {}
    try { renderNotificationsInbox(); } catch (e) { console.warn("inbox", e); }
    try { renderInboxPage(); } catch (e) { console.warn("inboxPage", e); }
    try { renderNotificationsPage(); } catch (e) { console.warn("page", e); }
    try { renderNotificationsDrawer(); } catch (e) { console.warn("drawer", e); }
  } catch (e) {
    console.error("[seedDemoInbox] failed", e);
    showToast("Demo failed: " + (e && e.message ? e.message : "Unknown error"), "error");
  }
}


/* ==========================
   PATCH43.1 - ADMIN NOTES + SNOOZE + AUTOMATIC TRIGGERS
========================== */

function createAdminNotePrompt(){
  if (!state.currentUser || state.currentUser.role !== "admin") return alert("Admins only.");
  const title = prompt("Admin note title:", "Admin note");
  if (!title) return;
  const message = prompt("Message:", "");
  if (message === null) return;

  const scope = prompt("Send to: all / current_project / role:<role>\nExample: role:qc", "current_project");
  if (!scope) return;

  let recipients = [];
  const users = Array.isArray(state.users) ? state.users : [];

  if (scope === "all") {
    recipients = users.map(u=>u.id).filter(Boolean);
  } else if (scope === "current_project") {
    const pid = state.currentProjectId;
    const p = pid ? (state.projects||[]).find(x=>x.id===pid) : null;
    const owners = [];
    if (p?.phaseOwners?.phases) {
      Object.values(p.phaseOwners.phases).forEach(arr=>{
        (arr||[]).forEach(uid=> owners.push(uid));
      });
    }
    const admins = users.filter(u=>u.role==="admin").map(u=>u.id);
    recipients = Array.from(new Set([...owners, ...admins])).filter(Boolean);
  } else if (scope.startsWith("role:")) {
    const role = scope.slice(5).trim();
    recipients = users.filter(u=>String(u.role||"").toLowerCase()===role.toLowerCase()).map(u=>u.id).filter(Boolean);
  }

  if (recipients.length === 0) {
    // fallback: at least notify self
    recipients = [state.currentUser.id].filter(Boolean);
  }

  addNotification(recipients, {
    type: "admin",
    title,
    message,
    requiresAction: false,
    createdBy: state.currentUser.fullName || state.currentUser.username || "Admin",
    projectId: (scope === "current_project" ? state.currentProjectId : null),
  });

  toast("Admin note sent.");
  updateNotificationsBadge();
}

/* ==========================
   Standalone Notification page
========================== */
function renderNotificationStandalonePage(){
  const mount = document.getElementById("notificationStandalone");
  if (!mount) return;

  ensureNotificationsState();
  const params = new URLSearchParams(window.location.search || "");
  const id = params.get("id");
  const n = id ? getNotificationByIdForCurrentUser(id) : null;

  const box = document.getElementById("notifStandaloneDetails");
  if (box) {
    box.innerHTML = `<div id="__standaloneDetails"></div>`;
    renderNotifDetails(n, "__standaloneDetails");
  }

  const backBtn = document.getElementById("notifStandaloneBack");
  if (backBtn) {
    backBtn.onclick = () => {
      // try to go back, otherwise go to inbox page
      if (history.length > 1) history.back();
      else { navigateToPage('notificationsPage'); }
    };
  }
}

// Try render standalone page on load