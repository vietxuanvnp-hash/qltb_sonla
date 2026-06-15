// ─── Cấu hình từ github-config.js ────────────────────────────────────────────
// GITHUB_CONFIG được load từ github-config.js (phải khai báo trước script này)

const GITHUB_API = "https://api.github.com";

// ─── Schema thiết bị (hardcoded từ biểu mẫu Ra_Soat_Thiet_Bi_CNTT.xlsx) ─────

// Trường bắt buộc – Mã vật tư / Năm sản xuất / Office / Tình trạng là KHÔNG bắt buộc
const COMMON_REQUIRED = [
  "Mã Bưu cục", "Tên Bưu cục",
  "Họ và tên người sử dụng",
  "Mã HRM",
  "Bộ phận / Phòng ban",
  "Tên tài sản\n(Theo danh mục CCDC)",
];

const COMPUTER_FORM_EXCLUDED = new Set([
  "Mã Tỉnh", "Tên Tỉnh",
  "Số thẻ\n(Số mã trên phiếu kê tài sản)",
  "Tên tài sản\n(Theo danh mục CCDC)",
  "Nước sản xuất",
  "Đơn vị bảo hành\n(Nhà cung cấp)",
  "Năm sử dụng", "Thời gian bảo hành (Còn/Hết)", "Tình trạng",
]);

const ATTACHMENT_TYPES = [
  { id: "laser_printer",   name: "Máy in laser",          category: "4A. Máy In Laser" },
  { id: "thermal_printer", name: "Máy in nhiệt",          category: "4B. Máy In Nhiệt" },
  { id: "photocopier",     name: "Máy photocopy",         category: "4C. Máy Photocopy" },
  { id: "barcode_reader",  name: "Đầu đọc mã vạch",      category: "5. Đầu Đọc Mã Vạch" },
  { id: "scanner",         name: "Máy quét (Scanner)",    category: "8. Thiết Bị Khác" },
  { id: "ups",             name: "Bộ lưu điện (UPS)",     category: "8. Thiết Bị Khác" },
];

const EXCLUDED_KEYWORDS = ["server", "máy chủ", "nas", "san"];

function buildSchema() {
  const allSheets = [
    {
      id: "3. Máy Tính", name: "Máy Tính",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Tên máy (Hostname)","Loại máy","Hãng","Model","Năm sản xuất","Serial Number","CPU","RAM","Ổ cứng","Hệ điều hành","Office","Antivirus","IP Address","MAC Address","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)","Tình trạng"],
    },
    {
      id: "4A. Máy In Laser", name: "Máy In Laser",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Khổ giấy","Hãng","Model","Năm sản xuất","Serial Number","Kết nối","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)"],
    },
    {
      id: "4B. Máy In Nhiệt", name: "Máy In Nhiệt",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Khổ giấy","Hãng","Model","Năm sản xuất","Serial Number","Kết nối","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)"],
    },
    {
      id: "4C. Máy Photocopy", name: "Máy Photocopy",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Hình thức sử dụng","Khổ giấy tối đa","Chức năng (In/Scan/Copy)","Số Counter hiện tại\n(Số bản đã in)","Hãng","Model","Năm sản xuất","Serial Number","Kết nối","IP Address","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)"],
    },
    {
      id: "5. Đầu Đọc Mã Vạch", name: "Đầu Đọc Mã Vạch",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Hãng","Model","Năm sản xuất","Serial Number","Loại đọc","Kết nối","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)"],
    },
    {
      id: "6. Thiết Bị Mạng", name: "Thiết Bị Mạng",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Loại thiết bị mạng","Hãng","Model","Năm sản xuất","Serial Number","Cấu hình chi tiết (Số cổng/Wi-Fi...)","Phiên bản Firmware/OS","IP Quản trị","Vị trí lắp đặt","Trạng thái giám sát (Zabbix/ManageEngine...)","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)"],
    },
    {
      id: "7. Cân Điện Tử", name: "Cân Điện Tử",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Loại cân (Cân bàn/Cân bưu kiện...)","Hãng sản xuất","Model","Tải trọng tối đa (vd: 30kg, 100kg)","Serial Number","Năm sản xuất","Ngày kiểm định gần nhất","Thời hạn kiểm định tiếp theo","Vị trí lắp đặt/Sử dụng","Tình trạng hoạt động","Ghi chú"],
    },
    {
      id: "8. Thiết Bị Khác", name: "Thiết Bị Khác",
      fields: ["STT","Mã Tỉnh","Tên Tỉnh","Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục","Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban","Mã vật tư","Số thẻ\n(Số mã trên phiếu kê tài sản)","Tên tài sản\n(Theo danh mục CCDC)","Loại thiết bị","Hãng","Model","Năm sản xuất","Serial Number","Vị trí sử dụng","Nước sản xuất","Đơn vị bảo hành\n(Nhà cung cấp)","Năm sử dụng","Thời gian bảo hành (Còn/Hết)","Tình trạng","Ghi chú"],
    },
  ];

  return allSheets.map((sheet) => {
    const fields = sheet.fields.filter((f) => f !== "STT");
    const formFields = sheet.id === "3. Máy Tính"
      ? fields.filter((f) => !COMPUTER_FORM_EXCLUDED.has(f))
      : fields;
    return {
      id: sheet.id,
      name: sheet.name,
      fields,
      formFields,
      required: formFields.filter((f) => COMMON_REQUIRED.includes(f)),
    };
  });
}

// ─── Trạng thái ứng dụng ─────────────────────────────────────────────────────

const state = {
  schema: buildSchema(),
  attachmentTypes: ATTACHMENT_TYPES,
  records: [],
  stats: null,
  selectedId: null,
  isAdmin: false,
  inventorySha: null,   // SHA của inventory.json trên GitHub (cần cho PUT)
};

// ─── DOM refs ────────────────────────────────────────────────────────────────

const category        = document.getElementById("category");
const categoryFilter  = document.getElementById("category-filter");
const dynamicFields   = document.getElementById("dynamic-fields");
const attachmentSection = document.getElementById("attachment-section");
const attachmentOptions = document.getElementById("attachment-options");
const form            = document.getElementById("record-form");
const formStatus      = document.getElementById("form-status");
const recordBody      = document.getElementById("record-body");
const recordCount     = document.getElementById("record-count");
const search          = document.getElementById("search");
const excelFile       = document.getElementById("excel-file");
const importButton    = document.getElementById("btn-import");
const fileName        = document.getElementById("file-name");
const importStatus    = document.getElementById("import-status");
const adminDialog     = document.getElementById("admin-dialog");
const adminPassword   = document.getElementById("admin-password");
const adminLoginStatus = document.getElementById("admin-login-status");

// ─── Helpers ─────────────────────────────────────────────────────────────────

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("vi-VN");
}

function nowIso() {
  return new Date().toISOString().replace("T", "T").slice(0, 19);
}

function uuid() {
  return crypto.randomUUID ? crypto.randomUUID().replace(/-/g, "") :
    Math.random().toString(36).slice(2).padEnd(32, "0").slice(0, 32);
}

function categoryName(id) {
  return state.schema.find((item) => item.id === id)?.name || id;
}

function currentSchema() {
  return state.schema.find((item) => item.id === category.value);
}

function fieldsForCurrentCategory() {
  const s = currentSchema();
  return s?.formFields || s?.fields || [];
}

function isLongField(field) {
  return /cấu hình|chức năng|trạng thái giám sát/i.test(field);
}

function fieldInputType(field) {
  if (/ngày|thời hạn kiểm định/i.test(field)) return "date";
  if (/năm sản xuất|năm sử dụng|counter/i.test(field)) return "number";
  return "text";
}

// ─── GitHub API ───────────────────────────────────────────────────────────────

function getWriteToken() {
  // Thứ tự ưu tiên:
  // 1. Admin session (sessionStorage) – đăng nhập quản trị
  // 2. localStorage (thiết lập thủ công trước đó)
  // 3. Token inject từ GitHub Actions Secret (tự động cho mọi client)
  return sessionStorage.getItem("gh_admin_token")
    || localStorage.getItem("gh_write_token")
    || (typeof GITHUB_CONFIG !== "undefined" && GITHUB_CONFIG.writeToken)
    || "";
}

function getAdminToken() {
  return sessionStorage.getItem("gh_admin_token") || "";
}

function getWriteTokenLabel() {
  if (sessionStorage.getItem("gh_admin_token")) return "admin";
  if (localStorage.getItem("gh_write_token")) return "write";
  return null;
}

async function githubGet(path, token) {
  const headers = { Accept: "application/vnd.github+json" };
  if (token) headers["Authorization"] = `Bearer ${token}`;
  const res = await fetch(`${GITHUB_API}${path}`, { headers });
  return res;
}

// Giải mã base64 UTF-8 đúng cách – atob() chỉ xử lý Latin-1, không xử lý tiếng Việt!
function decodeBase64UTF8(b64) {
  const binary = atob(b64.replace(/\n/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new TextDecoder("utf-8").decode(bytes);
}

async function loadInventoryFromGitHub() {
  const token = getWriteToken();

  // Nếu có token: dùng Contents API (không bị cache CDN, luôn có dữ liệu mới nhất)
  if (token && token !== "PASTE_YOUR_FINE_GRAINED_PAT_HERE") {
    try {
      const res = await githubGet(
        `/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/contents/${GITHUB_CONFIG.dataFile}`,
        token
      );
      if (res.ok) {
        const data = await res.json();
        // Dùng decodeBase64UTF8 thay vì atob() để xử lý đúng tiếng Việt
        const decoded = decodeBase64UTF8(data.content);
        return JSON.parse(decoded);
      }
    } catch { /* fall through to raw URL */ }
  }

  // Không có token (public read): dùng raw URL với cache-bust
  const rawUrl = `https://raw.githubusercontent.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}/${GITHUB_CONFIG.dataFile}?t=${Date.now()}`;
  try {
    const res = await fetch(rawUrl);
    if (!res.ok) return { records: [] };
    return await res.json();
  } catch {
    return { records: [] };
  }
}

async function getInventorySha(token) {
  const res = await githubGet(
    `/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/contents/${GITHUB_CONFIG.dataFile}`,
    token
  );
  if (!res.ok) return null;
  const data = await res.json();
  return data.sha || null;
}

async function saveInventoryToGitHub(records, commitMessage) {
  const token = getWriteToken();
  if (!token) {
    throw new Error(
      "Chưa có token ghi dữ liệu. " +
      "Vui lòng liên hệ quản trị viên để thiết lập."
    );
  }

  // Lấy SHA hiện tại (bắt buộc cho GitHub PUT API)
  const sha = await getInventorySha(token);

  const content = JSON.stringify({ records }, null, 2);
  const encoded = btoa(unescape(encodeURIComponent(content)));

  const body = { message: commitMessage, content: encoded, branch: GITHUB_CONFIG.branch };
  if (sha) body.sha = sha;

  const res = await fetch(
    `${GITHUB_API}/repos/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/contents/${GITHUB_CONFIG.dataFile}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    }
  );

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.message || `Lỗi GitHub API: ${res.status}`);
  }
  return res.json();
}

// ─── Render ───────────────────────────────────────────────────────────────────

function renderCategoryOptions() {
  const options = state.schema
    .map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name)}</option>`)
    .join("");
  category.innerHTML = options;
  categoryFilter.innerHTML = `<option value="all">Tất cả nhóm thiết bị</option>${options}`;
}

function renderAttachmentOptions(record = null) {
  const existing = new Map((record?.attachments || []).map((item) => [item.type, item]));
  attachmentOptions.innerHTML = state.attachmentTypes
    .map((item) => {
      const value = existing.get(item.id);
      const checked = Boolean(value);
      return `<div class="attachment-item" data-attachment="${escapeHtml(item.id)}">
        <label class="attachment-toggle">
          <input type="checkbox" data-attachment-toggle="${escapeHtml(item.id)}" ${checked ? "checked" : ""} />
          <span>${escapeHtml(item.name)}</span>
        </label>
        <div class="attachment-fields ${checked ? "" : "is-hidden"}">
          <label>
            <span>Tên thiết bị / Hãng - Model</span>
            <input class="input" data-attachment-name="${escapeHtml(item.id)}" value="${escapeHtml(value?.name || "")}" />
          </label>
          <label>
            <span>Số S/N</span>
            <input class="input" data-attachment-serial="${escapeHtml(item.id)}" value="${escapeHtml(value?.serial || "")}" />
          </label>
        </div>
      </div>`;
    })
    .join("");
}

function renderFields(record = null) {
  const values = record?.fields || {};
  const reqSet = new Set(currentSchema()?.required || []);
  dynamicFields.innerHTML = fieldsForCurrentCategory()
    .map((field) => {
      const id = `field-${btoa(unescape(encodeURIComponent(field))).replace(/=+$/g, "")}`;
      const value = values[field] || "";
      const wide = isLongField(field) ? " is-wide" : "";
      const isReq = reqSet.has(field);
      const reqAttr = isReq ? " required" : "";
      const reqMark = isReq ? ` <span class="req-mark" title="Bắt buộc">*</span>` : "";
      if (isLongField(field)) {
        return `<label class="${wide}${isReq ? " is-required" : ""}">
          <span>${escapeHtml(field)}${reqMark}</span>
          <textarea class="input" name="${escapeHtml(field)}" id="${id}"${reqAttr}>${escapeHtml(value)}</textarea>
        </label>`;
      }
      return `<label class="${wide}${isReq ? " is-required" : ""}">
        <span>${escapeHtml(field)}${reqMark}</span>
        <input class="input" name="${escapeHtml(field)}" id="${id}" type="${fieldInputType(field)}" value="${escapeHtml(value)}"${reqAttr} />
      </label>`;
    })
    .join("");

  const isComputer = category.value === "3. Máy Tính";
  attachmentSection.classList.toggle("is-hidden", !isComputer);
  if (isComputer) renderAttachmentOptions(record);
  else attachmentOptions.innerHTML = "";

  // Cập nhật banner cảnh báo lấy thông tin máy
  updateMachineInfoBanner();
}

function readFormFields() {
  const fields = {};
  for (const field of fieldsForCurrentCategory()) {
    const input = dynamicFields.querySelector(`[name="${CSS.escape(field)}"]`);
    fields[field] = input?.value?.trim() || "";
  }
  return fields;
}

function readAttachments() {
  if (category.value !== "3. Máy Tính") return [];
  return state.attachmentTypes
    .filter((item) => attachmentOptions.querySelector(`[data-attachment-toggle="${item.id}"]`)?.checked)
    .map((item) => ({
      type: item.id,
      name: attachmentOptions.querySelector(`[data-attachment-name="${item.id}"]`)?.value?.trim() || "",
      serial: attachmentOptions.querySelector(`[data-attachment-serial="${item.id}"]`)?.value?.trim() || "",
    }));
}

function recordLabel(record) {
  const fields = record.fields || {};
  return fields["Serial Number"] || fields["Tên máy (Hostname)"] || fields["Số thẻ\n(Số mã trên phiếu kê tài sản)"] || "";
}

function recordStatus(record) {
  return record.fields?.["Tình trạng"] || record.fields?.["Tình trạng hoạt động"] || "";
}

function filteredRecords() {
  const term = search.value.trim().toLowerCase();
  return state.records.filter((record) => {
    if (categoryFilter.value !== "all" && record.category !== categoryFilter.value) return false;
    if (!term) return true;
    return JSON.stringify(record).toLowerCase().includes(term);
  });
}

function renderRecords() {
  const rows = filteredRecords();
  recordCount.textContent = `${formatNumber(rows.length)} phiếu`;
  recordBody.innerHTML = rows
    .map((record) => {
      const fields = record.fields || {};
      const unit = fields["Tên Bưu cục"] || fields["Bộ phận / Phòng ban"] || "";
      const asset = fields["Tên tài sản\n(Theo danh mục CCDC)"] || fields.Model || categoryName(record.category);
      const attachCount = record.attachments?.length ? ` + ${record.attachments.length} thiết bị kèm` : "";
      return `<tr>
        <td><span class="pill">${escapeHtml(categoryName(record.category))}</span></td>
        <td>${escapeHtml(unit)}</td>
        <td>${escapeHtml(asset)}${escapeHtml(attachCount)}</td>
        <td>${escapeHtml(recordLabel(record))}</td>
        <td>${escapeHtml(recordStatus(record))}</td>
        <td>
          <div class="row-actions">
            <button class="ghost" data-edit="${record.id}">Sửa</button>
            <button class="danger" data-delete="${record.id}">Xóa</button>
          </div>
        </td>
      </tr>`;
    })
    .join("");
}

// ─── Stats ────────────────────────────────────────────────────────────────────

function expandedRecords(records) {
  const output = [...records];
  const copiedFields = [
    "Mã BĐ Xã","Tên Xã","Mã Bưu cục","Tên Bưu cục",
    "Họ và tên người sử dụng","Mã HRM","Bộ phận / Phòng ban",
  ];
  for (const record of records) {
    if (record.category !== "3. Máy Tính") continue;
    const source = record.fields || {};
    for (const att of (record.attachments || [])) {
      const fields = Object.fromEntries(copiedFields.map((f) => [f, source[f] || ""]));
      const attTypes = ATTACHMENT_TYPES.find((t) => t.id === att.type);
      fields["Tên tài sản\n(Theo danh mục CCDC)"] = attTypes?.name || "";
      fields["Model"] = att.name || "";
      fields["Serial Number"] = att.serial || "";
      fields["Loại thiết bị"] = attTypes?.name || "";
      output.push({
        id: `${record.id}:${att.type}`,
        category: attTypes?.category || "8. Thiết Bị Khác",
        fields,
        derived: true,
      });
    }
  }
  return output;
}

function computeStats(records) {
  const expanded = expandedRecords(records);
  const schemaById = Object.fromEntries(state.schema.map((s) => [s.id, s]));
  const byCategory = Object.fromEntries(state.schema.map((s) => [s.id, 0]));
  const byStatus = {}, byUnit = {};
  let missingRequired = 0, excludedCandidates = 0, monitoringYes = 0, monitoringTotal = 0;

  for (const record of expanded) {
    const cat = record.category || "";
    const fields = record.fields || {};
    byCategory[cat] = (byCategory[cat] || 0) + 1;

    const status = fields["Tình trạng"] || fields["Tình trạng hoạt động"] || "Chưa nhập tình trạng";
    byStatus[status] = (byStatus[status] || 0) + 1;

    const unit = fields["Tên Bưu cục"] || fields["Bộ phận / Phòng ban"] || "Chưa xác định";
    byUnit[unit] = (byUnit[unit] || 0) + 1;

    const schema = schemaById[cat];
    if (!record.derived && schema && schema.required.some((f) => !fields[f])) missingRequired++;

    const searchable = Object.values(fields).join(" ").toLowerCase();
    if (EXCLUDED_KEYWORDS.some((kw) => searchable.includes(kw))) excludedCandidates++;

    const monitor = fields["Trạng thái giám sát (Zabbix/ManageEngine...)"];
    if (monitor) {
      monitoringTotal++;
      if (/có|zabbix|manage/i.test(monitor)) monitoringYes++;
    }
  }

  // Lấy top 12 đơn vị
  const topUnit = Object.fromEntries(
    Object.entries(byUnit).sort((a, b) => b[1] - a[1]).slice(0, 12)
  );

  return {
    total: expanded.length, surveyForms: records.length,
    attachedDevices: expanded.length - records.length,
    byCategory, byStatus, byUnit: topUnit,
    missingRequired, excludedCandidates, monitoringYes, monitoringTotal,
  };
}

function renderKpis() {
  const stats = state.stats || {};
  const rate = stats.monitoringTotal ? Math.round((stats.monitoringYes / stats.monitoringTotal) * 100) : 0;
  const values = [
    ["Tổng thiết bị", stats.total || 0, ""],
    ["Phiếu khảo sát", stats.surveyForms || 0, ""],
    ["Thiết bị dùng kèm", stats.attachedDevices || 0, ""],
    ["Thiếu trường bắt buộc", stats.missingRequired || 0, stats.missingRequired ? "warn" : ""],
    ["Cần loại khỏi phạm vi", stats.excludedCandidates || 0, stats.excludedCandidates ? "bad" : ""],
    ["Thiết bị mạng có giám sát", `${rate}%`, ""],
  ];
  document.getElementById("kpis").innerHTML = values
    .map(([label, value, tone]) =>
      `<div class="kpi ${tone}"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`
    ).join("");
}

function renderBars(targetId, values) {
  const entries = Object.entries(values || {}).filter(([, v]) => Number(v) > 0);
  const max = Math.max(1, ...entries.map(([, v]) => Number(v)));
  document.getElementById(targetId).innerHTML = entries.length
    ? entries.map(([label, value]) =>
        `<div class="bar-row">
          <div class="bar-label" title="${escapeHtml(categoryName(label) || label)}">${escapeHtml(categoryName(label) || label)}</div>
          <strong>${formatNumber(value)}</strong>
          <div class="bar-track"><div class="bar-fill" style="width:${Math.max(4, (Number(value) / max) * 100)}%"></div></div>
        </div>`
      ).join("")
    : `<p class="status">Chưa có dữ liệu.</p>`;
}

function renderStats() {
  renderKpis();
  renderBars("category-bars", state.stats?.byCategory);
  renderBars("unit-bars", state.stats?.byUnit);
  renderBars("status-bars", state.stats?.byStatus);
}

// ─── Admin UI ─────────────────────────────────────────────────────────────────

function setAdminUi(authenticated) {
  state.isAdmin = authenticated;
  document.querySelectorAll(".admin-only").forEach((el) =>
    el.classList.toggle("is-hidden", !authenticated)
  );
  document.getElementById("btn-admin-login").classList.toggle("is-hidden", authenticated);
  document.getElementById("btn-admin-logout").classList.toggle("is-hidden", !authenticated);
  if (!authenticated) {
    state.records = [];
    state.stats = null;
    state.selectedId = null;
  }
}

// ─── Tải & lưu dữ liệu ───────────────────────────────────────────────────────

async function loadRecords() {
  const payload = await loadInventoryFromGitHub();
  state.records = payload.records || [];
  state.stats = computeStats(state.records);
  renderRecords();
  renderStats();
}

function resetForm({ keepStatus = false } = {}) {
  state.selectedId = null;
  form.reset();
  category.value = state.schema[0]?.id || "";
  renderFields();
  if (!keepStatus) formStatus.textContent = "";
}

async function saveRecord(event) {
  event.preventDefault();

  // ── Kiểm tra trường bắt buộc trước khi gửi ──────────────────────────────
  const schema = currentSchema();
  const formValues = readFormFields();
  const missing = (schema?.required || []).filter((f) => !formValues[f]?.trim());

  // Xóa highlight cũ
  dynamicFields.querySelectorAll(".input-error").forEach((el) =>
    el.classList.remove("input-error")
  );

  if (missing.length > 0) {
    // Highlight từng trường còn trống
    missing.forEach((field) => {
      const el = dynamicFields.querySelector(`[name="${CSS.escape(field)}"]`);
      if (el) el.classList.add("input-error");
    });
    // Cuộn đến trường đầu tiên còn trống
    const firstEl = dynamicFields.querySelector(".input-error");
    firstEl?.scrollIntoView({ behavior: "smooth", block: "center" });

    const labels = missing.map((f) => f.split("\n")[0]);
    formStatus.textContent = `Vui lòng điền đầy đủ trường bắt buộc: ${labels.join(", ")}`;
    formStatus.className = "status form-status is-error";
    return;
  }

  // Kiểm tra bổ sung: Nhóm Máy Tính phải lấy thông tin phần cứng trước
  if (category.value === "3. Máy Tính" && !formValues["Tên máy (Hostname)"]?.trim()) {
    const banner = document.getElementById("machine-info-banner");
    if (banner) {
      banner.classList.remove("is-hidden");
      banner.scrollIntoView({ behavior: "smooth", block: "center" });
    }
    formStatus.textContent = "Vui lòng lấy thông tin máy tính trước khi gửi phiếu. Xem hướng dẫn phía trên.";
    formStatus.className = "status form-status is-error";
    return;
  }

  formStatus.textContent = "Đang lưu phiếu lên GitHub...";
  formStatus.className = "status form-status";


  try {
    // Tải bản hiện tại từ GitHub để tránh ghi đè mất dữ liệu
    const payload = await loadInventoryFromGitHub();
    const records = payload.records || [];

    const existing = records.find((r) => r.id === state.selectedId);
    const now = nowIso();

    if (existing && state.isAdmin) {
      // Cập nhật phiếu đã chọn
      Object.assign(existing, {
        category: category.value,
        fields: readFormFields(),
        attachments: readAttachments(),
        updatedAt: now,
      });
    } else {
      // Thêm phiếu mới
      records.push({
        id: uuid(),
        category: category.value,
        fields: readFormFields(),
        attachments: readAttachments(),
        createdAt: now,
        updatedAt: now,
      });
    }

    await saveInventoryToGitHub(records, `Cập nhật phiếu thiết bị – ${now}`);

    // Cập nhật state trực tiếp – không cần đọc lại từ GitHub (tránh cache CDN)
    state.records = records;
    state.stats = computeStats(records);
    renderRecords();
    renderStats();

    resetForm({ keepStatus: true });
    formStatus.textContent = "Đã gửi phiếu thành công lên GitHub.";
    formStatus.className = "status form-status is-success";
  } catch (err) {
    formStatus.textContent = err.message || "Không lưu được phiếu.";
    formStatus.className = "status form-status is-error";
  }
}

async function deleteRecord(id) {
  try {
    const payload = await loadInventoryFromGitHub();
    const records = (payload.records || []).filter((r) => r.id !== id);
    await saveInventoryToGitHub(records, `Xóa phiếu ${id}`);
    if (state.selectedId === id) resetForm();
    setTimeout(loadRecords, 1500);
  } catch (err) {
    alert("Không xóa được phiếu: " + err.message);
  }
}

function editRecord(id) {
  const record = state.records.find((item) => item.id === id);
  if (!record || !state.isAdmin) return;
  state.selectedId = id;
  category.value = record.category;
  renderFields(record);
  formStatus.textContent = "Đang sửa phiếu đã chọn.";
  formStatus.className = "status form-status";
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// ─── Admin Auth (GitHub PAT) ──────────────────────────────────────────────────

async function loginAdmin(event) {
  event.preventDefault();
  const token = adminPassword.value.trim();
  if (!token) return;

  adminLoginStatus.textContent = "Đang xác thực token GitHub...";
  try {
    // Xác minh token bằng cách gọi /user API
    const res = await fetch(`${GITHUB_API}/user`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/vnd.github+json" },
    });
    if (!res.ok) throw new Error("Token không hợp lệ hoặc đã hết hạn.");

    const user = await res.json();
    sessionStorage.setItem("gh_admin_token", token);
    sessionStorage.setItem("gh_admin_user", user.login || "");

    // Lưu luôn vào localStorage để dùng chung cho việc gửi phiếu (write token)
    const saveAsWrite = document.getElementById("chk-save-write-token")?.checked;
    if (saveAsWrite) {
      localStorage.setItem("gh_write_token", token);
    }

    adminPassword.value = "";
    adminLoginStatus.textContent = "";
    adminDialog.close();
    setAdminUi(true);
    await loadRecords();
  } catch (err) {
    adminLoginStatus.textContent = err.message || "Xác thực thất bại.";
  }
}

function exportExcel() {
  if (!state.isAdmin) return;
  const XLSX = window.XLSX;
  if (!XLSX) { alert("Thư viện SheetJS chưa được tải."); return; }
  if (!state.records || state.records.length === 0) {
    alert("Không có dữ liệu để xuất."); return;
  }

  const wb = XLSX.utils.book_new();
  const expanded = expandedRecords(state.records);

  // Map category id → schema (cộ định hóa bằng trim() để tránh lỗi khoảng trắng)
  const schemaMap = new Map(state.schema.map((s) => [s.id.trim(), s]));

  const grouped = new Map(state.schema.map((s) => [s.id.trim(), []]));
  for (const rec of expanded) {
    const cat = (rec.category || "").trim();
    if (grouped.has(cat)) {
      grouped.get(cat).push(rec);
    }
  }

  let sheetCount = 0;
  for (const schema of state.schema) {
    const rows = grouped.get(schema.id.trim()) || [];
    const headers = (schema.fields || []).filter((f) => f !== "STT");
    const data = [["STT", ...headers]];
    rows.forEach((rec, idx) => {
      data.push([idx + 1, ...headers.map((h) => rec.fields?.[h] ?? "")]);
    });
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Tên sheet Excel giới hạn 31 ký tự, loại bỏ ký tự không hợp lệ
    const sheetName = schema.id.replace(/[\[\]\*\/\\\?\:]/g, "").slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, sheetName || `Sheet${++sheetCount}`);
  }

  const ts = new Date().toISOString().slice(0, 16).replace(/[-T:]/g, "");
  XLSX.writeFile(wb, `Ra_Soat_Thiet_Bi_CNTT_${ts}.xlsx`);
}

async function logoutAdmin() {
  sessionStorage.removeItem("gh_admin_token");
  sessionStorage.removeItem("gh_admin_user");
  resetForm();
  setAdminUi(false);
}

// ─── Lấy thông tin máy từ URL params (lay_thong_tin.bat) ─────────────────────

const URL_PARAM_MAP = {
  hostname: "Tên máy (Hostname)",
  cpu:      "CPU",
  hang:     "Hãng",
  model:    "Model",
  ram:      "RAM",
  disk:     "Ổ cứng",
  serial:   "Serial Number",
  os:       "Hệ điều hành",
  ip:       "IP Address",
  mac:      "MAC Address",
  loaiMay:  "Loại máy",
  office:   "Office",
  antivirus: "Antivirus",
};

function applyClientInventory(fields) {
  if (category.value !== "3. Máy Tính") {
    category.value = "3. Máy Tính";
    renderFields();
  }
  for (const [field, value] of Object.entries(fields || {})) {
    const input = dynamicFields.querySelector(`[name="${CSS.escape(field)}"]`);
    if (input && value) input.value = value;
  }
  updateMachineInfoBanner();
}

function loadClientInfoFromUrlParams() {
  const params = new URLSearchParams(window.location.search);
  const fields = {};
  let found = false;
  for (const [param, fieldName] of Object.entries(URL_PARAM_MAP)) {
    const val = params.get(param);
    if (val) { fields[fieldName] = decodeURIComponent(val); found = true; }
  }
  if (!found) return false;

  applyClientInventory(fields);
  // Xóa params khỏi URL mà không reload
  const clean = `${window.location.pathname}`;
  window.history.replaceState({}, "", clean);
  formStatus.textContent = "Đã tự động điền thông tin máy tính. Hãy nhập bổ sung thông tin người sử dụng và gửi phiếu.";
  formStatus.className = "status form-status is-success";
  updateMachineInfoBanner();
  return true;
}

// ─── Banner cảnh báo chưa lấy thông tin máy ──────────────────────────────────

function updateMachineInfoBanner() {
  const banner = document.getElementById("machine-info-banner");
  if (!banner) return;

  // Chỉ hiển thị khi nhóm là Máy Tính
  if (category.value !== "3. Máy Tính") {
    banner.classList.add("is-hidden");
    return;
  }

  // Hostname trống = chưa chạy công cụ lấy thông tin
  const hostnameInput = dynamicFields.querySelector('[name="Tên máy (Hostname)"]');
  const hasInfo = hostnameInput && hostnameInput.value.trim();
  banner.classList.toggle("is-hidden", Boolean(hasInfo));
}

// ─── Tải công cụ lấy thông tin máy ───────────────────────────────────────────
// Dùng chung cho cả nút "Làm mới thông tin máy" và nút trong banner

async function triggerBatDownload() {
  // Thử giao thức vnpost:// (nếu đã đăng ký thì bat chạy tự động)
  const iframe = document.createElement("iframe");
  iframe.style.display = "none";
  iframe.src = "vnpost://collect";
  document.body.appendChild(iframe);
  setTimeout(() => document.body.removeChild(iframe), 3000);

  formStatus.innerHTML = `
    <div class="auto-setup">
      <strong>🔧 Đang tải công cụ thu thập thông tin...</strong><br>
      <span id="dl-progress">Vui lòng chờ...</span>
    </div>`;
  formStatus.className = "status form-status";

  const RAW_BAT_URL =
    `https://raw.githubusercontent.com/${GITHUB_CONFIG.owner}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}/lay_thong_tin.bat`;

  try {
    const res = await fetch(RAW_BAT_URL);
    if (!res.ok) throw new Error("Không tải được file");
    let text = await res.text();
    
    // Tự động thay thế URL máy chủ trong file .bat bằng URL hiện tại của trang web
    const currentUrl = window.location.origin + window.location.pathname.replace(/\/$/, "");
    text = text.replaceAll("https://nguyennam90.github.io/Check_thiet_bi", currentUrl);

    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const blobUrl = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = "lay_thong_tin.bat";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(blobUrl); }, 2000);

    const dlEl = document.getElementById("dl-progress");
    if (dlEl) dlEl.parentElement.innerHTML = `
      <strong>✅ Tải xuống hoàn tất!</strong><br>
      👇 Nhấn vào file <strong>lay_thong_tin.bat</strong> ở <strong>thanh tải xuống phía dưới trình duyệt</strong> để chạy ngay.<br>
      <small style="opacity:.75">
        • Chrome/Edge: thanh tải xuống xuất hiện ở góc dưới bên phải ▼<br>
        • Nếu Windows hỏi xác nhận → chọn <em>"Run anyway"</em> hoặc <em>"Vẫn chạy"</em><br>
        • Kaspersky: phiên bản mới không còn bị chặn (không dùng iex)
      </small>`;
    formStatus.className = "status form-status is-success";
  } catch (err) {
    const dlEl = document.getElementById("dl-progress");
    if (dlEl) dlEl.textContent = "Không tải được file tự động. Vui lòng thử lại.";
    formStatus.className = "status form-status is-error";
  }
}

// ─── Xuất Excel (SheetJS) ────────────────────────────────────────────────────

function exportExcel() {
  if (!state.isAdmin) return;
  const XLSX = window.XLSX;
  if (!XLSX) { alert("Thư viện SheetJS chưa được tải."); return; }

  const wb = XLSX.utils.book_new();
  const expanded = expandedRecords(state.records);

  const grouped = {};
  for (const s of state.schema) grouped[s.id] = [];
  for (const rec of expanded) {
    if (grouped[rec.category]) grouped[rec.category].push(rec);
  }

  for (const schema of state.schema) {
    const rows = grouped[schema.id] || [];
    const headers = schema.fields.filter((f) => f !== "STT");
    const data = [["STT", ...headers]];
    rows.forEach((rec, idx) => {
      data.push([idx + 1, ...headers.map((h) => rec.fields?.[h] || "")]);
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, schema.id);
  }

  const ts = new Date().toISOString().slice(0, 19).replace(/[-T:]/g, "").slice(0, 15);
  XLSX.writeFile(wb, `Ra_Soat_Thiet_Bi_CNTT_${ts}.xlsx`);
}

// ─── Xuất JSON ────────────────────────────────────────────────────────────────

function exportJson() {
  if (!state.isAdmin) return;
  const blob = new Blob(
    [JSON.stringify({ records: state.records }, null, 2)],
    { type: "application/json" }
  );
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  const ts = new Date().toISOString().slice(0, 19).replace(/[-T:]/g, "").slice(0, 15);
  a.download = `du_lieu_thiet_bi_${ts}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ─── Nhập Excel (SheetJS) ────────────────────────────────────────────────────

async function importExcel() {
  if (!excelFile.files[0] || !state.isAdmin) return;
  const XLSX = window.XLSX;
  if (!XLSX) { alert("Thư viện SheetJS chưa được tải."); return; }

  importButton.disabled = true;
  importStatus.textContent = "Đang đọc file Excel...";

  try {
    const buffer = await excelFile.files[0].arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });

    const schemaById = Object.fromEntries(state.schema.map((s) => [s.id, s]));
    const newRecords = [];

    for (const sheetName of state.schema.map((s) => s.id)) {
      if (!wb.SheetNames.includes(sheetName) || !schemaById[sheetName]) continue;
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (rows.length < 5) continue;

      const headers = rows[3].map((h) => String(h || "").trim());
      for (let i = 4; i < rows.length; i++) {
        const row = rows[i];
        const fields = {};
        let hasData = false;
        headers.forEach((h, idx) => {
          if (h && h !== "STT") {
            const val = String(row[idx] || "").trim();
            fields[h] = val;
            if (val) hasData = true;
          }
        });
        if (!hasData) continue;
        newRecords.push({
          id: uuid(), category: sheetName,
          fields, attachments: [],
          createdAt: nowIso(), updatedAt: nowIso(),
        });
      }
    }

    if (newRecords.length === 0) {
      importStatus.textContent = "Không tìm thấy dữ liệu hợp lệ trong file.";
      importButton.disabled = false;
      return;
    }

    // Merge với dữ liệu hiện có
    const payload = await loadInventoryFromGitHub();
    const existing = payload.records || [];

    // Upsert theo identity (category + serial/hostname/asset)
    function recordKey(r) {
      const f = r.fields || {};
      return [
        r.category,
        f["Số thẻ\n(Số mã trên phiếu kê tài sản)"],
        f["Serial Number"],
        f["Tên máy (Hostname)"],
        f["Tên Bưu cục"],
        f["Tên tài sản\n(Theo danh mục CCDC)"],
      ].map((v) => String(v || "").trim().toLowerCase()).join("|");
    }

    const index = Object.fromEntries(existing.filter(recordKey).map((r) => [recordKey(r), r]));
    let inserted = 0, updated = 0;

    for (const nr of newRecords) {
      const key = recordKey(nr);
      if (key && index[key]) {
        Object.assign(index[key].fields, nr.fields);
        index[key].updatedAt = nowIso();
        updated++;
      } else {
        existing.push(nr);
        if (key) index[key] = nr;
        inserted++;
      }
    }

    importStatus.textContent = "Đang lưu lên GitHub...";
    await saveInventoryToGitHub(existing, `Import Excel: +${inserted} mới, ${updated} cập nhật`);
    importStatus.textContent = `Nhập xong: ${newRecords.length} dòng đọc, thêm ${inserted}, cập nhật ${updated}.`;
    setTimeout(loadRecords, 1500);
  } catch (err) {
    importStatus.textContent = "Lỗi: " + err.message;
  } finally {
    importButton.disabled = false;
  }
}

// ─── Reset dữ liệu ────────────────────────────────────────────────────────────

async function resetData() {
  if (!state.isAdmin) return;
  if (!confirm("Xóa toàn bộ dữ liệu đang lưu trên GitHub?")) return;
  try {
    await saveInventoryToGitHub([], "Reset – xóa toàn bộ dữ liệu");
    resetForm();
    setTimeout(loadRecords, 1500);
  } catch (err) {
    alert("Không reset được: " + err.message);
  }
}

// ─── Event Listeners ──────────────────────────────────────────────────────────

document.getElementById("btn-admin-login").addEventListener("click", () => adminDialog.showModal());
document.getElementById("btn-close-dialog").addEventListener("click", () => adminDialog.close());
document.getElementById("admin-login-form").addEventListener("submit", loginAdmin);
document.getElementById("btn-admin-logout").addEventListener("click", logoutAdmin);
document.getElementById("btn-export-excel").addEventListener("click", exportExcel);
document.getElementById("btn-export-json").addEventListener("click", exportJson);
document.getElementById("btn-reset").addEventListener("click", resetData);

excelFile.addEventListener("change", () => {
  fileName.textContent = excelFile.files[0]?.name || "Chọn file .xlsx";
  importButton.disabled = !excelFile.files[0];
});

attachmentOptions.addEventListener("change", (event) => {
  const type = event.target.dataset.attachmentToggle;
  if (!type) return;
  const container = attachmentOptions.querySelector(`[data-attachment="${type}"] .attachment-fields`);
  container?.classList.toggle("is-hidden", !event.target.checked);
});

importButton.addEventListener("click", importExcel);
category.addEventListener("change", () => { renderFields(); updateMachineInfoBanner(); });
form.addEventListener("submit", saveRecord);
document.getElementById("btn-new").addEventListener("click", () => resetForm());
document.getElementById("btn-local-info").addEventListener("click", () => {
  if (loadClientInfoFromUrlParams()) return;
  triggerBatDownload();
});
document.getElementById("btn-mib-download").addEventListener("click", triggerBatDownload);
search.addEventListener("input", renderRecords);
categoryFilter.addEventListener("change", renderRecords);

recordBody.addEventListener("click", (event) => {
  const editId = event.target.dataset.edit;
  const deleteId = event.target.dataset.delete;
  if (editId) editRecord(editId);
  if (deleteId && confirm("Xóa phiếu đã chọn?")) deleteRecord(deleteId);
});

// ─── Khởi động ────────────────────────────────────────────────────────────────

(async function init() {
  renderCategoryOptions();
  renderFields();

  // Kiểm tra session admin đã lưu
  const savedToken = sessionStorage.getItem("gh_admin_token");
  if (savedToken) {
    // Xác minh token còn hợp lệ
    const res = await fetch(`${GITHUB_API}/user`, {
      headers: { Authorization: `Bearer ${savedToken}`, Accept: "application/vnd.github+json" },
    }).catch(() => null);
    if (res?.ok) {
      setAdminUi(true);
      await loadRecords();
    } else {
      sessionStorage.removeItem("gh_admin_token");
    }
  }

  // Đọc thông tin máy từ URL params (do lay_thong_tin.bat tạo ra)
  loadClientInfoFromUrlParams();
})();
