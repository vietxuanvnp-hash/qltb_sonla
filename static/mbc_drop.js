/* ============================================================
   MBC DROP – Autocomplete Mã Bưu cục + tự điền
   Khi chọn / gõ đúng Mã Bưu cục sẽ tự điền:
     Tên Bưu cục, Mã BĐ Xã, Tên Xã (và Mã Tỉnh, Tên Tỉnh nếu form có)
   Nguồn dữ liệu: static/assets/danh_muc_MBC.xlsx (sheet đầu tiên)
============================================================ */
(function () {
"use strict";

const MBC_EXCEL = "./assets/danh_muc_MBC.xlsx";
const MBC_EXCEL_FALLBACKS = ["assets/danh_muc_MBC.xlsx", "/assets/danh_muc_MBC.xlsx"];
const MAX_RESULT = 20;
const MBC_BUILD = "20261010";   // đổi số này mỗi lần sửa để dễ kiểm tra bản đang chạy

/* Tên trường trong form -> thuộc tính của dòng dữ liệu */
const FIELD_MAP = [
    { names: ["Tên Bưu cục"],                         key: "tenBC"   },
    { names: ["Mã BĐ Xã"],                            key: "maXa"    },
    { names: ["Tên Xã", "Tên xã", "Tên xã/phường"],   key: "tenXa"   },
    { names: ["Mã Tỉnh"],                             key: "maTinh"  },
    { names: ["Tên Tỉnh"],                            key: "tenTinh" }
];

const MBC = {
    rows: [],            // toàn bộ dòng đã chuẩn hóa
    map: new Map(),      // mã bưu cục -> dòng đầu tiên
    duplicates: [],      // các mã bị trùng trong file Excel
    loaded: false,
    loadError: null,
    dropdown: null,
    input: null,
    wrapper: null,
    selectedIndex: -1,
    filtered: [],
    autoFilled: false
};

let outsideClickBound = false;
let submitBound = false;
let typingTimer = null;
let observer = null;

/*==========================================================
    HELPERS
==========================================================*/
function codeKey(v) {
    return cleanText(v).replace(/\s+/g, "").toUpperCase();
}
function cleanText(v) {
    return String(v == null ? "" : v).replace(/\s+/g, " ").trim();
}
/* chuẩn hóa tiêu đề cột: bỏ xuống dòng, khoảng trắng thừa, chữ thường */
function normHeader(v) {
    return cleanText(v).toLowerCase().normalize("NFC");
}
/* bỏ dấu để tìm kiếm không phân biệt dấu */
function fold(v) {
    return String(v == null ? "" : v)
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d").replace(/Đ/g, "D")
        .toLowerCase();
}
function escapeHtml(text) {
    return String(text)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
function highlight(text, keyword) {
    const safe = escapeHtml(text);
    if (!keyword) return safe;
    const reg = new RegExp("(" + escapeHtml(keyword).replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "ig");
    return safe.replace(reg, "<mark>$1</mark>");
}
function fireEvents(el) {
    el.dispatchEvent(new Event("input",  { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
}

/*==========================================================
    LOAD EXCEL  (tự nhận diện cột theo tiêu đề)
==========================================================*/
function detectColumns(headers) {
    const cols = {};
    headers.forEach(h => {
        const n = normHeader(h);
        if (!n) return;
        if (/^mã\s*bưu cục/.test(n))            cols.maBC    = h;
        else if (/^tên\s*bưu cục/.test(n))      cols.tenBC   = h;
        else if (/^mã\s*bđ\s*xã/.test(n))       cols.maXa    = h;
        else if (/^tên\s*(xã|phường)/.test(n))  cols.tenXa   = h;
        else if (/^mã\s*tỉnh/.test(n))          cols.maTinh  = h;
        else if (/^tên\s*(bđt|tỉnh)/.test(n))   cols.tenTinh = h;
    });
    return cols;
}

async function loadExcel() {
    if (MBC.loaded) return;
    if (typeof XLSX === "undefined") {
        throw new Error("Chưa tải được thư viện SheetJS (XLSX).");
    }
    let response = null, lastErr = "";
    for (const url of [MBC_EXCEL].concat(MBC_EXCEL_FALLBACKS)) {
        try {
            const r = await fetch(url + "?v=" + Date.now(), { cache: "no-store" });
            if (r.ok) { response = r; break; }
            lastErr = url + " (HTTP " + r.status + ")";
        } catch (e) { lastErr = url + " (" + e.message + ")"; }
    }
    if (!response) throw new Error("Không đọc được file danh mục: " + lastErr);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    if (!raw.length) throw new Error("File danh mục bưu cục không có dữ liệu.");

    const cols = detectColumns(Object.keys(raw[0]));
    if (!cols.maBC || !cols.tenBC) {
        throw new Error("File danh mục thiếu cột 'Mã Bưu cục' hoặc 'Tên Bưu cục'.");
    }

    MBC.rows = [];
    MBC.map.clear();
    MBC.duplicates = [];

    raw.forEach(r => {
        const ma = cleanText(r[cols.maBC]);
        if (!ma) return;
        let tenTinh = cols.tenTinh ? cleanText(r[cols.tenTinh]) : "";
        tenTinh = tenTinh.replace(/^BĐ(TP|TT|T)\s+/i, "");   // "BĐT Sơn La" -> "Sơn La"
        const row = {
            ma: ma,
            tenBC:   cleanText(r[cols.tenBC]),
            maXa:    cols.maXa   ? cleanText(r[cols.maXa])   : "",
            tenXa:   cols.tenXa  ? cleanText(r[cols.tenXa])  : "",
            maTinh:  cols.maTinh ? cleanText(r[cols.maTinh]) : "",
            tenTinh: tenTinh,
            search:  fold(ma + " " + r[cols.tenBC])
        };
        MBC.rows.push(row);
        const key = codeKey(ma);
        if (MBC.map.has(key)) MBC.duplicates.push(ma);
        else MBC.map.set(key, row);          // trùng mã: giữ dòng xuất hiện đầu tiên
    });

    MBC.loaded = true;
    console.log("[MBC] Đã nạp", MBC.map.size, "mã bưu cục.");
    if (MBC.duplicates.length) {
        console.warn("[MBC] Mã bị trùng trong file Excel:", Array.from(new Set(MBC.duplicates)).join(", "));
    }
}

/*==========================================================
    DOM
==========================================================*/
function findInput() {
    return document.querySelector('input[name="Mã Bưu cục"]');
}
function getField(names) {
    for (const n of names) {
        const el = document.querySelector('[name="' + n + '"]');
        if (el) return el;
    }
    return null;
}

function injectCSS() {
    if (document.getElementById("mbc-style")) return;
    const style = document.createElement("style");
    style.id = "mbc-style";
    style.textContent = `
.mbc-wrapper{position:relative;width:100%}
.mbc-dropdown{position:absolute;left:0;right:0;top:calc(100% + 4px);background:#fff;border:1px solid #d0d7de;border-radius:8px;box-shadow:0 10px 30px rgba(0,0,0,.12);display:none;overflow-y:auto;max-height:280px;z-index:999999}
.mbc-item{padding:10px 14px;cursor:pointer;transition:.15s}
.mbc-item:hover{background:#f5f8ff}
.mbc-item.active{background:#e9f2ff}
.mbc-code{font-weight:600;color:#0b57d0;font-size:14px}
.mbc-name{font-size:13px;color:#666;margin-top:2px}
.mbc-sub{font-size:12px;color:#999;margin-top:1px}
.mbc-empty{padding:12px;color:#999;text-align:center}
.mbc-invalid{border-color:#d93025 !important;box-shadow:0 0 0 2px rgba(217,48,37,.15)}
.mbc-hint{font-size:11px;color:#8a94a6;margin-top:3px}
.mbc-dropdown mark{background:#ffe58f;padding:0}
`;
    document.head.appendChild(style);
}

function createDropdown() {
    const input = findInput();
    if (!input) return;
    if (input.dataset.mbcReady === "1") {
        MBC.input = input;
        MBC.wrapper = input.closest(".mbc-wrapper");
        if (MBC.wrapper) MBC.dropdown = MBC.wrapper.querySelector(".mbc-dropdown");
        return;
    }
    if (!input.parentNode) return;
    input.dataset.mbcReady = "1";
    input.setAttribute("autocomplete", "off");
    input.setAttribute("placeholder", "Nhập mã hoặc tên bưu cục...");
    const wrapper = document.createElement("div");
    wrapper.className = "mbc-wrapper";
    input.parentNode.insertBefore(wrapper, input);
    wrapper.appendChild(input);
    const dropdown = document.createElement("div");
    dropdown.className = "mbc-dropdown";
    wrapper.appendChild(dropdown);
    const hint = document.createElement("div");
    hint.className = "mbc-hint";
    hint.textContent = "Danh mục bưu cục: đang tải... · bản " + MBC_BUILD;
    wrapper.appendChild(hint);
    MBC.input = input;
    MBC.wrapper = wrapper;
    MBC.dropdown = dropdown;
}

/*==========================================================
    TÌM KIẾM + HIỂN THỊ
==========================================================*/
function filterData(keyword) {
    const k = fold(keyword.trim());
    if (!k) return MBC.rows.slice(0, MAX_RESULT);
    const starts = [], contains = [];
    for (const row of MBC.rows) {
        if (!row.search.includes(k)) continue;
        (fold(row.ma).startsWith(k) ? starts : contains).push(row);
        if (starts.length >= MAX_RESULT) break;
    }
    return starts.concat(contains).slice(0, MAX_RESULT);
}

function renderDropdown(keyword) {
    if (!MBC.dropdown) return;
    keyword = (keyword || "").trim();
    if (!MBC.loaded) {
        MBC.dropdown.innerHTML = '<div class="mbc-empty">' +
            escapeHtml(MBC.loadError ? "Không nạp được danh mục bưu cục: " + MBC.loadError : "Đang tải danh mục bưu cục...") + "</div>";
        MBC.dropdown.style.display = "block";
        return;
    }
    const data = filterData(keyword);
    MBC.filtered = data;
    MBC.selectedIndex = -1;
    if (!data.length) {
        MBC.dropdown.innerHTML = '<div class="mbc-empty">Không tìm thấy Mã Bưu cục</div>';
        MBC.dropdown.style.display = "block";
        return;
    }
    MBC.dropdown.innerHTML = data.map((row, i) =>
        '<div class="mbc-item" data-index="' + i + '">' +
        '<div class="mbc-code">' + highlight(row.ma, keyword) + "</div>" +
        '<div class="mbc-name">' + highlight(row.tenBC, keyword) + "</div>" +
        (row.tenXa ? '<div class="mbc-sub">' + escapeHtml(row.tenXa) + (row.maXa ? " · " + escapeHtml(row.maXa) : "") + "</div>" : "") +
        "</div>"
    ).join("");
    MBC.dropdown.style.display = "block";
}

function updateHint() {
    const el = MBC.wrapper ? MBC.wrapper.querySelector(".mbc-hint") : null;
    if (!el) return;
    el.textContent = MBC.loaded
        ? "Danh mục bưu cục: " + MBC.map.size + " mã · bản " + MBC_BUILD + " · nhập mã rồi Enter để lấy thông tin máy"
        : "Danh mục bưu cục: chưa nạp được · bản " + MBC_BUILD;
}
function hideDropdown() { if (MBC.dropdown) MBC.dropdown.style.display = "none"; }

function updateActiveItem() {
    if (!MBC.dropdown) return;
    MBC.dropdown.querySelectorAll(".mbc-item").forEach((item, i) => {
        item.classList.toggle("active", i === MBC.selectedIndex);
        if (i === MBC.selectedIndex) item.scrollIntoView({ block: "nearest" });
    });
}
function moveDown() {
    if (!MBC.filtered.length) return;
    MBC.selectedIndex = (MBC.selectedIndex + 1) % MBC.filtered.length;
    updateActiveItem();
}
function moveUp() {
    if (!MBC.filtered.length) return;
    MBC.selectedIndex = MBC.selectedIndex <= 0 ? MBC.filtered.length - 1 : MBC.selectedIndex - 1;
    updateActiveItem();
}

/*==========================================================
    TỰ ĐIỀN / XÓA CÁC TRƯỜNG LIÊN QUAN
==========================================================*/
function fillFromRow(row) {
    if (!row) return false;
    FIELD_MAP.forEach(f => {
        const el = getField(f.names);
        if (!el) return;
        el.value = row[f.key] || "";
        fireEvents(el);
    });
    MBC.autoFilled = true;
    return true;
}
function clearRelated() {
    FIELD_MAP.forEach(f => {
        const el = getField(f.names);
        if (!el || !el.value) return;
        el.value = "";
        fireEvents(el);
    });
    MBC.autoFilled = false;
}
function fillFields(maBC) {
    return fillFromRow(MBC.map.get(codeKey(maBC)));
}

function selectItem(index) {
    if (index < 0 || index >= MBC.filtered.length) return;
    const row = MBC.filtered[index];       // dùng đúng dòng được chọn (kể cả khi mã bị trùng)
    MBC.input.value = row.ma;
    fillFromRow(row);
    hideError();
    hideDropdown();
    MBC.selectedIndex = -1;
    MBC.filtered = [];
}

/*==========================================================
    SỰ KIỆN INPUT
==========================================================*/
function onInput(e) {
    const value = e.target.value;
    const code = cleanText(value);
    // 1) gõ đủ mã hợp lệ -> tự điền ngay
    if (MBC.loaded && MBC.map.has(codeKey(code))) {
        fillFields(code);
        hideError();
    } else if (MBC.autoFilled) {
        // mã đã sửa thành giá trị không khớp -> xóa dữ liệu tự điền cũ
        clearRelated();
    }
    // 2) cập nhật gợi ý
    clearTimeout(typingTimer);
    typingTimer = setTimeout(() => renderDropdown(value), 120);
}
function onChange() {
    // trường hợp dán mã hoặc trình duyệt tự điền
    const code = cleanText(MBC.input.value);
    if (MBC.loaded && MBC.map.has(codeKey(code))) { fillFields(code); hideError(); }
}
function onFocus()  { renderDropdown(MBC.input.value); }
function onClick()  { if (MBC.dropdown && MBC.dropdown.style.display === "none") renderDropdown(MBC.input.value); }
function onBlur()   { setTimeout(() => { validateMBC(); hideDropdown(); }, 150); }
/* Xác nhận mã bưu cục hợp lệ -> yêu cầu app chạy lay_thong_tin.bat lấy thông tin máy */
function confirmCode(code) {
    if (typeof window.requestMachineInfo === "function") {
        window.requestMachineInfo(codeKey(code) ? cleanText(code) : "");
    }
}
function onKeyDown(e) {
    switch (e.key) {
        case "ArrowDown": e.preventDefault(); moveDown(); break;
        case "ArrowUp":   e.preventDefault(); moveUp();   break;
        case "Escape":    hideDropdown(); break;
        case "Enter": {
            if (MBC.selectedIndex >= 0) {
                e.preventDefault();
                const row = MBC.filtered[MBC.selectedIndex];
                selectItem(MBC.selectedIndex);
                if (row) confirmCode(row.ma);
                break;
            }
            const code = cleanText(MBC.input.value);
            if (!code) break;                       // để trống: giữ hành vi mặc định
            e.preventDefault();                     // không gửi form khi đang nhập mã
            if (!MBC.loaded) { showError(MBC.loadError ? "Không nạp được danh mục bưu cục." : "Đang tải danh mục bưu cục, vui lòng thử lại."); break; }
            if (!MBC.map.has(codeKey(code))) { showError("Mã Bưu cục không tồn tại."); break; }
            fillFields(code);
            hideError();
            hideDropdown();
            confirmCode(code);
            break;
        }
    }
}

function bindInputEvents() {
    if (!MBC.input || MBC.input.dataset.mbcBound === "1") return;
    MBC.input.dataset.mbcBound = "1";
    MBC.input.addEventListener("input", onInput);
    MBC.input.addEventListener("change", onChange);
    MBC.input.addEventListener("focus", onFocus);
    MBC.input.addEventListener("blur", onBlur);
    MBC.input.addEventListener("click", onClick);
    MBC.input.addEventListener("keydown", onKeyDown);
    if (!outsideClickBound) {
        outsideClickBound = true;
        document.addEventListener("click", e => {
            if (MBC.wrapper && !MBC.wrapper.contains(e.target)) hideDropdown();
        });
    }
}
function bindDropdownClick() {
    if (!MBC.dropdown || MBC.dropdown.dataset.bound === "1") return;
    MBC.dropdown.dataset.bound = "1";
    // giữ focus ở ô nhập khi bấm vào gợi ý
    MBC.dropdown.addEventListener("mousedown", e => e.preventDefault());
    MBC.dropdown.addEventListener("click", e => {
        const item = e.target.closest(".mbc-item");
        if (item) selectItem(Number(item.dataset.index));
    });
}

/*==========================================================
    THÔNG BÁO LỖI + KIỂM TRA KHI GỬI PHIẾU
==========================================================*/
function getErrorBox() {
    if (!MBC.wrapper) return null;
    let box = MBC.wrapper.querySelector(".mbc-error");
    if (box) return box;
    box = document.createElement("div");
    box.className = "mbc-error";
    box.style.cssText = "display:none;color:#d93025;font-size:13px;margin-top:4px";
    MBC.wrapper.appendChild(box);
    return box;
}
function showError(message) {
    const box = getErrorBox();
    if (!box) return;
    box.textContent = message;
    box.style.display = "block";
    if (MBC.input) MBC.input.classList.add("mbc-invalid");
}
function hideError() {
    const box = MBC.wrapper ? MBC.wrapper.querySelector(".mbc-error") : null;
    if (box) box.style.display = "none";
    if (MBC.input) MBC.input.classList.remove("mbc-invalid");
}
function validateMBC() {
    if (!MBC.input || !MBC.loaded) return true;   // chưa nạp được danh mục thì không chặn gửi phiếu
    const value = cleanText(MBC.input.value);
    if (value === "") { hideError(); return true; }
    if (!MBC.map.has(codeKey(value))) { showError("Mã Bưu cục không tồn tại."); return false; }
    hideError();
    return true;
}
function bindSubmitValidation() {
    if (submitBound) return;
    submitBound = true;
    document.addEventListener("submit", function (e) {
        if (!MBC.input || !document.body.contains(MBC.input)) return;
        if (!validateMBC()) {
            e.preventDefault();
            e.stopPropagation();
            MBC.input.focus();
            return;
        }
        // mã hợp lệ nhưng các trường liên quan đang trống -> điền bổ sung trước khi gửi
        const code = cleanText(MBC.input.value);
        const tenBC = getField(["Tên Bưu cục"]);
        if (code && MBC.map.has(codeKey(code)) && tenBC && !tenBC.value) fillFields(code);
    }, true);
}

/*==========================================================
    THEO DÕI FORM ĐỘNG (form được vẽ lại khi đổi nhóm thiết bị)
==========================================================*/
function attach() {
    const input = findInput();
    if (!input || input === MBC.input) return;
    MBC.autoFilled = false;
    createDropdown();
    updateHint();
    bindInputEvents();
    bindDropdownClick();
}
function observeForm() {
    if (observer) observer.disconnect();
    observer = new MutationObserver(attach);
    observer.observe(document.getElementById("dynamic-fields") || document.body, { childList: true, subtree: true });
}

/*==========================================================
    INIT
==========================================================*/
async function init() {
    injectCSS();
    attach();                // gắn ô nhập ngay, danh mục nạp xong sẽ dùng được
    bindSubmitValidation();
    observeForm();
    try {
        MBC.loadPromise = loadExcel();
        await MBC.loadPromise;
        MBC.loadError = null;
        updateHint();
        // nếu người dùng đã lỡ gõ mã trước khi danh mục nạp xong
        if (MBC.input && MBC.map.has(codeKey(MBC.input.value)) && !(getField(["Tên Bưu cục"]) || {}).value) {
            fillFields(MBC.input.value);
        }
        console.log("[MBC] Module initialized.");
    } catch (err) {
        MBC.loadError = err.message;
        console.error("[MBC]", err);
        showError("Không nạp được danh mục bưu cục: " + err.message);
    }
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
else init();

/* Đặt mã bưu cục bằng code (dùng khi nhận thông tin máy ở tab mới) và tự điền các trường liên quan */
async function setCode(code) {
    try { if (MBC.loadPromise) await MBC.loadPromise; } catch (_) { return false; }
    attach();
    const input = findInput();
    if (!input) return false;
    input.value = cleanText(code);
    const ok = fillFields(code);
    if (ok) hideError();
    return ok;
}

window.MBCModule = {
    reload: function () { MBC.loaded = false; return init(); },
    validate: validateMBC,
    fill: fillFields,
    setCode: setCode
};
})();
