/* ============================================================
   MBC.JS
   Module Autocomplete Mã Bưu cục
   PHẦN 1/6
============================================================ */
(function () {
"use strict";
/*==========================================================
    CONFIG
==========================================================*/
const MBC_EXCEL = "./assets/mbc_up.xlsx";
const MBC = {
    rows: [],
    map: new Map(),
    loaded: false,
    dropdown: null,
    input: null,
    wrapper: null,
    selectedIndex: -1,
    filtered: []
};
	let outsideClickBound = false;
	let submitBound = false;
	let typingTimer = null;
	const MAX_RESULT = 20;
/*==========================================================
    LOAD EXCEL
==========================================================*/
async function loadExcel() {
    if (MBC.loaded) return;
    const response = await fetch(MBC_EXCEL);
    if (!response.ok) {
        throw new Error("Không đọc được " + MBC_EXCEL);
    }
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, {
        type: "array"
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: ""
    });
    MBC.rows = rows;
    MBC.map.clear();
    rows.forEach(row => {
        const maBC = String(row["Mã Bưu cục"]).trim();
        if (!maBC) return;
        MBC.map.set(maBC, {
            maTinh: String(row["Mã Tỉnh"]).trim(),
            tenTinh: String(row["Tên Tỉnh"]).trim(),
            maXa: String(row["Mã BĐ Xã"]).trim(),
            tenXa: String(row["Tên Xã"]).trim(),
            tenBC: String(row["Tên Bưu cục"]).trim()
        });
    });
    MBC.loaded = true;
    console.log(
        "[MBC]",
        "Loaded",
        MBC.map.size,
        "records"
    );
}

/*==========================================================
    FIND INPUT
==========================================================*/
function findInput() {
    return document.querySelector(
        'input[name="Mã Bưu cục"]'
    );
}

/*==========================================================
    INJECT CSS
==========================================================*/
function injectCSS() {
    if (document.getElementById("mbc-style"))
        return;
    const style = document.createElement("style");
    style.id = "mbc-style";
    style.textContent = `
.mbc-wrapper{
    position:relative;
    width:100%;
}
.mbc-dropdown{
    position:absolute;
    left:0;
    right:0;
    top:calc(100% + 4px);
    background:#ffffff;
    border:1px solid #d0d7de;
    border-radius:8px;
    box-shadow:0 10px 30px rgba(0,0,0,.12);
    display:none;
    overflow-y:auto;
    max-height:280px;
    z-index:999999;
}
.mbc-item{
    padding:10px 14px;
    cursor:pointer;
    transition:.15s;
}
.mbc-item:hover{
    background:#f5f8ff;
}
.mbc-item.active{
    background:#e9f2ff;
}
.mbc-code{
    font-weight:600;
    color:#0b57d0;
    font-size:14px;
}
.mbc-name{
    font-size:13px;
    color:#666;
    margin-top:2px;
}
.mbc-empty{
    padding:12px;
    color:#999;
    text-align:center;
}
.mbc-invalid{
    border-color:#d93025 !important;
    box-shadow:0 0 0 2px rgba(217,48,37,.15);
}
mark{
    background:#ffe58f;
    padding:0;
}
`; document.head.appendChild(style);

}
/*==========================================================
    CREATE DROPDOWN
==========================================================*/
function createDropdown() {
    const input = findInput();
    if (!input) return;
    // đã tạo rồi
    if (input.dataset.mbcReady === "1") {
        MBC.input = input;
        MBC.wrapper = input.closest(".mbc-wrapper");
	if (MBC.wrapper) {
	    MBC.dropdown = MBC.wrapper.querySelector(".mbc-dropdown");
	}
        return;
    }
    input.dataset.mbcReady = "1";
    if (!input.parentNode) return;
    const wrapper = document.createElement("div");
    wrapper.className = "mbc-wrapper";
    input.parentNode.insertBefore(wrapper, input);
    wrapper.appendChild(input);
    const dropdown = document.createElement("div");
    dropdown.className = "mbc-dropdown";
    wrapper.appendChild(dropdown);
    MBC.input = input;
    MBC.wrapper = wrapper;
    MBC.dropdown = dropdown;
}

/*==========================================================
    ESCAPE HTML
==========================================================*/

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

/*==========================================================
    HIGHLIGHT
==========================================================*/
function highlight(text, keyword) {
    if (!keyword) return escapeHtml(text);
    const safe = escapeHtml(text);
    const reg = new RegExp(
        "(" +
        keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        +
        ")",
        "ig"
    );
    return safe.replace(reg, "<mark>$1</mark>");
}

/*==========================================================
    FILTER DATA
==========================================================*/
function filterData(keyword) {
    keyword = keyword.trim().toLowerCase();
    if (!keyword) {
        return MBC.rows.slice(0, MAX_RESULT);
    }
    return MBC.rows.filter(row => {
        const ma =
            String(row["Mã Bưu cục"])
            .toLowerCase();
        const ten =
            String(row["Tên Bưu cục"])
            .toLowerCase();
        return (ma.includes(keyword) || ten.includes(keyword)
        );
    }).slice(0, MAX_RESULT);
}

/*==========================================================
    RENDER DROPDOWN
==========================================================*/
function renderDropdown(keyword) {
    if (!MBC.dropdown)
        return;
    const data = filterData(keyword);
    MBC.filtered = data;
    MBC.selectedIndex = -1;
    if (!data.length) {
        MBC.dropdown.innerHTML = `
<div class="mbc-empty">
Không tìm thấy Mã Bưu cục
</div>
`;
        MBC.dropdown.style.display = "block";
        return;
    }
    let html = "";
    data.forEach((row, index) => {
        const ma = String(row["Mã Bưu cục"]);
        const ten = String(row["Tên Bưu cục"]);
        html += `
<div
    class="mbc-item"
    data-index="${index}"
    data-code="${escapeHtml(ma)}"
>
<div class="mbc-code">
${highlight(ma, keyword)}
</div>
<div class="mbc-name">
${highlight(ten, keyword)}
</div>
</div>
`;
    });
    MBC.dropdown.innerHTML = html;
    MBC.dropdown.style.display = "block";
}

/*==========================================================
    HIDE DROPDOWN
==========================================================*/
function hideDropdown() {
    if (!MBC.dropdown)
        return;
    MBC.dropdown.style.display = "none";
}

/*==========================================================
    SHOW DROPDOWN
==========================================================*/
function showDropdown() {
    if (!MBC.dropdown)
        return;
    MBC.dropdown.style.display = "block";
}

/*==========================================================
    MOVE ACTIVE ITEM
==========================================================*/
function updateActiveItem() {
    if (!MBC.dropdown) return;
    const items = MBC.dropdown.querySelectorAll(".mbc-item");
    items.forEach((item, index) => {
        item.classList.toggle(
            "active",
            index === MBC.selectedIndex
        );
        if (index === MBC.selectedIndex) {
            item.scrollIntoView({
                block: "nearest"
            });
        }
    });
}

/*==========================================================
    MOVE DOWN
==========================================================*/

function moveDown() {
    if (!MBC.filtered.length) return;
    MBC.selectedIndex++;
    if (MBC.selectedIndex >= MBC.filtered.length) {
        MBC.selectedIndex = 0;
    }
    updateActiveItem();
}

/*==========================================================
    MOVE UP
==========================================================*/
function moveUp() {
    if (!MBC.filtered.length) return;
    MBC.selectedIndex--;
    if (MBC.selectedIndex < 0) {
        MBC.selectedIndex =
            MBC.filtered.length - 1;

    }
    updateActiveItem();
}

/*==========================================================
    INPUT EVENT
==========================================================*/
function onInput(e) {
    const keyword = e.target.value.trim();
    clearTimeout(typingTimer);
    typingTimer = setTimeout(() => {
        renderDropdown(keyword);
    }, 120);
}
/*==========================================================
    FOCUS EVENT
==========================================================*/
function onFocus() {
    renderDropdown(MBC.input.value);
}

/*==========================================================
    BLUR EVENT
==========================================================*/
function onBlur() {
    setTimeout(() => {
        validateMBC();
        hideDropdown();
    },150);
}

/*==========================================================
    CLICK EVENT
==========================================================*/
function onClick() {
    if (!MBC.dropdown) return;
    if (MBC.dropdown.style.display === "none") {
        renderDropdown(MBC.input.value);
    }
}

/*==========================================================
    KEYDOWN
==========================================================*/
function onKeyDown(e) {
    switch (e.key) {
        case "ArrowDown":
            e.preventDefault();
            moveDown();
            break;
        case "ArrowUp":
            e.preventDefault();
            moveUp();
            break;
        case "Escape":
            hideDropdown();
            break;
        case "Enter":
	    if (MBC.selectedIndex >= 0) {
	        e.preventDefault();
	        selectItem(MBC.selectedIndex);
 	   }
    break;
    }
}

/*==========================================================
    CLICK OUTSIDE
==========================================================*/
function bindOutsideClick() {
  if(outsideClickBound) return;
    outsideClickBound = true;  
	document.addEventListener("click", e => {
        if (!MBC.wrapper) return;
        if (!MBC.wrapper.contains(e.target)) {
            hideDropdown();
        }
    });
}

/*==========================================================
    BIND INPUT EVENTS
==========================================================*/
function bindInputEvents() {
    if (!MBC.input) return;
    if (MBC.input.dataset.mbcBound === "1")
        return;
    MBC.input.dataset.mbcBound = "1";
    MBC.input.addEventListener(
        "input",
        onInput
    );
    MBC.input.addEventListener(
        "focus",
        onFocus
    );
    MBC.input.addEventListener(
        "blur",
        onBlur
    );
    MBC.input.addEventListener(
        "click",
        onClick
    );
    MBC.input.addEventListener(
        "keydown",
        onKeyDown
    );
    bindOutsideClick();
}

/*==========================================================
    FIND FIELD
==========================================================*/
function getField(name) {
    return document.querySelector(
        '[name="' + name + '"]'
    );
}

/*==========================================================
    FILL RELATED FIELDS
==========================================================*/
function fillFields(maBC) {
    maBC = String(maBC).trim();
    const info = MBC.map.get(maBC);
    if (!info) return false;
    const tenBC = getField("Tên Bưu cục");
    const maXa = getField("Mã BĐ Xã");
    const tenXa = getField("Tên Xã")
    || getField("Tên xã")
    || getField("Tên xã/phường");
    if (tenBC) {
        tenBC.value = info.tenBC;
        tenBC.dispatchEvent( new Event("input",{bubbles:true}) );
        tenBC.dispatchEvent( new Event("change",{bubbles:true}) );
    }
    if (maXa) {
        maXa.value = info.maXa;
        maXa.dispatchEvent( new Event("input",{bubbles:true}) );
        maXa.dispatchEvent( new Event("change",{bubbles:true}) );
    }
    if (tenXa) {
        tenXa.value = info.tenXa;
        tenXa.dispatchEvent( new Event("input",{bubbles:true}) );
        tenXa.dispatchEvent( new Event("change",{bubbles:true}) );
    }
    return true;
}

/*==========================================================
    SELECT ITEM
==========================================================*/
function selectItem(index) {
    if (
        index < 0 ||
        index >= MBC.filtered.length
    ) {
        return;
    }
    const row = MBC.filtered[index];
    const maBC = String(row["Mã Bưu cục"]).trim();
    MBC.input.value = maBC;
    fillFields(maBC);
    hideError();
    hideDropdown();
    // Reset trạng thái chọn
    MBC.selectedIndex = -1;
    MBC.filtered = [];
}

/*==========================================================
    CLICK ITEM
==========================================================*/
function bindDropdownClick() {
    if (!MBC.dropdown) return;
    if (MBC.dropdown.dataset.bound === "1") return;
    MBC.dropdown.dataset.bound = "1";
    MBC.dropdown.addEventListener("click", e => {
        const item = e.target.closest(".mbc-item");
        if (!item) return;
        selectItem(Number(item.dataset.index));
    });
}
/*==========================================================
    ERROR UI
==========================================================*/
function getErrorBox() {
    if (!MBC.wrapper) return null;
    let box = MBC.wrapper.querySelector(".mbc-error");
    if (box) return box;
    box = document.createElement("div");
    box.className = "mbc-error";
    box.style.display = "none";
    box.style.color = "#d93025";
    box.style.fontSize = "13px";
    box.style.marginTop = "4px";
    MBC.wrapper.appendChild(box);
    return box;
}

/*==========================================================
    SHOW ERROR
==========================================================*/
function showError(message) {
    const box = getErrorBox();
	if (!box) return;
		box.textContent = message;
		box.style.display = "block";
   	if (MBC.input) {
        	MBC.input.classList.add("mbc-invalid");
    }
}

/*==========================================================
    HIDE ERROR
==========================================================*/
function hideError() {
    const box = MBC.wrapper ? MBC.wrapper.querySelector(".mbc-error")  : null;
    if (box) {
        box.style.display = "none";
    }
    if (MBC.input) {
        MBC.input.classList.remove("mbc-invalid");
    }
}

/*==========================================================
    VALIDATE
==========================================================*/
function validateMBC() {
    if (!MBC.input) return true;
    const value =
        MBC.input.value.trim();
    if (value === "") {
        hideError();
        return true;
    }
    if (!MBC.map.has(value)) {
        showError("Mã Bưu cục không tồn tại.");
        return false;
    }
    hideError();
    return true;
}

/*==========================================================
    BLOCK SUBMIT
==========================================================*/
function bindSubmitValidation() {
    if (submitBound) return;
    submitBound = true;
    document.addEventListener(
        "submit",
        function (e) {
            if (!validateMBC()) {
                e.preventDefault();
                e.stopPropagation();
                MBC.input.focus();
            }
        },
        true
    );
}
/*==========================================================
    AUTO FILL WHEN USER TYPES FULL CODE
==========================================================*/
function bindAutoFill() {
    if (!MBC.input) return;
    if (MBC.input.dataset.autofillBound === "1") return;
    MBC.input.dataset.autofillBound = "1";
    MBC.input.addEventListener(
        "input",
        () => {
            const code = MBC.input.value.trim();
            if (MBC.map.has(code)) {
                fillFields(code);
                hideError();
            }
        }
    );

}/*==========================================================
    OBSERVE FORM
==========================================================*/
let observer = null;
function observeForm() {
    if (observer) {
        observer.disconnect();
    }
    observer = new MutationObserver(() => {
        const input = findInput();
        if (!input) return;
        if (input !== MBC.input) {
	    createDropdown();
	    bindInputEvents();
	    bindDropdownClick();
	    bindAutoFill();
	}
    });
    // Ưu tiên theo dõi vùng form động
    const target =
        document.getElementById("dynamic-fields") ||
        document.body;
    observer.observe(target, {
        childList: true,
        subtree: true
    });
}
/*==========================================================
    INIT
==========================================================*/
async function init() {
    try {
        await loadExcel();
        injectCSS();
        createDropdown();
	bindDropdownClick();
        bindInputEvents();
        bindAutoFill();
        bindSubmitValidation();
	observeForm();
        console.log(  "[MBC] Module initialized." );
    }
    catch (err) {
        console.error(  "[MBC]", err);
    }
}
/*==========================================================
    START
==========================================================*/
if (
    document.readyState === "loading") {
    document.addEventListener( "DOMContentLoaded", init);
}
else {
    init();
}
/*==========================================================
    EXPORT
==========================================================*/
window.MBCModule = {
    reload: init,
    validate: validateMBC,
    fill: fillFields
};
})();