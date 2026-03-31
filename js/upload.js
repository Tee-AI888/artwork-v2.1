// =============================================================================
// File Upload & Processing
// =============================================================================
async function autoLoadTemplate() {
    const status = document.getElementById("autoLoadStatus");
    const title  = document.getElementById("excelTitle");
    const badge  = document.getElementById("excelFileName");
    const box    = document.getElementById("excelUploadBox");
    const icon   = document.getElementById("excelIcon");

    status.textContent = "⏳ กำลังตรวจสอบ Template จากระบบ...";
    status.classList.remove("hidden");

    try {
        const res = await fetchWithAuth("/api/get-template");
        if (!res.ok) throw new Error("not found");
        const blob = await res.blob();
        state.excelFile = new File([blob], "Templates.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

        badge.innerHTML = `<i class="fa-solid fa-check mr-1"></i> Templates.xlsx (Auto-Loaded)`;
        badge.classList.remove("hidden");
        title.textContent = "Template Ready";
        status.textContent = "โหลดอัตโนมัติสำเร็จ พร้อมใช้งาน";
        status.className = "text-[10px] mt-2 font-bold italic text-emerald-500";
        box.classList.replace("border-dashed", "border-solid");
        box.classList.replace("border-slate-300", "border-emerald-400");
        box.classList.replace("bg-slate-50/50", "bg-emerald-50/30");
        icon.classList.replace("bg-slate-200", "bg-emerald-100");
        icon.classList.replace("text-slate-500", "text-emerald-600");

        await parseExcelHeaders(blob);
    } catch {
        status.textContent = "⚠️ ไม่พบไฟล์ (กรุณาอัปโหลดเอง)";
        status.className = "text-[10px] mt-2 font-bold italic text-orange-500";
    }
}

async function parseExcelHeaders(blobOrFile) {
    const data  = await blobOrFile.arrayBuffer();
    const wb    = XLSX.read(data);
    if (!wb.SheetNames || !wb.SheetNames.length) {
        await customAlert("ไฟล์ Excel ไม่ถูกต้อง: ไม่พบ Sheet", "error");
        return;
    }
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet || !sheet["!ref"]) {
        await customAlert("ไฟล์ Excel ไม่ถูกต้อง: Sheet ว่างเปล่า", "error");
        return;
    }
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    state.columnHeaders = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ c: C, r: 0 });
        if (sheet[addr]) {
            state.columnHeaders[XLSX.utils.encode_col(C)] = sheet[addr].v.toString().replace(/\s+/g, " ").trim();
        }
    }
}

async function handleAutoMapped(autoMapped) {
    const fresh = autoMapped.filter(am => !state.batchMappedCrops.some(c => c.itemId === am.itemId && c.col === am.col));
    if (!fresh.length) { renderAutoMappedUI(); return; }

    for (const am of fresh) {
        const resp = await fetch(am.imageBase64);
        const blob = await resp.blob();
        state.batchMappedCrops.push({
            filename: `auto_${am.col}_${am.itemId}_${Date.now()}.png`,
            blob,
            col: am.col,
            colName: state.columnHeaders[am.col] || `Col ${am.col}`,
            itemId: am.itemId,
            colorDesc: "Auto-Loaded",
            artboardIndex: "auto",
            previewUrl: am.imageBase64
        });
    }
    renderAutoMappedUI();
}

document.getElementById("excelFile").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    state.excelFile = file;

    const badge  = document.getElementById("excelFileName");
    const box    = document.getElementById("excelUploadBox");
    const icon   = document.getElementById("excelIcon");
    const status = document.getElementById("autoLoadStatus");

    badge.innerHTML = `<i class="fa-solid fa-check mr-1"></i> ${file.name}`;
    badge.classList.remove("hidden");
    document.getElementById("excelTitle").textContent = "Template Uploaded";
    status.textContent = "อัปโหลดทับไฟล์ Template เดิมสำเร็จ";
    status.className = "text-[10px] mt-2 font-bold italic text-emerald-500";
    status.classList.remove("hidden");

    box.classList.replace("border-dashed", "border-solid");
    box.classList.replace("border-slate-300", "border-emerald-400");
    box.classList.replace("bg-slate-50/50", "bg-emerald-50/30");
    icon.classList.replace("bg-slate-200", "bg-emerald-100");
    icon.classList.replace("text-slate-500", "text-emerald-600");

    await parseExcelHeaders(file);
});

document.getElementById("aiFiles").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const span  = document.getElementById("aiFileName");
    const box   = document.getElementById("aiUploadBox");
    const icon  = document.getElementById("aiIcon");
    const title = document.getElementById("aiTitle");

    span.innerHTML = `<i class="fa-solid fa-check mr-1"></i> ${file.name}`;
    span.classList.remove("hidden");
    title.textContent = "Artwork Ready";

    box.classList.replace("border-dashed", "border-solid");
    box.classList.replace("border-slate-300", "border-indigo-400");
    box.classList.replace("bg-slate-50/50", "bg-indigo-50/30");
    icon.classList.replace("bg-slate-200", "bg-indigo-100");
    icon.classList.replace("text-slate-500", "text-indigo-600");

    hideSection("textDataWorkspace");
    hideSection("mainWorkspace");
    hideSection("cropModal");
    hideSection("autoMappedSection");
    document.getElementById("artboardList").innerHTML = "";
    document.getElementById("recordCardsContainer").innerHTML = "";

    if (state.cropper) { state.cropper.destroy(); state.cropper = null; }

    if (state.processedAIFiles.length === 0) {
        hideSection("exportWorkspace");
    } else {
        hideSection("btnSaveProgress");
        showSection("exportWorkspace");
        document.getElementById("exportPanelTitle").innerHTML = `<i class="fa-solid fa-flag-checkered text-emerald-500 mr-2"></i> มีไฟล์รอ Export อยู่ในระบบ (${state.processedAIFiles.length} ไฟล์)`;
        document.getElementById("exportPanelDesc").textContent = "กดปุ่ม 'ดึงข้อมูลมาเริ่มทำงาน' เพื่อทำไฟล์ใหม่ หรือกด Export งานเก่าได้เลย";
    }
});

document.getElementById("btnStart").addEventListener("click", async () => {
    const aiInput = document.getElementById("aiFiles");
    if (!state.excelFile || !aiInput.files.length) return await customAlert("กรุณาอัปโหลดทั้งไฟล์ Excel และ Artwork ครับ", "warning");

    const aiFileName = aiInput.files[0].name;
    if (state.processedAIFiles.includes(aiFileName)) {
        if (!(await customConfirm(`ไฟล์ ${aiFileName} ถูกบันทึกไว้ในระบบแล้ว! ดึงซ้ำอีกครั้งหรือไม่?`))) return;
    }
    if (state.currentMappedCrops.length > 0) {
        if (!(await customConfirm("มีงานค้างบนหน้าจอที่ยังไม่ได้ Save!\nยืนยันที่จะทิ้งงานเดิมหรือไม่?"))) return;
    }

    showLoader("กำลังสกัดข้อมูลและ AI จากไฟล์ Artwork...");
    const formData = new FormData();
    formData.append("file", aiInput.files[0]);
    formData.append("excel_file", state.excelFile);
    formData.append("use_ai", state.useAiMode ? "true" : "false"); 

    try {
        const res  = await fetchWithAuth(`${API_BASE}/extract-ai`, { method: "POST", body: formData });
        const data = await res.json();
        if (data.status !== "success") {
            hideLoader();
            await customAlert(data.message, "error");
            return;
        }

        const batchIds    = state.batchTextData.map(r => r.item_id);
        const dupsInCache = data.text_data.filter(r => batchIds.includes(r.item_id));
        if (dupsInCache.length) {
            hideLoader();
            await customAlert(`❌ พบ ITEM_ID ซ้ำกับไฟล์ที่บันทึกไว้!\nรายการ: ${dupsInCache.map(d => d.item_id).join(", ")}`, "error");
            return;
        }

        state.currentFileName      = aiFileName;
        state.currentColors        = data.colors || {};
        state.currentTextData      = data.text_data.map(r => ({ ...r, source_file: aiFileName }));
        state.currentArtboards     = data.artboards.map(ab => ({ ...ab, source_file: aiFileName, global_id: Math.random() }));
        state.currentMappedCrops   = [];
        state.activeArtboardIndex  = null;
        state.globalCropMemory     = data.crop_memory || {};
        state.manualCompletedArtboards = [];

        showSection("textDataWorkspace");
        showSection("exportWorkspace");
        document.getElementById("exportPanelTitle").innerHTML = `<i class="fa-solid fa-flag-checkered text-indigo-500 mr-2"></i> ทำไฟล์นี้เสร็จแล้วใช่ไหม?`;
        document.getElementById("exportPanelDesc").textContent = "บันทึกเพื่อเคลียร์หน้าจอและไปอัปโหลดไฟล์ถัดไป หรือกด Export ทีเดียวตอนจบงานได้เลย";
        showSection("btnSaveProgress");

        await handleAutoMapped(data.auto_mapped || []);
        renderTextDataCards();
        renderArtboardGallery();

        aiInput.value = "";
        const aiSpan = document.getElementById("aiFileName");
        aiSpan.innerHTML = `<i class="fa-solid fa-file-import mr-1"></i> ดึง ${aiFileName} ขึ้นจอแล้ว`;
        aiSpan.className = "text-sm font-semibold text-indigo-600 bg-indigo-50 px-3 py-1 rounded-full text-center inline-block";
    } catch (err) {
        await customAlert("Cannot connect to server. " + err, "error");
    } finally {
        hideLoader();
    }
});
