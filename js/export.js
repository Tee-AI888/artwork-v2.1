// =============================================================================
// Export & Save Progress
// =============================================================================
document.getElementById("btnSaveProgress").addEventListener("click", async () => {
    if (!state.currentMappedCrops.length && !state.batchMappedCrops.length) return await customAlert("คุณยังไม่มีรูปเตรียมส่งเข้า Excel เลยครับ!", "warning");
    
    const missingColors = [], missingCrops = [];
    state.currentTextData.forEach(rec => {
        if (!rec.color || rec.color.trim() === "") missingColors.push(rec.item_id);
        const hasCrop = state.currentMappedCrops.some(c => c.itemId === rec.item_id) || state.batchMappedCrops.some(c => c.itemId === rec.item_id);
        if (!hasCrop) missingCrops.push(rec.item_id);
    });

    if (missingColors.length > 0 || missingCrops.length > 0) {
        let warnMsg = "⚠️ พบข้อมูลไม่ครบถ้วนในไฟล์นี้:\n\n";
        if (missingColors.length > 0) warnMsg += `• ลืมใส่สี (Color): ${missingColors.join(', ')}\n`; 
        if (missingCrops.length > 0) warnMsg += `• ยังไม่ได้ตัดรูปลงคอลัมน์ใดๆ: ${missingCrops.join(', ')}\n`;
        warnMsg += "\nคุณต้องการบันทึกงานต่อไปทั้งๆ ที่ข้อมูลไม่ครบหรือไม่?"; 
        if (!(await customConfirm(warnMsg, "ข้อมูลไม่ครบถ้วน", "danger"))) return; 
    }
    
    state.batchTextData = [...state.batchTextData, ...state.currentTextData]; 
    state.batchMappedCrops = [...state.batchMappedCrops, ...state.currentMappedCrops];
    if (!state.processedAIFiles.includes(state.currentFileName)) state.processedAIFiles.push(state.currentFileName);

    showSection("cacheCounter"); 
    document.getElementById("fileCountVal").textContent = state.processedAIFiles.length;
    
    state.currentArtboards = []; 
    state.currentTextData = []; 
    state.currentMappedCrops = []; 
    state.currentFileName = ""; 
    state.currentColors = {}; 
    state.activeArtboardIndex = null; 
    state.manualCompletedArtboards = [];
    
    if (state.cropper) { 
        state.cropper.destroy(); 
        state.cropper = null; 
    }
    
    hideSection("textDataWorkspace"); 
    hideSection("mainWorkspace"); 
    hideSection("cropModal");
    
    document.getElementById("exportPanelTitle").innerHTML = `<i class="fa-solid fa-flag-checkered text-emerald-500 mr-2"></i> มีไฟล์รอ Export อยู่ในระบบ (${state.processedAIFiles.length} ไฟล์)`; 
    document.getElementById("exportPanelDesc").textContent = "หน้าจอถูกเคลียร์แล้ว! อัปโหลดไฟล์ .ai ถัดไป หรือกด Export ทีเดียวตอนจบงานได้เลย";
    hideSection("btnSaveProgress"); 
    document.getElementById("aiFiles").value = ""; 
    hideSection("aiFileName");
    
    await customAlert(`✅ บันทึกข้อมูลของไฟล์นี้ลงระบบเรียบร้อย!\n(เคลียร์หน้าจอพร้อมรับไฟล์ใหม่แล้วครับ)`, "success");
    document.getElementById("mainScrollArea").scrollTo({ top: 0, behavior: "smooth" }); 
    setTimeout(() => flashBox("aiUploadBox"), 500);
});

document.getElementById("btnClear").addEventListener("click", async () => {
    if (!(await customConfirm("แน่ใจหรือไม่ว่าต้องการล้างข้อมูลทั้งหมดเพื่อเริ่มใหม่?", "ยืนยันการล้างข้อมูล", "danger"))) return;
    
    Object.assign(state, { 
        batchTextData: [], batchMappedCrops: [], processedAIFiles: [], 
        currentArtboards: [], currentTextData: [], currentMappedCrops: [], 
        currentFileName: "", currentColors: {}, activeArtboardIndex: null, 
        cmPerPx: 0, isColorLocked: false, manualCompletedArtboards: [] 
    });
    
    if (state.cropper) { state.cropper.destroy(); state.cropper = null; }
    
    ["textDataWorkspace","mainWorkspace","exportWorkspace","cropArea"].forEach(hideSection);
    
    document.getElementById("aiFiles").value = ""; 
    hideSection("aiFileName"); 
    document.getElementById("cacheCounter").classList.replace("flex","hidden");
    
    const exBox = document.getElementById("excelUploadBox"), aiBox = document.getElementById("aiUploadBox");
    exBox.className = "border-2 border-dashed border-slate-300 bg-slate-50/50 rounded-xl p-6 flex flex-col items-center justify-center text-center hover:border-emerald-400 hover:bg-emerald-50 cursor-pointer transition-all duration-300 h-full min-h-[160px] group";
    aiBox.className = "border-2 border-dashed border-slate-300 bg-slate-50/50 rounded-xl p-6 flex flex-col items-center justify-center text-center hover:border-indigo-400 hover:bg-indigo-50 cursor-pointer transition-all duration-300 h-full min-h-[160px] group";
    document.getElementById("excelIcon").className = "w-12 h-12 bg-slate-200 text-slate-500 rounded-full flex items-center justify-center text-xl mb-3 group-hover:scale-110 transition shadow-sm"; 
    document.getElementById("aiIcon").className = "w-12 h-12 bg-slate-200 text-slate-500 rounded-full flex items-center justify-center text-xl mb-3 group-hover:scale-110 transition shadow-sm";
    document.getElementById("excelTitle").textContent = "Upload Excel Template"; 
    document.getElementById("aiTitle").textContent = "Upload Artwork File (.ai)";
    document.getElementById("excelFileName").classList.add("hidden"); 
    document.getElementById("autoLoadStatus").classList.add("hidden");
    
    autoLoadTemplate();
});

document.getElementById("btnGenerateZip").addEventListener("click", async () => {
    const finalCrops = getAllCrops(); 
    const finalTextData = [...state.batchTextData, ...state.currentTextData];
    
    if (!finalCrops.length) return await customAlert("ยังไม่มีรูปที่ถูกหั่นเลยครับ!", "warning");
    
    showLoader("กำลังสร้างไฟล์ ZIP...");
    const formData = new FormData(); 
    formData.append("text_data", JSON.stringify(finalTextData));
    
    const mappingData = {};
    finalCrops.forEach(crop => { 
        mappingData[crop.filename] = { col: crop.col, col_name: crop.colName, item_id: crop.itemId }; 
        formData.append("images", new File([crop.blob], crop.filename, { type: "image/png" })); 
    });
    formData.append("mapping_data", JSON.stringify(mappingData));
    
    try {
        const res = await fetchWithAuth(`${API_BASE}/generate-zip`, { method: "POST", body: formData });
        if (res.ok) { 
            const url = URL.createObjectURL(await res.blob()); 
            const a = Object.assign(document.createElement("a"), { href: url, download: "Artwork_Assets.zip" }); 
            document.body.appendChild(a); 
            a.click(); 
            a.remove(); 
            URL.revokeObjectURL(url); 
        } 
        else await customAlert(`Export ZIP ล้มเหลว`, "error");
    } catch { 
        await customAlert("Backend connection failed.", "error"); 
    } finally { 
        hideLoader(); 
    }
});

document.getElementById("btnGenerateExcel").addEventListener("click", async () => {
    const finalCrops = getAllCrops(); 
    const finalTextData = [...state.batchTextData, ...state.currentTextData];
    
    if (!finalCrops.length) return await customAlert("ยังไม่มีรูปที่ถูกหั่นเลยครับ!", "warning");
    
    let html = `<div class="overflow-x-auto"><table class="w-full text-left border-collapse min-w-[800px]"><thead><tr class="bg-slate-100 text-slate-600 text-xs uppercase tracking-wider"><th class="p-3 border-b border-slate-200 w-2/12">ITEM ID</th><th class="p-3 border-b border-slate-200 w-2/12">Color</th><th class="p-3 border-b border-slate-200 w-2/12">Fabric</th><th class="p-3 border-b border-slate-200 w-5/12">รูปที่หั่น (Columns)</th><th class="p-3 border-b border-slate-200 text-center w-1/12">Status</th></tr></thead><tbody class="text-sm">`;
    let hasWarning = false;
    
    finalTextData.forEach(rec => {
        const itemCrops = finalCrops.filter(c => c.itemId === rec.item_id);
        const colBadgesHtml = itemCrops.length > 0 
            ? itemCrops.sort((a,b) => a.col.localeCompare(b.col)).map(c => `<span class="inline-flex items-center gap-1.5 bg-white border border-slate-200 px-2 py-1 rounded shadow-sm text-[10px] text-slate-600 mr-1.5 mb-1.5" title="${state.columnHeaders[c.col] || ''}"><span class="font-black text-indigo-600">${c.col}</span>${state.columnHeaders[c.col] ? `<span class="truncate max-w-[120px] text-slate-400 font-medium">${state.columnHeaders[c.col]}</span>` : ''}</span>`).join('') 
            : '<span class="text-slate-300 italic text-xs">ไม่มีรูป</span>';
            
        const isMissingColor = !rec.color || rec.color.trim() === ""; 
        const isMissingCrops = itemCrops.length === 0;
        let statusHtml = `<span class="bg-emerald-100 text-emerald-700 px-2 py-1 rounded text-[10px] font-bold whitespace-nowrap"><i class="fa-solid fa-check"></i> พร้อม</span>`;
        
        if (isMissingColor || isMissingCrops) { 
            hasWarning = true; 
            let errs = []; 
            if (isMissingColor) errs.push("ขาดสี"); 
            if (isMissingCrops) errs.push("ขาดรูป"); 
            statusHtml = `<span class="bg-red-100 text-red-700 px-2 py-1 rounded text-[10px] font-bold whitespace-nowrap"><i class="fa-solid fa-triangle-exclamation"></i> ${errs.join(', ')}</span>`; 
        }
        const fabricDisplay = rec.fabric ? `<div class="text-indigo-600 font-medium truncate max-w-[150px]" title="${rec.fabric}">${rec.fabric}</div>` : `<div class="text-slate-300 italic">- ไม่พบ -</div>`;
        html += `<tr class="border-b border-slate-100 hover:bg-slate-50 transition"><td class="p-3 font-bold text-indigo-700 align-top">${rec.item_id}</td><td class="p-3 ${isMissingColor ? 'text-red-500 italic' : 'text-slate-700'} align-top">${rec.color || '- รอระบุ -'}</td><td class="p-3 align-top">${fabricDisplay}</td><td class="p-3 font-medium text-slate-600 align-top flex flex-wrap pt-3.5">${colBadgesHtml}</td><td class="p-3 text-center align-top pt-3.5">${statusHtml}</td></tr>`;
    });
    
    html += `</tbody></table></div>`;
    if (hasWarning) html = `<div class="mb-4 bg-amber-50 border border-amber-200 text-amber-800 p-3 rounded-lg text-sm font-bold flex items-center shadow-sm"><i class="fa-solid fa-circle-exclamation text-lg mr-2"></i> คำเตือน: พบรายการที่ข้อมูลไม่ครบ โปรดตรวจสอบตารางด้านล่างก่อนยืนยัน!</div>` + html;
    
    document.getElementById("exportPreviewContent").innerHTML = html; 
    document.getElementById("exportPreviewModal").classList.remove("hidden");
});

window.closeExportPreview = () => { 
    document.getElementById("exportPreviewModal").classList.add("hidden"); 
};

document.getElementById("btnConfirmExport").addEventListener("click", async () => {
    closeExportPreview(); 
    const finalCrops = getAllCrops(); 
    const finalTextData = [...state.batchTextData, ...state.currentTextData]; 
    showLoader("กำลังประกอบไฟล์ Excel...");
    
    const formData = new FormData(); 
    formData.append("excel_file", state.excelFile); 
    formData.append("text_data", JSON.stringify(finalTextData));
    
    const mappingData = {}; 
    finalCrops.forEach(crop => { 
        mappingData[crop.filename] = { col: crop.col, item_id: crop.itemId }; 
        formData.append("images", new File([crop.blob], crop.filename, { type: "image/png" })); 
    }); 
    formData.append("mapping_data", JSON.stringify(mappingData));
    
    try {
        const res = await fetchWithAuth(`${API_BASE}/generate-excel`, { method: "POST", body: formData });
        if (res.ok) { 
            const url = URL.createObjectURL(await res.blob()); 
            const a = Object.assign(document.createElement("a"), { href: url, download: "Tracking_Update.xlsx" }); 
            document.body.appendChild(a); 
            a.click(); 
            a.remove(); 
            URL.revokeObjectURL(url); 
        } 
        else { 
            const err = await res.json().catch(() => ({})); 
            await customAlert(`Export ล้มเหลว: ${err.message || res.statusText}`, "error"); 
        }
    } catch { 
        await customAlert("Backend connection failed.", "error"); 
    } finally { 
        hideLoader(); 
    }
});