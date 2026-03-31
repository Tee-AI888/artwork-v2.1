// =============================================================================
// Application Entry Point & Keyboard Bindings
// =============================================================================

window.addEventListener("load", async () => {
    if (getAuthToken()) {
        try {
            const res = await fetchWithAuth("/api/me");
            if (res.ok) {
                setupUserProfile(await res.json());
                document.getElementById("loginModal").classList.add("hidden");
                loadMasterTeamsForMapping(); 
                autoLoadTemplate(); 
            } else logout();
        } catch (e) { logout(); }
    } else document.getElementById("loginModal").classList.remove("hidden");
});

document.addEventListener('keydown', function(e) {
    const cropModal = document.getElementById("cropModal");
    if (!state.cropper || cropModal.classList.contains('hidden') || ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
    
    const step = e.shiftKey ? 10 : 1; 
    const cropData = state.cropper.getData();
    
    switch(e.key) {
        case 'ArrowUp': e.preventDefault(); state.cropper.setData({ y: cropData.y - step }); break;
        case 'ArrowDown': e.preventDefault(); state.cropper.setData({ y: cropData.y + step }); break;
        case 'ArrowLeft': e.preventDefault(); state.cropper.setData({ x: cropData.x - step }); break;
        case 'ArrowRight': e.preventDefault(); state.cropper.setData({ x: cropData.x + step }); break;
        case 'Enter': e.preventDefault(); document.getElementById('btnAddCrop').click(); break;
    }
    
    if (e.key >= '1' && e.key <= '9') { 
        const colButtons = document.querySelectorAll(".col-card"); 
        const index = parseInt(e.key) - 1; 
        if (colButtons[index]) { 
            e.preventDefault(); 
            colButtons[index].click(); 
        } 
    }
});

document.getElementById("chkLockRatio").addEventListener("change", function () { 
    document.getElementById("iconLockRatio").className = this.checked ? "fa-solid fa-link" : "fa-solid fa-link-slash"; 
    this.parentNode.classList.toggle("text-indigo-600",  this.checked); 
    this.parentNode.classList.toggle("text-slate-400",  !this.checked); 
});

document.getElementById("cropWidthCm").addEventListener("change", function () { 
    if (!state.cropper || state.cmPerPx === 0) return; 
    const newW = parseFloat(this.value); 
    if (isNaN(newW) || newW <= 0) return; 
    
    const cur = state.cropper.getData(); 
    let newH = parseFloat(document.getElementById("cropHeightCm").value); 
    if (document.getElementById("chkLockRatio").checked) { 
        newH = newW / (cur.width / cur.height); 
        document.getElementById("cropHeightCm").value = newH.toFixed(2); 
    } 
    state.cropper.setData({ width: newW / state.cmPerPx, height: newH / state.cmPerPx }); 
});

document.getElementById("cropHeightCm").addEventListener("change", function () { 
    if (!state.cropper || state.cmPerPx === 0) return; 
    const newH = parseFloat(this.value); 
    if (isNaN(newH) || newH <= 0) return; 
    
    const cur = state.cropper.getData(); 
    let newW = parseFloat(document.getElementById("cropWidthCm").value); 
    if (document.getElementById("chkLockRatio").checked) { 
        newW = newH * (cur.width / cur.height); 
        document.getElementById("cropWidthCm").value = newW.toFixed(2); 
    } 
    state.cropper.setData({ width: newW / state.cmPerPx, height: newH / state.cmPerPx }); 
});

document.getElementById("btnToggleLock").addEventListener("click", () => {
    state.isColorLocked = !state.isColorLocked; 
    const btn = document.getElementById("btnToggleLock"), icon = document.getElementById("lockIcon"), text = document.getElementById("lockText");
    if (state.isColorLocked) { 
        btn.classList.replace("bg-slate-100","bg-amber-100"); 
        btn.classList.replace("text-slate-500","text-amber-700"); 
        btn.classList.replace("border-slate-200","border-amber-300"); 
        icon.className = "fa-solid fa-lock"; 
        text.textContent = "Locked"; 
    } else { 
        btn.classList.replace("bg-amber-100","bg-slate-100"); 
        btn.classList.replace("text-amber-700","text-slate-500"); 
        btn.classList.replace("border-amber-300","border-slate-200"); 
        icon.className = "fa-solid fa-unlock"; 
        text.textContent = "Auto-Clear"; 
    }
});

document.getElementById("btnSelectAllColors").addEventListener("click", () => { 
    const all = document.querySelectorAll(".rec-card"); 
    const allSelected = Array.from(all).every(b => b.classList.contains("selected")); 
    all.forEach(b => b.classList.toggle("selected", !allSelected)); 
    renderAvailableColumns(); 
});

document.getElementById("btnAddCrop").addEventListener("click", async () => {
    const selectedColBtn  = document.querySelector(".col-card.selected"); 
    const selectedRecBtns = document.querySelectorAll(".rec-card.selected");
    
    if (!selectedColBtn || !selectedRecBtns.length) return await customAlert("กรุณาเลือก 'Fabric Color' และ 'Target Column' ให้ครบก่อนกดหั่นรูปครับ!", "warning");
    
    const col = selectedColBtn.dataset.col;
    for (const recBtn of selectedRecBtns) { 
        if (isMapped(recBtn.dataset.itemid, col)) return await customAlert(`คอลัมน์ ${col} สำหรับสี ${recBtn.dataset.color} ถูกใช้งานไปแล้วครับ!`, "warning"); 
    }

    const cropData = state.cropper.getData();
    const memory = JSON.parse(localStorage.getItem(COL_MEMORY_KEY) || "{}"); 
    memory[col] = cropData; 
    localStorage.setItem(COL_MEMORY_KEY, JSON.stringify(memory)); 
    state.sessionCropMemory[col] = cropData;

    const croppedCanvas = state.cropper.getCroppedCanvas();
    if (!croppedCanvas) return await customAlert("พื้นที่ตัดเล็กเกินไป กรุณาลากใหม่", "warning");
    croppedCanvas.toBlob((blob) => {
        selectedRecBtns.forEach(recBtn => {
            const itemId = recBtn.dataset.itemid; 
            const styleCode = getStyleFromItemId(itemId); 
            const filename = `crop_${Date.now()}_${Math.floor(Math.random() * 10000)}.png`;
            
            if (!state.globalCropMemory[styleCode]) state.globalCropMemory[styleCode] = {}; 
            state.globalCropMemory[styleCode][col] = cropData;
            
            state.currentMappedCrops.push({ 
                filename, 
                blob, 
                col, 
                colName: selectedColBtn.dataset.colname, 
                itemId: itemId, 
                colorDesc: recBtn.dataset.color, 
                artboardIndex: state.activeArtboardIndex, 
                previewUrl: URL.createObjectURL(blob), 
                cropData: cropData 
            });
            
            const fdMem = new FormData(); 
            fdMem.append("item_id", styleCode); 
            fdMem.append("column", col); 
            fdMem.append("crop_data", JSON.stringify(cropData)); 
            fetchWithAuth(`${API_BASE}/save-crop-memory`, { method: "POST", body: fdMem }).catch(console.warn);
            
            const fd = new FormData(); 
            fd.append("file", blob, "training_crop.png"); 
            fd.append("label", col); 
            fd.append("item_id", itemId); 
            fetchWithAuth(`${API_BASE}/save-dataset`, { method: "POST", body: fd }).catch(console.warn);
        });
        
        document.querySelectorAll(".col-card").forEach(b => b.classList.remove("selected"));
        if (!state.isColorLocked) document.querySelectorAll(".rec-card").forEach(b => b.classList.remove("selected"));
        
        renderMappedItems(); 
        renderAvailableColumns(); 
        renderArtboardGallery(); 
        renderTextDataCards();
    }, "image/png");
});
// =============================================================================
// AI Mode Toggle — sync state.useAiMode กับ checkbox
// =============================================================================
document.getElementById("toggleAiMode").addEventListener("change", function () {
    state.useAiMode = this.checked;

    if (!this.checked) {
        // ปิด AI → abort การสแกนที่กำลังทำ + ซ่อน panels
        if (state.aiAbortController) state.aiAbortController.abort();
        document.getElementById("aiStatusBadge").classList.add("hidden");
        document.getElementById("detectedObjectsPanel").classList.add("hidden");
        const list = document.getElementById("detectedObjectsList");
        if (list) list.innerHTML = "";
    }
});
