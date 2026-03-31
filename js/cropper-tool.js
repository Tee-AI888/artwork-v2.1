// =============================================================================
// Cropper Engine & AI Detection
// =============================================================================
function activateCropTool() { 
    document.getElementById("btnToolCrop").classList.add("active"); 
    document.getElementById("btnToolMove").classList.remove("active"); 
}

window.cropperAction = (action) => {
    if (!state.cropper) return;
    switch (action) {
        case "zoom_in": state.cropper.zoom(0.1); break; 
        case "zoom_out": state.cropper.zoom(-0.1); break;
        case "move": 
            state.cropper.setDragMode("move"); 
            document.getElementById("btnToolMove").classList.add("active"); 
            document.getElementById("btnToolCrop").classList.remove("active"); 
            break;
        case "crop": 
            state.cropper.setDragMode("crop"); 
            activateCropTool(); 
            break; 
        case "reset": 
            state.cropper.reset(); 
            activateCropTool(); 
            break;
    }
};

window.openCropperWorkspace = (index) => {
    state.activeArtboardIndex = index; 
    const ab = state.currentArtboards[index]; 
    const cropModal = document.getElementById("cropModal"); 
    const cropArea = document.getElementById("cropArea"); 
    // Move modal to body to escape ancestor overflow clipping
    if (cropModal.parentElement !== document.body) {
        document.body.appendChild(cropModal);
    }
    cropModal.classList.remove("hidden");
    cropArea.classList.remove("hidden");
    document.body.style.overflow = "hidden";

    document.getElementById("artboardDisplay").innerHTML = `Artboard ${ab.artboard_number} <span class="text-xs text-slate-400 font-normal ml-2">(${ab.source_file})</span>`;
    
    hideSection("detectedObjectsPanel"); 
    const dObjList = document.getElementById("detectedObjectsList"); 
    if (dObjList) dObjList.innerHTML = "";
    
    renderAvailableRecords(); 
    renderAvailableColumns(); 
    renderMappedItems();

    const imageEl = document.getElementById("imageToCrop");
    if (state.cropper) { 
        state.cropper.destroy(); 
        state.cropper = null; 
    }
    imageEl.src = "";
    
    imageEl.onload = () => {
        document.querySelectorAll(".tool-btn").forEach(b => b.classList.remove("active")); 
        document.getElementById("btnToolCrop").classList.add("active"); 
        
        setTimeout(() => {
            state.cropper = new Cropper(imageEl, {
                viewMode: 1,
                background: true,
                zoomable: true,
                wheelZoomRatio: 0.1,
                autoCrop: false,
                autoCropArea: 0,
                ready() {
                    const imgData = state.cropper.getImageData();
                    state.cmPerPx = ab.width_cm / imgData.naturalWidth;
                    // Force clear — บาง version ของ Cropper.js ignore autoCrop:false
                    state.cropper.clear();
                    if (state.pendingEditCrop) applyPendingEdit();
                },
                crop(event) {
                    if (state.cmPerPx > 0) {
                        const wEl = document.getElementById("cropWidthCm"); 
                        const hEl = document.getElementById("cropHeightCm");
                        if (document.activeElement !== wEl) wEl.value = (event.detail.width  * state.cmPerPx).toFixed(2);
                        if (document.activeElement !== hEl) hEl.value = (event.detail.height * state.cmPerPx).toFixed(2);
                    }
                    
                    if (state.crosshairEnabled && state.cropper) {
                        const box = state.cropper.getCropBoxData();
                        document.getElementById('crossLineH').style.display = 'block';
                        document.getElementById('crossLineV').style.display = 'block';
                        document.getElementById('crossLineH').style.top = (box.top + box.height / 2) + 'px';
                        document.getElementById('crossLineV').style.left = (box.left + box.width / 2) + 'px';
                    }
                },
            });
            // Modal is already centered — no scroll needed
            
            if (state.useAiMode) {
                runAutoDetect(ab);
            } else {
                document.getElementById("aiStatusBadge").classList.add("hidden");
                document.getElementById("detectedObjectsPanel").classList.add("hidden");
            }
        }, 150);
    };
    imageEl.src = ab.image_base64;
};

async function runAutoDetect(ab) {
    if (state.aiAbortController) state.aiAbortController.abort();
    state.aiAbortController = new AbortController();

    const panel = document.getElementById("detectedObjectsPanel");
    const list = document.getElementById("detectedObjectsList");
    const badge = document.getElementById("aiStatusBadge");
    
    badge.innerHTML = `<i class="fa-solid fa-spinner fa-spin mr-1"></i> AI กำลังสแกน...`;
    badge.classList.remove("hidden");
    list.innerHTML = ""; // เคลียร์ของเก่าก่อน
    
    try {
        const res = await fetch(ab.image_base64);
        const blob = await res.blob();
        const fd = new FormData();
        fd.append("file", blob, "artboard.png");
        
        // 1. ส่งรูปใหญ่ให้ OpenCV หาพิกัด Object เบื้องต้น
        const apiRes = await fetchWithAuth(`${API_BASE}/auto_detect`, { 
            method: "POST", 
            body: fd,
            signal: state.aiAbortController.signal 
        });
        const data = await apiRes.json();
        
        if (data.status === "success" && data.objects.length > 0) {
            panel.classList.remove("hidden");
            const img = new Image();
            img.src = ab.image_base64;
            await new Promise(r => img.onload = r);
            
            data.objects.forEach(obj => {
                // สร้าง Canvas ตัดรูปเฉพาะส่วนนั้น
                const canvas = document.createElement("canvas");
                canvas.width = obj.w; canvas.height = obj.h;
                const ctx = canvas.getContext("2d");
                ctx.drawImage(img, obj.x, obj.y, obj.w, obj.h, 0, 0, obj.w, obj.h);
                
                // สร้างปุ่ม UI
                const btn = document.createElement("div");
                btn.className = "relative cursor-pointer border-2 border-transparent hover:border-indigo-400 rounded bg-white p-1 shadow-sm shrink-0 transition-all hover:scale-105";
                
                const thumbImg = document.createElement("img");
                thumbImg.src = canvas.toDataURL();
                thumbImg.className = "h-10 object-contain";
                btn.appendChild(thumbImg);

                // ป้าย Badge โหลด AI Predict
                const predBadge = document.createElement("span");
                predBadge.className = "absolute -top-2 -right-2 bg-slate-200 text-slate-500 text-[8px] font-bold px-1.5 py-0.5 rounded shadow-sm";
                predBadge.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i>`;
                btn.appendChild(predBadge);
                
                list.appendChild(btn);

                // 🌟 2. ส่งรูปที่ตัดแล้วให้ ResNet18 ทำนายแบบ Asynchronous!
                canvas.toBlob(async (cropBlob) => {
                    try {
                        const predFd = new FormData();
                        predFd.append("file", cropBlob, "predict.png");
                        const predRes = await fetchWithAuth(`${API_BASE}/ai/predict`, { 
                            method: "POST", body: predFd, signal: state.aiAbortController.signal 
                        });
                        const pData = await predRes.json();

                        if (pData.status === "success" && pData.label !== "Unknown") {
                            const conf = Math.round(pData.confidence * 100);
                            // อัปเดต UI ถ้า AI เดาถูก
                            predBadge.className = "absolute -top-2 -right-2 bg-violet-600 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow-md border border-white";
                            predBadge.innerHTML = `${pData.label} <span class="text-[7px] text-violet-200">${conf}%</span>`;
                            
                            // เก็บข้อมูลไว้ใน dataset ของปุ่ม
                            btn.dataset.predictedCol = pData.label;
                            btn.dataset.predictedConf = conf;
                        } else {
                            predBadge.remove(); // ถ้า AI ไม่มั่นใจ (Unknown) ให้ซ่อนป้ายไป
                        }
                    } catch (e) {
                        predBadge.remove();
                    }
                }, "image/png");

                // 3. Event ตอน User กดรูป Thumbnail
                btn.onclick = () => {
                    if (state.cropper) {
                        // เลื่อนกรอบไปหาเป้าหมาย
                        state.cropper.setData({ x: obj.x, y: obj.y, width: obj.w, height: obj.h });
                        state.cropper.setDragMode("crop");
                        activateCropTool();

                        const pCol = btn.dataset.predictedCol;
                        if (pCol) {
                            // 🌟 AI Auto-Select Column!
                            const colBtn = document.querySelector(`.col-card[data-col="${pCol}"]`);
                            if (colBtn) {
                                document.querySelectorAll(".col-card").forEach(b => b.classList.remove("selected"));
                                colBtn.classList.add("selected");
                                showSnapAlert(`AI ตีกรอบ + เลือกคอลัมน์ ${pCol} ให้แล้ว! (${btn.dataset.predictedConf}%)`, 'item');
                            } else {
                                showSnapAlert('AI ดึงกรอบล้อมชิ้นส่วนอัตโนมัติ!', 'item');
                            }
                        } else {
                            showSnapAlert('AI ดึงกรอบล้อมชิ้นส่วนอัตโนมัติ!', 'item');
                        }
                    }
                };
            });
            badge.innerHTML = `<i class="fa-solid fa-robot mr-1 text-indigo-500"></i> เจอ ${data.objects.length} ชิ้นส่วน`;
        } else {
            badge.innerHTML = `<i class="fa-solid fa-check text-emerald-500 mr-1"></i> พร้อมทำงาน`;
        }
    } catch (e) {
        if (e.name === 'AbortError') return;
        badge.innerHTML = `<i class="fa-solid fa-check text-emerald-500 mr-1"></i> พร้อมทำงาน`;
    }
}

window.toggleAutoCrop = () => {
    if (!state.cropper) return;
    const btn = document.getElementById("btnAutoCrop");
    const isOn = btn.classList.contains("bg-amber-500");

    if (isOn) {
        // ปิด — clear กรอบออก
        state.cropper.clear();
        btn.classList.replace("bg-amber-500",    "bg-white");
        btn.classList.replace("text-white",      "text-slate-500");
        btn.classList.replace("border-amber-500","border-slate-200");
    } else {
        // เปิด — crop เต็มภาพ (autoCropArea = 1)
        state.cropper.crop();
        state.cropper.setCropBoxData({
            left:   0,
            top:    0,
            width:  state.cropper.getContainerData().width,
            height: state.cropper.getContainerData().height,
        });
        btn.classList.replace("bg-white",        "bg-amber-500");
        btn.classList.replace("text-slate-500",  "text-white");
        btn.classList.replace("border-slate-200","border-amber-500");
    }
};


window.toggleCrosshair = () => {
    state.crosshairEnabled = !state.crosshairEnabled;
    const btn = document.getElementById("btnCrosshair");
    const container = document.getElementById("cropper-container");
    
    if (state.crosshairEnabled) {
        btn.classList.replace("text-slate-500", "text-pink-600");
        btn.classList.replace("bg-white", "bg-pink-50");
        btn.classList.replace("border-slate-200", "border-pink-200");
        container.classList.add("crosshair-active");
        
        if(state.cropper) {
            const box = state.cropper.getCropBoxData();
            document.getElementById('crossLineH').style.display = 'block';
            document.getElementById('crossLineV').style.display = 'block';
            document.getElementById('crossLineH').style.top = (box.top + box.height / 2) + 'px';
            document.getElementById('crossLineV').style.left = (box.left + box.width / 2) + 'px';
        }
    } else {
        btn.classList.replace("text-pink-600", "text-slate-500");
        btn.classList.replace("bg-pink-50", "bg-white");
        btn.classList.replace("border-pink-200", "border-slate-200");
        container.classList.remove("crosshair-active");
        document.getElementById('crossLineH').style.display = 'none';
        document.getElementById('crossLineV').style.display = 'none';
    }
};

window.markCurrentArtboardDone = () => { 
    if (state.activeArtboardIndex !== null && !state.manualCompletedArtboards.includes(state.activeArtboardIndex)) {
        state.manualCompletedArtboards.push(state.activeArtboardIndex); 
    }
    closeCropperWorkspace(); 
};

window.closeCropperWorkspace = () => { 
    document.getElementById("cropModal").classList.add("hidden"); 
    document.body.style.overflow = ""; 
    renderArtboardGallery(); 
};

window.restoreLastCrop = async () => {
    const selectedColBtn = document.querySelector(".col-card.selected");
    if (!selectedColBtn) return await customAlert("💡 กรุณาเลือก 'Target Column' ก่อนกดปุ่มนี้\nระบบจะได้รู้ว่าต้องดึงตำแหน่งความจำของคอลัมน์ไหนมาครับ!", "warning");
    
    const col = selectedColBtn.dataset.col; 
    const memory = JSON.parse(localStorage.getItem(COL_MEMORY_KEY) || "{}");
    
    if (state.cropper && memory[col]) { 
        state.cropper.setData(memory[col]); 
        state.cropper.setDragMode("crop"); 
        activateCropTool(); 
    } else {
        await customAlert(`ยังไม่มีประวัติพิกัดการตัดรูปล่าสุดสำหรับคอลัมน์ [ ${col} ] ครับ`, "warning");
    }
};