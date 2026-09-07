document.addEventListener("DOMContentLoaded", function () {
    const prompt = document.getElementById("filePrompt");
    const promptCount = document.getElementById("promptCount");
    const fileType = document.getElementById("fileType");
    const fileName = document.getElementById("fileName");
    const extension = document.getElementById("extension");
    const createBtn = document.getElementById("createFileBtn");
    const createStatus = document.getElementById("createStatus");
    const resultCard = document.getElementById("resultCard");
    const resultTitle = document.getElementById("resultTitle");
    const resultDescription = document.getElementById("resultDescription");
    const downloadBtn = document.getElementById("downloadBtn");
    const previewBtn = document.getElementById("previewBtn");
    const toast = document.getElementById("toast");

    const extensions = {
        auto: ".txt", pdf: ".pdf", docx: ".docx", xlsx: ".xlsx", pptx: ".pptx",
        html: ".html", css: ".css", js: ".js", ts: ".ts", jsx: ".jsx", tsx: ".tsx",
        python: ".py", c: ".c", cpp: ".cpp", java: ".java", php: ".php", sh: ".sh", bat: ".bat",
        json: ".json", xml: ".xml", yaml: ".yaml", yml: ".yml", csv: ".csv", sql: ".sql",
        ini: ".ini", conf: ".conf", env: ".env", txt: ".txt", md: ".md", log: ".log", zip: ".zip"
    };

    function showToast(message) {
        if (!toast) return;
        toast.textContent = message;
        toast.classList.add("show");
        clearTimeout(showToast.timer);
        showToast.timer = setTimeout(() => toast.classList.remove("show"), 3000);
    }

    function updateExtension() {
        if (extension && fileType) extension.textContent = extensions[fileType.value] || ".txt";
    }

    if (prompt) {
        prompt.addEventListener("input", function () {
            if (promptCount) promptCount.textContent = `${prompt.value.length} / 4000`;
        });
    }
    if (fileType) fileType.addEventListener("change", updateExtension);
    updateExtension();

    if (!createBtn) return;
    let creating = false;

    createBtn.addEventListener("click", async function () {
        if (creating) return;
        const userPrompt = (prompt ? prompt.value : "").trim();
        if (!userPrompt) { showToast("Please describe what you want to create."); if (prompt) prompt.focus(); return; }

        creating = true;
        createBtn.classList.add("loading");
        createBtn.disabled = true;
        createBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i><span>Creating File...</span>';
        if (createStatus) createStatus.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Zyntra AI is creating your file...';
        if (resultCard) resultCard.hidden = true;

        try {
            const response = await fetch("/api/file-creator", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Accept": "application/json" },
                body: JSON.stringify({
                    prompt: userPrompt,
                    file_type: fileType ? fileType.value : "auto",
                    file_name: fileName && fileName.value.trim() ? fileName.value.trim() : "my_file",
                    output_format: document.getElementById("outputFormat")?.value || "standard",
                    auto_format: document.getElementById("autoFormat")?.checked ?? true,
                    ai_optimize: document.getElementById("aiOptimize")?.checked ?? true,
                    professional_layout: document.getElementById("professionalLayout")?.checked ?? true
                })
            });

            let data = {};
            try { data = await response.json(); } catch (_) { throw new Error("Server returned an invalid response."); }
            if (!response.ok || !data.success) throw new Error(data.error || "File creation failed.");

            if (resultCard) resultCard.hidden = false;
            if (resultTitle) resultTitle.textContent = "File Created Successfully";
            if (resultDescription) resultDescription.textContent = `${data.filename} • ${data.provider || "AI"}`;
            if (downloadBtn) { downloadBtn.href = data.download_url; downloadBtn.download = data.filename; }
            if (previewBtn) {
                previewBtn.onclick = function () {
                    if (data.preview_url) window.open(data.preview_url, "_blank", "noopener,noreferrer");
                    else showToast("Preview is not available for this file.");
                };
            }
            if (createStatus) createStatus.innerHTML = '<i class="fa-solid fa-circle-check"></i> File created successfully';
            showToast(`Created ${data.filename}`);
        } catch (error) {
            console.error(error);
            if (createStatus) createStatus.innerHTML = '<i class="fa-solid fa-circle-exclamation"></i> ' + error.message;
            showToast(error.message);
        } finally {
            creating = false;
            createBtn.classList.remove("loading");
            createBtn.disabled = false;
            createBtn.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i><span>Create File</span>';
        }
    });
});
