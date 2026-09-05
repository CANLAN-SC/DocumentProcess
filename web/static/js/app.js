const state = {
  files: [],
  selectedIndex: -1,
  previewCache: new Map(),
  previewRequest: null,
};

const elements = {
  fileInput: document.querySelector("#file-input"),
  dropzone: document.querySelector("#dropzone"),
  fileList: document.querySelector("#file-list"),
  imageScale: document.querySelector("#image-scale"),
  imageScaleOutput: document.querySelector("#image-scale-output"),
  marginTop: document.querySelector("#margin-top"),
  marginBottom: document.querySelector("#margin-bottom"),
  marginLeft: document.querySelector("#margin-left"),
  marginRight: document.querySelector("#margin-right"),
  resetLayout: document.querySelector("#reset-layout"),
  mergeFiles: document.querySelector("#merge-files"),
  convertButton: document.querySelector("#convert-button"),
  previewTitle: document.querySelector("#preview-title"),
  previewMeta: document.querySelector("#preview-meta"),
  emptyState: document.querySelector("#empty-state"),
  loadingState: document.querySelector("#loading-state"),
  pages: document.querySelector("#pages"),
  resultCard: document.querySelector("#result-card"),
  errorMessage: document.querySelector("#error-message"),
};

const allowedExtensions = new Set(["pdf", "jpg", "jpeg", "png"]);

function fileKey(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}

function extensionOf(name) {
  return name.includes(".") ? name.split(".").pop().toLowerCase() : "";
}

function addFiles(fileList) {
  const existing = new Set(state.files.map(fileKey));
  const rejected = [];
  for (const file of fileList) {
    if (!allowedExtensions.has(extensionOf(file.name))) {
      rejected.push(file.name);
      continue;
    }
    if (!existing.has(fileKey(file)) && state.files.length < 50) {
      state.files.push(file);
      existing.add(fileKey(file));
    }
  }
  if (rejected.length) showError(`已忽略不支持的文件：${rejected.join("、")}`);
  // 按文件名字典序（升序、自然排序）重排，保证列表顺序稳定可预期。
  // 若已有选中文件，按文件标识重新定位，避免排序后选中项错位。
  const selectedFile = state.selectedIndex >= 0 ? state.files[state.selectedIndex] : null;
  state.files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
  if (selectedFile) {
    const located = state.files.findIndex((file) => fileKey(file) === fileKey(selectedFile));
    state.selectedIndex = located >= 0 ? located : 0;
  } else if (state.files.length) {
    state.selectedIndex = 0;
  }
  renderFileList();
  updateControls();
  loadSelectedPreview();
}

function renderFileList() {
  elements.fileList.replaceChildren();
  state.files.forEach((file, index) => {
    const row = document.createElement("div");
    row.className = `file-item${index === state.selectedIndex ? " active" : ""}`;
    row.tabIndex = 0;
    row.setAttribute("role", "button");
    row.setAttribute("aria-label", `预览 ${file.name}`);
    const extension = extensionOf(file.name);
    const badge = document.createElement("span");
    badge.className = `file-badge${extension === "pdf" ? " pdf" : ""}`;
    badge.textContent = extension === "pdf" ? "PDF" : "IMG";
    const copy = document.createElement("span");
    copy.className = "file-copy";
    const name = document.createElement("strong");
    name.textContent = file.name;
    const size = document.createElement("small");
    size.textContent = formatBytes(file.size);
    copy.append(name, size);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove-file";
    remove.textContent = "×";
    remove.setAttribute("aria-label", `移除 ${file.name}`);
    remove.addEventListener("click", (event) => {
      event.stopPropagation();
      removeFile(index);
    });
    row.append(badge, copy, remove);
    row.addEventListener("click", () => selectFile(index));
    row.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectFile(index);
      }
    });
    elements.fileList.append(row);
  });
}

function removeFile(index) {
  const [removed] = state.files.splice(index, 1);
  state.previewCache.delete(fileKey(removed));
  if (!state.files.length) state.selectedIndex = -1;
  else if (state.selectedIndex >= state.files.length) state.selectedIndex = state.files.length - 1;
  else if (index < state.selectedIndex) state.selectedIndex -= 1;
  renderFileList();
  updateControls();
  loadSelectedPreview();
}

function selectFile(index) {
  if (index === state.selectedIndex) return;
  state.selectedIndex = index;
  renderFileList();
  loadSelectedPreview();
}

function updateControls() {
  elements.convertButton.disabled = state.files.length === 0;
  elements.mergeFiles.disabled = state.files.length < 2;
  if (state.files.length < 2) elements.mergeFiles.checked = false;
}

function showPreviewState(mode) {
  elements.emptyState.hidden = mode !== "empty";
  elements.loadingState.hidden = mode !== "loading";
  elements.pages.hidden = mode !== "pages";
}

async function loadSelectedPreview() {
  clearError();
  if (state.previewRequest) state.previewRequest.abort();
  if (state.selectedIndex < 0) {
    elements.previewTitle.textContent = "Word 页面";
    elements.previewMeta.textContent = "等待添加文件";
    elements.pages.replaceChildren();
    showPreviewState("empty");
    return;
  }
  const file = state.files[state.selectedIndex];
  elements.previewTitle.textContent = file.name.replace(/\.[^.]+$/, "");
  const cached = state.previewCache.get(fileKey(file));
  if (cached) {
    renderPreview(cached);
    return;
  }

  showPreviewState("loading");
  elements.previewMeta.textContent = "正在生成预览";
  const controller = new AbortController();
  state.previewRequest = controller;
  const form = new FormData();
  form.append("file", file);
  try {
    const response = await fetch("/api/preview", { method: "POST", body: form, signal: controller.signal });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "无法生成预览");
    state.previewCache.set(fileKey(file), data);
    if (file === state.files[state.selectedIndex]) renderPreview(data);
  } catch (error) {
    if (error.name === "AbortError") return;
    showPreviewState("empty");
    elements.previewMeta.textContent = "预览失败";
    showError(error.message);
  } finally {
    if (state.previewRequest === controller) state.previewRequest = null;
  }
}

function renderPreview(data) {
  elements.pages.replaceChildren();
  for (const page of data.pages) {
    const paper = document.createElement("article");
    paper.className = "word-page";
    paper.innerHTML = `
      <div class="word-content">
        <h3 class="word-title"></h3>
        <div class="word-image-area"><img alt="第 ${page.page} 页预览"></div>
      </div>
      <span class="page-number">${page.page}</span>`;
    paper.querySelector(".word-title").textContent = data.title;
    const image = paper.querySelector("img");
    image.src = page.url;
    image.width = page.width;
    image.height = page.height;
    elements.pages.append(paper);
  }
  const suffix = data.total_pages > data.preview_limit ? `，显示前 ${data.preview_limit} 页` : "";
  elements.previewMeta.textContent = `共 ${data.total_pages} 页${suffix}`;
  showPreviewState("pages");
  applyLayoutPreview();
}

function numericValue(input, fallback) {
  const value = Number.parseFloat(input.value);
  return Number.isFinite(value) ? value : fallback;
}

function applyLayoutPreview() {
  const scale = numericValue(elements.imageScale, 90);
  elements.imageScaleOutput.textContent = `${Math.round(scale)}%`;
  const values = {
    "--margin-top": `${numericValue(elements.marginTop, 2.54) / 29.7 * 100}%`,
    "--margin-bottom": `${numericValue(elements.marginBottom, 2.54) / 29.7 * 100}%`,
    "--margin-left": `${numericValue(elements.marginLeft, 2.54) / 21 * 100}%`,
    "--margin-right": `${numericValue(elements.marginRight, 2.54) / 21 * 100}%`,
    "--image-scale": `${scale}%`,
  };
  for (const page of elements.pages.children) {
    Object.entries(values).forEach(([property, value]) => page.style.setProperty(property, value));
  }
}

function validateSettings() {
  const scale = numericValue(elements.imageScale, NaN);
  const margins = [elements.marginTop, elements.marginBottom, elements.marginLeft, elements.marginRight]
    .map((input) => numericValue(input, NaN));
  if (!Number.isFinite(scale) || scale < 10 || scale > 200) return "图片大小必须在 10% 到 200% 之间";
  if (margins.some((value) => !Number.isFinite(value) || value < 0 || value > 8)) return "页边距必须在 0 到 8 厘米之间";
  if (margins[2] + margins[3] >= 20) return "左右页边距之和过大";
  if (margins[0] + margins[1] >= 28.35) return "上下页边距之和过大";
  return null;
}

async function convertFiles() {
  clearError();
  elements.resultCard.hidden = true;
  const settingsError = validateSettings();
  if (settingsError) {
    showError(settingsError);
    return;
  }
  const form = new FormData();
  state.files.forEach((file) => form.append("files", file));
  form.append("image_scale", elements.imageScale.value);
  form.append("margin_top", elements.marginTop.value);
  form.append("margin_bottom", elements.marginBottom.value);
  form.append("margin_left", elements.marginLeft.value);
  form.append("margin_right", elements.marginRight.value);
  form.append("quality", document.querySelector('input[name="quality"]:checked').value);
  form.append("merge", elements.mergeFiles.checked ? "true" : "false");

  elements.convertButton.disabled = true;
  elements.convertButton.classList.add("loading");
  try {
    const response = await fetch("/api/convert", { method: "POST", body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "转换失败，请重试");
    elements.resultCard.replaceChildren();
    const message = document.createElement("strong");
    message.textContent = `转换完成 · ${data.count} 个文件`;
    const detail = document.createElement("div");
    detail.textContent = data.kind === "merged" ? "文档已合并，可以下载。" : "文件已准备好，可以下载。";
    const link = document.createElement("a");
    link.className = "download-link";
    link.href = data.download_url;
    link.textContent = `下载 ${data.filename}`;
    elements.resultCard.append(message, detail, link);
    elements.resultCard.hidden = false;
  } catch (error) {
    showError(error.message);
  } finally {
    elements.convertButton.classList.remove("loading");
    elements.convertButton.disabled = state.files.length === 0;
  }
}

function showError(message) {
  elements.errorMessage.textContent = message;
  elements.errorMessage.hidden = false;
}

function clearError() {
  elements.errorMessage.hidden = true;
  elements.errorMessage.textContent = "";
}

elements.fileInput.addEventListener("change", () => {
  addFiles(elements.fileInput.files);
  elements.fileInput.value = "";
});

for (const eventName of ["dragenter", "dragover"]) {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("dragging");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("dragging");
  });
}
elements.dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

for (const input of [elements.imageScale, elements.marginTop, elements.marginBottom, elements.marginLeft, elements.marginRight]) {
  input.addEventListener("input", applyLayoutPreview);
}

elements.resetLayout.addEventListener("click", () => {
  elements.imageScale.value = "90";
  elements.marginTop.value = "2.54";
  elements.marginBottom.value = "2.54";
  elements.marginLeft.value = "2.54";
  elements.marginRight.value = "2.54";
  applyLayoutPreview();
});

elements.convertButton.addEventListener("click", convertFiles);
