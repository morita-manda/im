(function () {
  const form = document.getElementById("im-form");
  const submitBtn = document.getElementById("submit-btn");
  const btnText = document.getElementById("btn-text");
  const btnLoading = document.getElementById("btn-loading");
  const progressArea = document.getElementById("progress-area");
  const progressMsg = document.getElementById("progress-msg");
  const progressFill = document.getElementById("progress-fill");
  const errorArea = document.getElementById("error-area");
  const errorMsg = document.getElementById("error-msg");
  const uploadArea = document.getElementById("upload-area");
  const fileInput = document.getElementById("pdfs");
  const fileList = document.getElementById("file-list");

  let selectedFiles = [];

  // ========== 株主・役員 動的行 ==========
  function createRow(container, fields, placeholder) {
    const row = document.createElement("div");
    row.className = "dynamic-row";
    fields.forEach(({ cls, ph }) => {
      const inp = document.createElement("input");
      inp.type = "text";
      inp.className = cls;
      inp.placeholder = ph;
      row.appendChild(inp);
    });
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "row-remove";
    btn.textContent = "✕";
    btn.addEventListener("click", () => {
      if (container.children.length > 1) row.remove();
    });
    row.appendChild(btn);
    container.appendChild(row);
  }

  function addShareholderRow() {
    createRow(document.getElementById("shareholder-list"), [
      { cls: "field-main", ph: "例：山田太郎" },
      { cls: "field-sub",  ph: "例：60%" },
    ]);
  }

  // 初期行を追加
  addShareholderRow();

  document.getElementById("add-shareholder").addEventListener("click", addShareholderRow);

  function collectRows(containerId, keys) {
    const rows = document.getElementById(containerId).querySelectorAll(".dynamic-row");
    const result = [];
    rows.forEach((row) => {
      const inputs = row.querySelectorAll("input");
      const obj = {};
      keys.forEach((k, i) => { obj[k] = inputs[i] ? inputs[i].value.trim() : ""; });
      if (Object.values(obj).some((v) => v)) result.push(obj);
    });
    return result;
  }

  // ========== ドラッグ&ドロップ ==========
  uploadArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadArea.classList.add("drag-over");
  });
  uploadArea.addEventListener("dragleave", () => uploadArea.classList.remove("drag-over"));
  uploadArea.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadArea.classList.remove("drag-over");
    const files = Array.from(e.dataTransfer.files).filter((f) => f.type === "application/pdf");
    addFiles(files);
  });
  uploadArea.addEventListener("click", (e) => {
    if (e.target.tagName !== "LABEL" && e.target.tagName !== "INPUT") {
      fileInput.click();
    }
  });
  fileInput.addEventListener("change", () => {
    addFiles(Array.from(fileInput.files));
    fileInput.value = "";
  });

  function addFiles(files) {
    files.forEach((f) => {
      if (selectedFiles.length >= 3) return;
      if (!selectedFiles.find((x) => x.name === f.name && x.size === f.size)) {
        selectedFiles.push(f);
      }
    });
    renderFileList();
  }

  function renderFileList() {
    fileList.innerHTML = "";
    selectedFiles.forEach((f, i) => {
      const item = document.createElement("div");
      item.className = "file-item";
      item.innerHTML = `
        <span>📄</span>
        <span class="file-name">${f.name} (${(f.size / 1024).toFixed(0)}KB)</span>
        <span class="file-remove" data-i="${i}" title="削除">✕</span>
      `;
      fileList.appendChild(item);
    });
    fileList.querySelectorAll(".file-remove").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const idx = parseInt(e.target.dataset.i);
        selectedFiles.splice(idx, 1);
        renderFileList();
      });
    });
  }

  // ========== フォーム送信 ==========
  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const url = document.getElementById("url").value.trim();
    if (!url) {
      alert("URLを入力してください");
      return;
    }

    // UI: 生成中
    form.style.display = "none";
    errorArea.style.display = "none";
    progressArea.style.display = "block";
    startProgress();

    const formData = new FormData();
    formData.append("url", url);
    formData.append("companyName", document.getElementById("companyName").value.trim());
    formData.append("repName", document.getElementById("repName").value.trim());
    formData.append("reason", document.getElementById("reason").value.trim());
    formData.append("price", document.getElementById("price").value.trim());
    formData.append("scheme", document.getElementById("scheme").value);
    formData.append("managementIntent", document.getElementById("managementIntent").value);
    formData.append("shareholdersJson", JSON.stringify(collectRows("shareholder-list", ["name", "ratio"])));
    formData.append("empFull", document.getElementById("empFull").value.trim());
    formData.append("empPart", document.getElementById("empPart").value.trim());
    formData.append("author", document.getElementById("author").value.trim());
    selectedFiles.forEach((f) => formData.append("pdfs", f));

    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: "不明なエラー" }));
        throw new Error(err.error || `HTTPエラー: ${res.status}`);
      }

      // ダウンロード
      const blob = await res.blob();
      const disposition = res.headers.get("Content-Disposition") || "";
      let filename = "企業概要書.pptx";
      const fnMatch = disposition.match(/filename\*=UTF-8''(.+)/) || disposition.match(/filename="?([^"]+)"?/);
      if (fnMatch) filename = decodeURIComponent(fnMatch[1]);

      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();

      // 完了画面を表示
      stopProgress(true);
      setTimeout(() => {
        progressArea.style.display = "none";
        document.getElementById("success-filename").textContent = filename;
        document.getElementById("success-area").style.display = "block";
      }, 800);
    } catch (err) {
      stopProgress(false);
      progressArea.style.display = "none";
      errorArea.style.display = "block";
      errorMsg.textContent = err.message;
      form.style.display = "block";
    }
  });

  // ========== プログレス演出 ==========
  let progressTimer = null;
  const steps = [
    { pct: 10, msg: "HPの情報を取得しています..." },
    { pct: 30, msg: "決算書PDFを解析しています..." },
    { pct: 55, msg: "Claude AIでデータを構造化しています..." },
    { pct: 80, msg: "PPTXスライドを生成しています..." },
    { pct: 95, msg: "ファイルを仕上げています..." },
  ];
  let stepIdx = 0;

  function startProgress() {
    stepIdx = 0;
    setProgress(steps[0]);
    progressTimer = setInterval(() => {
      stepIdx = Math.min(stepIdx + 1, steps.length - 1);
      setProgress(steps[stepIdx]);
    }, 15000);
  }

  function setProgress(step) {
    progressFill.style.width = step.pct + "%";
    progressMsg.textContent = step.msg;
  }

  function stopProgress(success) {
    clearInterval(progressTimer);
    if (success) {
      progressFill.style.width = "100%";
      progressMsg.textContent = "✅ 生成完了！ダウンロードが始まります。";
    }
  }

  window.resetForm = function () {
    errorArea.style.display = "none";
    document.getElementById("success-area").style.display = "none";
    form.reset();
    selectedFiles = [];
    renderFileList();
    document.getElementById("shareholder-list").innerHTML = "";
    addShareholderRow();
    form.style.display = "block";
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
})();
