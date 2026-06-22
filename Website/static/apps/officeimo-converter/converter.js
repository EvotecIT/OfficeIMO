(function () {
  const roots = Array.from(document.querySelectorAll("[data-officeimo-converter-app]"));
  if (roots.length === 0) {
    return;
  }

  const supported = {
    docx: {
      label: "Word document",
      api: "OfficeIMO.Word.Pdf",
      status: "Partial",
      statusClass: "warn",
      note: "Basic DOCX browser fixtures passed. Rich Word files still need Unicode font embedding diagnostics."
    },
    xlsx: {
      label: "Excel workbook",
      api: "OfficeIMO.Excel.Pdf",
      status: "Proof passed",
      statusClass: "good",
      note: "Representative XLSX fixture produced PDF bytes in Blazor WebAssembly."
    },
    pptx: {
      label: "PowerPoint deck",
      api: "OfficeIMO.PowerPoint.Pdf",
      status: "Proof passed",
      statusClass: "good",
      note: "Representative PPTX fixture produced PDF bytes in Blazor WebAssembly."
    }
  };

  const state = {
    file: null,
    target: "pdf",
    quality: "balanced",
    preserveLayout: true,
    diagnostics: true,
    stage: "idle"
  };

  function icon(name) {
    const icons = {
      upload: '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 16V4m0 0 4.5 4.5M12 4 7.5 8.5M5 20h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>',
      file: '<svg width="30" height="30" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M14 3v5a2 2 0 0 0 2 2h5M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-9l-5-7Z" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg>',
      spark: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="m12 3 1.8 5.1L19 10l-5.2 1.9L12 17l-1.8-5.1L5 10l5.2-1.9L12 3Zm7 12 .9 2.4 2.1.8-2.1.8L19 21l-.9-2-2.1-.8 2.1-.8L19 15Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>',
      download: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 4v10m0 0 4-4m-4 4-4-4M5 20h14" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>'
    };
    return icons[name] || icons.spark;
  }

  function extension(file) {
    if (!file || !file.name || !file.name.includes(".")) {
      return "";
    }
    return file.name.split(".").pop().toLowerCase();
  }

  function formatBytes(bytes) {
    if (!bytes) {
      return "0 B";
    }
    const units = ["B", "KB", "MB", "GB"];
    let value = bytes;
    let unit = 0;
    while (value >= 1024 && unit < units.length - 1) {
      value /= 1024;
      unit += 1;
    }
    return `${value.toFixed(value >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`;
  }

  function getProfile() {
    const ext = extension(state.file);
    return supported[ext] || null;
  }

  function diagnostics() {
    const profile = getProfile();
    if (!state.file) {
      return [
        ["cold", "Waiting for a document", "Drop a DOCX, XLSX, or PPTX file to inspect the conversion path locally."],
        ["cold", "Browser privacy boundary", "Selected files stay in this tab. The static app does not upload document contents."],
        ["warn", "WASM engine not bundled", "This shell is ready for the Blazor WebAssembly engine, but the converter payload is not published yet."]
      ];
    }

    const ext = extension(state.file);
    if (!profile) {
      return [
        ["bad", "Unsupported source format", `.${ext || "unknown"} is not in the initial browser conversion matrix.`],
        ["good", "File stayed local", "The browser only read metadata for this readiness report."],
        ["warn", "Next step", "Route this through local MCP/CLI conversion or add the format to the OfficeIMO browser support matrix."]
      ];
    }

    const rows = [
      ["good", `${profile.label} recognized`, `${profile.api} is the expected OfficeIMO conversion package.`],
      [profile.statusClass, profile.status, profile.note],
      ["good", "Static hosting compatible", "The planned path uses byte and stream APIs with no Office, LibreOffice, Redis, or server process."]
    ];

    if (state.file.size > 25 * 1024 * 1024) {
      rows.push(["warn", "Large browser workload", "Files over 25 MB should be tested against browser memory budgets before public release."]);
    }

    if (ext === "docx") {
      rows.push(["warn", "Word font coverage", "Private-use bullet glyphs and richer typography need the Unicode font embedding path."]);
    }

    return rows;
  }

  function buildReport() {
    const profile = getProfile();
    return {
      generatedAt: new Date().toISOString(),
      file: state.file ? {
        name: state.file.name,
        size: state.file.size,
        type: state.file.type || null,
        extension: extension(state.file)
      } : null,
      target: state.target,
      quality: state.quality,
      preserveLayout: state.preserveLayout,
      diagnostics: state.diagnostics,
      profile: profile ? {
        label: profile.label,
        api: profile.api,
        status: profile.status,
        note: profile.note
      } : null,
      engine: {
        status: "not-bundled",
        expectedMount: "/apps/officeimo-converter/",
        expectedRuntime: "Blazor WebAssembly"
      },
      checks: diagnostics().map(([level, title, detail]) => ({ level, title, detail }))
    };
  }

  function downloadReport() {
    const report = buildReport();
    const name = state.file ? state.file.name.replace(/\.[^.]+$/, "") : "officeimo-conversion";
    const blob = new Blob([JSON.stringify(report, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${name}-readiness.json`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function selectSample(kind) {
    const samples = {
      docx: { name: "quarterly-report.docx", size: 184320, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      xlsx: { name: "regional-sales.xlsx", size: 327680, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      pptx: { name: "board-update.pptx", size: 512000, type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
    };
    state.file = samples[kind] || samples.docx;
    state.stage = "selected";
  }

  function apiSample() {
    const profile = getProfile();
    const api = profile ? profile.api : "OfficeIMO.Word.Pdf";
    const load = extension(state.file) === "xlsx"
      ? "ExcelDocument.Load(inputStream)"
      : extension(state.file) === "pptx"
        ? "PowerPointPresentation.Load(inputStream)"
        : "WordDocument.Load(inputStream)";
    return `using ${api};\n\nawait using var inputStream = file.OpenReadStream(maxAllowedSize);\nusing var document = ${load};\nbyte[] pdfBytes = document.SaveAsPdfBytes();\nawait browser.DownloadAsync(pdfBytes, \"converted.pdf\");`;
  }

  function render(root) {
    const standalone = root.dataset.mode === "standalone";
    const profile = getProfile();
    const ext = extension(state.file);
    const statusLabel = profile ? profile.status : (state.file ? "Unsupported" : "Waiting");
    const statusClass = profile ? profile.statusClass : (state.file ? "bad" : "cold");
    const runDisabled = !state.file || !profile;

    root.innerHTML = `
      <section class="ocx-shell">
        <div class="ocx-container">
          ${standalone ? `
          <div class="ocx-topbar">
            <a class="ocx-brand" href="/">
              <span class="ocx-brand-mark">${icon("spark")}</span>
              <span>OfficeIMO Converter</span>
            </a>
            <div class="ocx-topbar-links">
              <a class="ocx-link" href="/playground/">Playground</a>
              <a class="ocx-link" href="/docs/converters/browser-playground/">Docs</a>
            </div>
          </div>` : ""}
          <div class="ocx-hero">
            <div>
              <h1 class="ocx-title">Convert Office files in the browser.</h1>
              <p class="ocx-lede">Drop a DOCX, XLSX, or PPTX file to inspect the OfficeIMO WebAssembly conversion path, get local diagnostics, and prepare a PDF conversion without uploading the document.</p>
              <div class="ocx-hero-actions">
                <a class="ocx-button ocx-button--primary" href="#converter-workbench">${icon("upload")} Choose a file</a>
                <a class="ocx-button" href="/docs/converters/browser-playground/">Read browser docs</a>
              </div>
            </div>
            <div class="ocx-status-strip" aria-label="Conversion status">
              <div class="ocx-status-card">
                <span class="ocx-status-label">Engine</span>
                <span class="ocx-status-value">WASM-ready shell</span>
              </div>
              <div class="ocx-status-card">
                <span class="ocx-status-label">Privacy</span>
                <span class="ocx-status-value">Local file only</span>
              </div>
              <div class="ocx-status-card">
                <span class="ocx-status-label">Selected</span>
                <span class="ocx-status-value">${state.file ? ext.toUpperCase() : "None yet"}</span>
              </div>
            </div>
          </div>

          <div id="converter-workbench" class="ocx-workbench">
            <div class="ocx-panel">
              <h2>Configuration</h2>
              <div class="ocx-field">
                <span class="ocx-label">Source document</span>
                <label class="ocx-dropzone" data-dropzone>
                  <input type="file" data-file accept=".docx,.xlsx,.pptx" />
                  <span class="ocx-dropzone-inner">
                    <span class="ocx-drop-icon">${icon("upload")}</span>
                    <strong>${state.file ? state.file.name : "Drop an Office file here"}</strong>
                    <span class="ocx-muted">${state.file ? `${formatBytes(state.file.size)} selected` : "or click to select DOCX, XLSX, or PPTX"}</span>
                  </span>
                </label>
                <p class="ocx-hint">The app reads metadata and prepares a readiness report now. The Blazor WASM converter can bind to this same surface when the engine bundle is published.</p>
              </div>

              <div class="ocx-field">
                <span class="ocx-label">Try a sample</span>
                <div class="ocx-sample-grid">
                  <button type="button" data-sample="docx">DOCX report</button>
                  <button type="button" data-sample="xlsx">XLSX workbook</button>
                  <button type="button" data-sample="pptx">PPTX deck</button>
                </div>
              </div>

              <div class="ocx-field">
                <span class="ocx-label">Output</span>
                <div class="ocx-segmented" role="group" aria-label="Output format">
                  ${["pdf", "html", "markdown"].map(target => `<button type="button" data-target="${target}" class="${state.target === target ? "is-selected" : ""}">${target.toUpperCase()}</button>`).join("")}
                </div>
              </div>

              <div class="ocx-field">
                <label class="ocx-label" for="ocx-quality">Conversion profile</label>
                <select id="ocx-quality" data-quality>
                  <option value="balanced" ${state.quality === "balanced" ? "selected" : ""}>Balanced fidelity</option>
                  <option value="fast" ${state.quality === "fast" ? "selected" : ""}>Fast preview</option>
                  <option value="strict" ${state.quality === "strict" ? "selected" : ""}>Strict diagnostics</option>
                </select>
              </div>

              <div class="ocx-field">
                <span class="ocx-label">Options</span>
                <div class="ocx-checks">
                  <label><input type="checkbox" data-preserve ${state.preserveLayout ? "checked" : ""} /> Preserve layout where the engine supports it</label>
                  <label><input type="checkbox" data-diagnostics ${state.diagnostics ? "checked" : ""} /> Include browser safety diagnostics</label>
                </div>
              </div>

              <button type="button" class="ocx-button ocx-button--primary" data-run ${runDisabled ? "disabled" : ""}>${icon("spark")} Run readiness check</button>
              <button type="button" class="ocx-button" data-report>${icon("download")} Download readiness report</button>
            </div>

            <div class="ocx-preview-column">
              <div class="ocx-output">
                <h2>Preview</h2>
                <div class="ocx-output-stage">
                  <div class="ocx-file-card">
                    <div style="color: var(--imo-accent, #38bdf8); margin-bottom: 0.85rem;">${icon("file")}</div>
                    <p class="ocx-file-title">${state.file ? state.file.name : "No document selected"}</p>
                    <p class="ocx-muted">${state.file ? `${formatBytes(state.file.size)} • ${profile ? profile.label : "unknown format"}` : "Choose a file to see the conversion route, expected package, and browser-specific warnings."}</p>
                    <div class="ocx-file-meta">
                      <span class="ocx-chip ocx-chip--${statusClass}">${statusLabel}</span>
                      <span class="ocx-chip">Target: ${state.target.toUpperCase()}</span>
                      <span class="ocx-chip">${state.quality}</span>
                    </div>
                  </div>
                </div>
                <div class="ocx-diagnostics">
                  ${diagnostics().map(([level, title, detail]) => `
                    <div class="ocx-diagnostic">
                      <span class="ocx-dot ocx-dot--${level}"></span>
                      <span><strong>${title}</strong><span>${detail}</span></span>
                    </div>`).join("")}
                </div>
                <div class="ocx-log" aria-live="polite">
                  <span>&gt; ${state.stage === "checked" ? "Readiness check completed." : "Waiting for local file selection."}</span>
                  <span>&gt; Engine payload: Blazor WebAssembly bundle pending.</span>
                  <span>&gt; Upload policy: none.</span>
                </div>
              </div>
              <div class="ocx-code-card">
                <h2>Core API shape</h2>
                <pre><code>${escapeHtml(apiSample())}</code></pre>
              </div>
              <div class="ocx-matrix">
                ${Object.entries(supported).map(([key, value]) => `
                  <div class="ocx-matrix-item">
                    <strong>${key.toUpperCase()} to PDF</strong>
                    <span>${value.status}. ${value.api}</span>
                  </div>`).join("")}
              </div>
            </div>
          </div>
        </div>
      </section>
    `;

    wire(root);
  }

  function escapeHtml(value) {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function wire(root) {
    const fileInput = root.querySelector("[data-file]");
    const dropzone = root.querySelector("[data-dropzone]");

    fileInput?.addEventListener("change", (event) => {
      state.file = event.target.files && event.target.files[0] ? event.target.files[0] : null;
      state.stage = "selected";
      render(root);
    });

    if (dropzone) {
      ["dragenter", "dragover"].forEach(name => {
        dropzone.addEventListener(name, (event) => {
          event.preventDefault();
          dropzone.classList.add("is-active");
        });
      });
      ["dragleave", "drop"].forEach(name => {
        dropzone.addEventListener(name, (event) => {
          event.preventDefault();
          dropzone.classList.remove("is-active");
        });
      });
      dropzone.addEventListener("drop", (event) => {
        const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
        if (file) {
          state.file = file;
          state.stage = "selected";
          render(root);
        }
      });
    }

    root.querySelectorAll("[data-target]").forEach(button => {
      button.addEventListener("click", () => {
        state.target = button.dataset.target;
        state.stage = "selected";
        render(root);
      });
    });

    root.querySelectorAll("[data-sample]").forEach(button => {
      button.addEventListener("click", () => {
        selectSample(button.dataset.sample);
        render(root);
      });
    });

    root.querySelector("[data-quality]")?.addEventListener("change", (event) => {
      state.quality = event.target.value;
      render(root);
    });

    root.querySelector("[data-preserve]")?.addEventListener("change", (event) => {
      state.preserveLayout = event.target.checked;
      render(root);
    });

    root.querySelector("[data-diagnostics]")?.addEventListener("change", (event) => {
      state.diagnostics = event.target.checked;
      render(root);
    });

    root.querySelector("[data-run]")?.addEventListener("click", () => {
      state.stage = "checked";
      render(root);
    });

    root.querySelector("[data-report]")?.addEventListener("click", downloadReport);
  }

  roots.forEach(render);
})();
