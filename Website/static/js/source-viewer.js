(function () {
  "use strict";
  var dialog = document.querySelector(".imo-source-viewer");
  if (!dialog || typeof dialog.showModal !== "function") return;
  var code = dialog.querySelector("code");
  var pre = dialog.querySelector("pre");
  var copy = dialog.querySelector("[data-source-copy]");
  var download = dialog.querySelector("[data-source-download]");
  var status = dialog.querySelector("[data-source-status]");
  var request;
  var source = "";

  document.addEventListener("click", async function (event) {
    var link = event.target.closest("a[href]");
    if (!link || dialog.contains(link) || event.defaultPrevented || event.button !== 0 ||
        event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    var url = new URL(link.href, location.href);
    if (url.origin !== location.origin || !url.pathname.startsWith("/downloads/showcase/") ||
        !/\.(?:cs|source)\.txt$/i.test(url.pathname)) return;
    event.preventDefault();
    if (request) request.abort();
    var current = new AbortController();
    request = current;
    source = "";
    copy.disabled = true;
    code.textContent = "";
    var csharp = /\.cs\.txt$/i.test(url.pathname);
    var filename = decodeURIComponent(url.pathname.split("/").pop()).replace(/\.txt$/i, "");
    if (!csharp) filename = filename.replace(/\.source$/i, ".html");
    dialog.querySelector("[data-source-filename]").textContent = filename;
    download.href = url.href;
    download.download = filename;
    pre.className = code.className = "language-" + (csharp ? "csharp" : "markup");
    status.textContent = "Loading source…";
    dialog.showModal();
    pre.scrollTop = pre.scrollLeft = 0;
    try {
      var response = await fetch(url.href, { signal: current.signal });
      if (!response.ok) throw new Error("Source unavailable");
      var text = await response.text();
      if (current.signal.aborted) return;
      source = text;
      code.textContent = text;
      copy.disabled = false;
      status.textContent = text.split(/\r?\n/).length + " lines";
      if (window.Prism && typeof window.Prism.highlightElement === "function") {
        window.Prism.highlightElement(code);
      }
    } catch (error) {
      if (!current.signal.aborted) status.textContent = "Could not load the preview. Try Download source or reopen it.";
    }
  });

  dialog.querySelector("[data-source-close]").addEventListener("click", function () { dialog.close(); });
  dialog.addEventListener("click", function (event) {
    if (event.target !== dialog) return;
    var rect = dialog.getBoundingClientRect();
    if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) dialog.close();
  });
  dialog.addEventListener("close", function () {
    if (request) request.abort();
    source = "";
    copy.disabled = true;
  });
  copy.addEventListener("click", async function () {
    var current = request;
    try {
      await navigator.clipboard.writeText(source);
      if (dialog.open && request === current) status.textContent = "Code copied.";
    } catch (_) {
      if (dialog.open && request === current) status.textContent = "Copy unavailable. Select the code or download the source.";
    }
  });
})();
