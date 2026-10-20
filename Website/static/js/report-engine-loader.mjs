try {
  await import("./report-engine.mjs");
} catch {
  document.getElementById("report-status").textContent = "The report controls could not load. Reload the controls or open the JavaScript API guide below.";
  const reload = document.getElementById("report-generate");
  reload.type = "button"; reload.textContent = "Reload controls"; reload.disabled = false;
  reload.addEventListener("click", () => location.reload());
}
