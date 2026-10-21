// Browser policy stays in this opt-in verifier; shaping and PDF writing use their public owners.
globalThis.runShapingContracts = async function ({ moduleBase, shapingBase, fontBase }) {
  const check = (condition, message) => { if (!condition) throw new Error(message); };
  const hash = async bytes => Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), byte => byte.toString(16).padStart(2, "0")).join("");
  const encodedHash = value => hash(new TextEncoder().encode(JSON.stringify(value)));
  const get = async url => {
    const response = await fetch(url);
    check(response.ok, "Shaping fixture fetch failed: " + response.status + " " + url);
    return new Uint8Array(await response.arrayBuffer());
  };
  const api = await import(moduleBase + "/pdf/index.js");
  const { createHarfBuzzTextShaper } = await import(moduleBase + "/pdf/providers/harfbuzz.js");
  const { corpus, createReport, controlRequests, normalizeGlyphRuns } = await import(shapingBase + "/corpus.mjs");
  const hb = await import(shapingBase + "/packages/harfbuzzjs/package/dist/index.mjs");
  const { default: bidiFactory } = await import(shapingBase + "/packages/bidi-js/package/dist/bidi.mjs");
  const latinBytes = await get(fontBase + "/Carlito-Regular.ttf"), arabicBytes = await get(fontBase + "/NotoSansArabic-Regular.ttf");
  const provider = createHarfBuzzTextShaper({ harfbuzz: hb, bidi: bidiFactory() });
  let calls = 0;
  const lineRequests = [];
  const recording = { shapeLine(request) {
    calls++;
    lineRequests.push({ text: request.text, start: request.start, end: request.end, fonts: request.runs.length, direction: request.direction });
    return provider.shapeLine(request);
  } };
  const report = await createReport(api, recording, latinBytes, arabicBytes);
  const controls = [];
  for (const request of controlRequests(report.latin, report.arabic)) {
    const normalized = normalizeGlyphRuns(await provider.shapeLine(request), { latin: report.latin, arabic: report.arabic });
    check(normalized.length > 0 && normalized.every(glyph => glyph[1] > 0), "Shaping control returned missing glyphs.");
    controls.push({ name: request.name, normalized, sha256: await encodedHash(normalized) });
  }
  check(controls.some(control => control.name === "arabic" && control.normalized.some(glyph => glyph[4] === 0 && (glyph[6] !== 0 || glyph[7] !== 0))), "Arabic control lost positioned combining marks.");
  check(lineRequests.some(request => request.text === corpus.mixed && (request.start > 0 || request.end < request.text.length)), "Mixed text did not exercise core-owned line wrapping.");

  let defaultRejected = false;
  try { await api.writePdfDocument([{ kind: "paragraph", text: [{ text: corpus.arabic, font: report.arabic }] }], { fonts: { regular: report.latin } }); }
  catch (error) { defaultRejected = error?.code === "NOT_SUPPORTED"; }
  check(defaultRejected, "Arabic shaping became an implicit default dependency or silent scalar fallback.");

  const bytes = new Uint8Array(await report.pdf.arrayBuffer());
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 32768) binary += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
  const contract = { corpus, controls, calls, finalizedLineRequests: lineRequests.filter(request => request.start > 0 || request.end < request.text.length) };
  await saveShapingFixture("shaped-report.pdf", btoa(binary), JSON.stringify(contract));
  const pdfSha256 = await hash(bytes);
  provider.dispose();
  let disposedRejected = false;
  try { provider.shapeLine(controlRequests(report.latin, report.arabic)[0]); } catch (error) { disposedRejected = error instanceof TypeError; }
  check(disposedRejected, "Disposed text provider still accepts native shaping work.");
  return { harfbuzzVersion: hb.versionString(), controls, pdfSha256, pdfBytes: bytes.length, calls,
    defaultWithoutProviderRejected: defaultRejected, disposedRejected,
    explicitFonts: { latin: await hash(latinBytes), arabic: await hash(arabicBytes) } };
};
