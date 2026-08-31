using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlManagedRendererGallery(string folderPath, bool openPdf) {
            Console.WriteLine("[*] Managed HTML renderer gallery: one source to interactive PDF, PNG, and SVG");

            const string html = """
                <!doctype html>
                <html lang="en">
                <head>
                  <meta charset="utf-8">
                  <style>
                    @layer reset, theme, components;
                    @page renderer-gallery {
                      size: A4 landscape;
                      margin: 18px;
                      @bottom-left { content: "OFFICEIMO / MANAGED RENDERER"; color:#64748b; font-size:8px; }
                      @bottom-right { content: "PAGE " counter(page) " OF " counter(pages); color:#64748b; font-size:8px; }
                    }
                    @counter-style proof-steps {
                      system: fixed;
                      symbols: "01" "02" "03";
                      suffix: ". ";
                    }
                    @layer reset {
                      * { box-sizing:border-box; }
                      body { margin:0; }
                    }
                    @layer theme {
                      body {
                        page:renderer-gallery;
                        font:12px/1.45 Arial,sans-serif;
                        color:#172033;
                        background:#eaf0f8;
                      }
                      .sheet {
                        min-height:720px;
                        padding:22px;
                        border:1px solid #cbd7ea;
                        border-radius:18px;
                        background:#f8fbff;
                        box-shadow:0 4px 10px rgba(30,64,175,.1);
                      }
                    }
                    @layer components {
                      .hero {
                        display:grid;
                        grid-template-columns:minmax(0,1.8fr) minmax(230px,.8fr);
                        gap:18px;
                        align-items:center;
                        padding:20px 22px;
                        border-radius:15px;
                        color:white;
                        background:linear-gradient(120deg,oklch(38% .16 258),oklch(58% .2 255) 56%,oklch(70% .17 197));
                        box-shadow:0 4px 10px rgba(29,78,216,.18);
                      }
                      .eyebrow { margin:0 0 5px; font-size:9px; font-weight:800; letter-spacing:.18em; text-transform:uppercase; color:#bfdbfe; }
                      h1 { margin:0; font-size:28px; line-height:1.04; letter-spacing:-.025em; }
                      .hero p { max-width:620px; margin:8px 0 0; color:#e6f0ff; }
                      .hero-art {
                        display:grid;
                        grid-template-columns:76px 1fr;
                        gap:12px;
                        align-items:center;
                        min-height:94px;
                        padding:12px;
                        border:1px solid rgba(255,255,255,.28);
                        border-radius:13px;
                        background:rgba(255,255,255,.1);
                      }
                      .hero-art__ring {
                        width:76px;
                        height:76px;
                        border:8px solid rgba(255,255,255,.2);
                        border-radius:50%;
                        background:conic-gradient(from 28deg at center,oklch(78% .19 145) 0 37%,oklch(78% .17 70) 37% 66%,oklch(74% .17 250) 66% 100%);
                      }
                      .hero-art strong { display:block; font-size:19px; }
                      .hero-art small { display:block; margin-top:4px; color:#dbeafe; }
                      .metrics {
                        display:grid;
                        grid-template-columns:repeat(4,minmax(0,1fr));
                        gap:10px;
                        margin-top:12px;
                      }
                      .metric {
                        padding:10px 12px;
                        border:1px solid #d8e2f1;
                        border-radius:11px;
                        background:white;
                      }
                      .metric span { display:block; color:#64748b; font-size:8px; font-weight:800; letter-spacing:.08em; text-transform:uppercase; }
                      .metric strong { display:block; margin-top:3px; color:#1d4ed8; font-size:15px; }
                      .workspace {
                        display:grid;
                        grid-template-columns:minmax(0,1.35fr) minmax(340px,.75fr);
                        gap:12px;
                        margin-top:12px;
                      }
                      .panel {
                        padding:14px;
                        border:1px solid #d8e2f1;
                        border-radius:13px;
                        background:white;
                      }
                      .panel-head { display:flex; justify-content:space-between; gap:12px; align-items:start; margin-bottom:10px; }
                      .panel h2 { margin:0; font-size:15px; }
                      .panel-head p { margin:2px 0 0; color:#64748b; font-size:9px; }
                      .status { padding:3px 8px; border-radius:999px; color:#047857; background:#d1fae5; font-size:8px; font-weight:800; }
                      .pipeline {
                        display:grid;
                        grid-template-columns:repeat(3,minmax(0,1fr));
                        gap:8px;
                        align-items:stretch;
                      }
                      .step {
                        display:grid;
                        grid-template-columns:subgrid;
                        grid-column:span 1;
                        gap:5px;
                        padding:10px;
                        border:1px solid #d8e2f1;
                        border-radius:10px;
                        background:#f8fbff;
                      }
                      .step b { color:#2563eb; font-size:9px; letter-spacing:.08em; }
                      .step strong { font-size:11px; }
                      .step p { margin:0; color:#64748b; font-size:8px; }
                      .capability { container-type:inline-size; margin-top:10px; }
                      .capability-grid { display:grid; grid-template-columns:1fr; gap:8px; }
                      @container (min-width: 480px) { .capability-grid { grid-template-columns:1fr 1fr; } }
                      .capability-card { padding:10px; border-radius:10px; background:#eff6ff; }
                      .capability-card strong { display:block; margin-bottom:4px; color:#1e40af; }
                      .capability-card p { margin:0; color:#475569; font-size:8px; }
                      .formula {
                        display:flex;
                        gap:10px;
                        align-items:center;
                        margin-top:8px;
                        padding:8px 10px;
                        border-left:4px solid #7c3aed;
                        border-radius:0 9px 9px 0;
                        background:#f5f3ff;
                      }
                      .formula span { color:#6d28d9; font-size:8px; font-weight:800; text-transform:uppercase; }
                      math { font-size:15px; }
                      .proof-list { margin:0; padding-left:28px; list-style:proof-steps; }
                      .proof-list li { margin:0 0 7px; padding-left:4px; color:#475569; font-size:9px; }
                      .proof-list li::marker { color:#2563eb; font-weight:900; }
                      .proof-list strong { color:#172033; }
                      .hyphen-proof {
                        margin-top:8px;
                        max-width:150px;
                        padding:8px;
                        border-radius:9px;
                        background:#fff7ed;
                        color:#9a3412;
                        font-size:8px;
                        hyphens:auto;
                        hyphenate-limit-chars:6 3 3;
                      }
                      .controls { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:8px; margin-top:9px; }
                      .controls label { color:#475569; font-size:8px; font-weight:700; }
                      .controls > input[type=text],.controls > select {
                        display:block;
                        width:100%;
                        height:27px;
                        padding:5px 7px;
                        border:1px solid #94a3b8;
                        border-radius:7px;
                        background:#f8fafc;
                        color:#172033;
                      }
                      .choice-row { grid-row:3; grid-column:1 / span 2; display:grid; grid-template-columns:1.35fr 1fr 1fr; gap:7px; align-items:center; min-height:18px; }
                      .choice-row label { display:flex; gap:4px; align-items:center; min-height:18px; white-space:nowrap; }
                      .choice-row input { width:12px; height:12px; flex:0 0 12px; }
                      .footnote { margin:10px 2px 0; color:#64748b; font-size:8px; }
                    }
                  </style>
                </head>
                <body>
                  <main class="sheet">
                    <header class="hero">
                      <div>
                        <p class="eyebrow">OfficeIMO managed document renderer</p>
                        <h1>One authored document.<br>Four dependable outputs.</h1>
                        <p>AngleSharp parsing, bounded document layout, searchable vector output, and native form fields—without a browser process.</p>
                      </div>
                      <div class="hero-art">
                        <div class="hero-art__ring" aria-label="Three-color conic gradient"></div>
                        <div><strong>PDF · PNG · SVG</strong><small>Same parsed source and policy</small></div>
                      </div>
                    </header>

                    <section class="metrics" aria-label="Renderer capabilities">
                      <div class="metric"><span>Cascade</span><strong>Layers + nesting</strong></div>
                      <div class="metric"><span>Layout</span><strong>Grid + subgrid</strong></div>
                      <div class="metric"><span>Paint</span><strong>CSS Color 4</strong></div>
                      <div class="metric"><span>Output</span><strong>Real AcroForms</strong></div>
                    </section>

                    <section class="workspace">
                      <article class="panel">
                        <div class="panel-head"><div><h2>Render pipeline</h2><p>One managed scene keeps output adapters thin.</p></div><span class="status">Validated</span></div>
                        <div class="pipeline">
                          <div class="step"><b>01 / PARSE</b><strong>HTML + CSS</strong><p>AngleSharp, layers, nesting, media and container queries.</p></div>
                          <div class="step"><b>02 / LAYOUT</b><strong>Document scene</strong><p>Paged grid, subgrid, typography, counters, SVG and MathML.</p></div>
                          <div class="step"><b>03 / EXPORT</b><strong>Portable output</strong><p>Searchable PDF, vector SVG, PNG and native form fields.</p></div>
                        </div>
                        <div class="capability">
                          <div class="capability-grid">
                            <div class="capability-card"><strong>Container-aware composition</strong><p>This pair becomes two columns only when its own container has room, not merely when the page is wide.</p></div>
                            <div class="capability-card"><strong>Bounded by design</strong><p>Resource limits, deterministic fallbacks, structured diagnostics, and no script execution are part of the rendering contract.</p></div>
                          </div>
                        </div>
                        <div class="formula"><span>MathML → OfficeMath</span><math><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math></div>
                      </article>

                      <aside class="panel">
                        <div class="panel-head"><div><h2>Proof inside the PDF</h2><p>These controls remain fillable.</p></div></div>
                        <ol class="proof-list">
                          <li><strong>Stylesheet paint:</strong> layered conic and linear gradients stay vector.</li>
                          <li><strong>Unicode text:</strong> line breaking and fallback remain searchable.</li>
                          <li><strong>Page semantics:</strong> named geometry and margin counters are resolved.</li>
                        </ol>
                        <div id="hyphen-proof" class="hyphen-proof" lang="de">Donaudampfschifffahrtsgesellschaftskapitän demonstrates language-aware hyphenation in a deliberately narrow measure.</div>
                        <form class="controls">
                          <label for="gallery-reviewer">Reviewer</label>
                          <label for="gallery-decision">Decision</label>
                          <input id="gallery-reviewer" type="text" name="reviewer" value="Ada Lovelace">
                          <select id="gallery-decision" name="decision"><option selected>Approved</option><option>Needs changes</option></select>
                          <div class="choice-row">
                            <label><input type="checkbox" name="verified" checked><span>Evidence verified</span></label>
                            <label><input type="radio" name="lane" value="managed" checked><span>Managed</span></label>
                            <label><input type="radio" name="lane" value="browser"><span>Browser</span></label>
                          </div>
                        </form>
                      </aside>
                    </section>
                    <p class="footnote">Generated from a runnable OfficeIMO example. The downloadable PDF, PNG, SVG, and HTML source are produced together and hashed in the website evidence manifest.</p>
                  </main>
                </body>
                </html>
                """;

            var options = new HtmlPdfSaveOptions {
                PageSize = OfficePageSizes.A4.Landscape(),
                Margins = HtmlRenderMargins.All(18D),
                BackgroundColor = OfficeColor.White,
                Scale = 1D,
                ConicGradientQualitySegments = 72
            };
            options.UseTextHyphenationLexicon(new OfficeTextHyphenationLexicon(new[] {
                "Do-nau-dampf-schiff-fahrts-ge-sell-schafts-ka-pi-tän"
            }, minimumPrefixLength: 2, minimumSuffixLength: 2));

            string htmlPath = Path.Combine(folderPath, "HtmlManagedRendererGallery.html");
            string pdfPath = Path.Combine(folderPath, "HtmlManagedRendererGallery.pdf");
            string pngPath = Path.Combine(folderPath, "HtmlManagedRendererGallery.png");
            string svgPath = Path.Combine(folderPath, "HtmlManagedRendererGallery.svg");

            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
            HtmlDiagnostic[] fidelityDiagnostics = rendered.Diagnostics
                .Where(diagnostic => diagnostic.Severity != HtmlDiagnosticSeverity.Info)
                .ToArray();
            if (fidelityDiagnostics.Length > 0) {
                throw new InvalidOperationException("Managed renderer gallery emitted fidelity diagnostics: " + string.Join("; ", fidelityDiagnostics.Select(diagnostic =>
                    diagnostic.Code + " (" + diagnostic.Source + ": " + diagnostic.Detail + ")")));
            }
            bool hasHyphenationProof = rendered.Pages
                .SelectMany(page => EnumerateManagedRendererGalleryVisuals(page.Scene))
                .OfType<HtmlRenderText>()
                .Any(text => text.Text.StartsWith("Donaudampf", StringComparison.Ordinal)
                    && text.Text.EndsWith("-", StringComparison.Ordinal));
            if (!hasHyphenationProof) {
                throw new InvalidOperationException("Managed renderer gallery expected the German proof word to use the configured automatic hyphenation lexicon.");
            }

            File.WriteAllText(htmlPath, html);
            document.SaveAsPdf(pdfPath, options);
            document.ToImage(options)
                .AsPng()
                .OnFileConflict(OfficeImageExportFileConflictPolicy.Replace)
                .Save(pngPath);
            document.ToImage(options)
                .AsSvg()
                .OnFileConflict(OfficeImageExportFileConflictPolicy.Replace)
                .Save(svgPath);

            byte[] pdf = File.ReadAllBytes(pdfPath);
            int formFieldCount = global::OfficeIMO.Pdf.PdfDocument.Load(pdf).Inspect().FormFieldCount;
            if (formFieldCount < 4) {
                throw new InvalidOperationException($"Managed renderer gallery expected at least four PDF fields but found {formFieldCount}.");
            }

            Console.WriteLine($"    HTML: {htmlPath}");
            Console.WriteLine($"    PDF: {pdfPath} ({formFieldCount} interactive fields)");
            Console.WriteLine($"    PNG: {pngPath}");
            Console.WriteLine($"    SVG: {svgPath}");

            if (openPdf) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }

        private static IEnumerable<HtmlRenderVisual> EnumerateManagedRendererGalleryVisuals(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (HtmlRenderVisual visual in visuals) {
                yield return visual;
                IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clip
                    ? clip.Visuals
                    : visual is HtmlRenderPathClipGroup pathClip
                        ? pathClip.Visuals
                        : visual is HtmlRenderEffectGroup effect
                            ? effect.Visuals
                            : visual is HtmlRenderSemanticGroup semantic
                                ? semantic.Visuals
                                : visual is HtmlRenderLogicalTextGroup logical
                                    ? logical.Visuals
                                    : visual is HtmlRenderFormField form ? form.Visuals : null;
                if (children == null) continue;
                foreach (HtmlRenderVisual child in EnumerateManagedRendererGalleryVisuals(children)) yield return child;
            }
        }
    }
}
