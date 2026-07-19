using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Generates one inspectable proof bundle: editable PPTX, PNG/SVG previews, PDF, HTML, and structured reports.
    /// </summary>
    public static class EndToEndPowerPointProof {
        public static void Example_EndToEndPowerPointProof(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - end-to-end proof bundle");
            string outputFolder = Path.Combine(folderPath, "PowerPoint End To End Proof");
            Directory.CreateDirectory(outputFolder);
            string presentationPath = Path.Combine(outputFolder, "Executive Delivery Review.pptx");
            string pdfPath = Path.Combine(outputFolder, "Executive Delivery Review.pdf");
            string handoutPath = Path.Combine(outputFolder, "Executive Delivery Review.handout.pdf");
            string htmlPath = Path.Combine(outputFolder, "Executive Delivery Review.html");
            string reportPath = Path.Combine(outputFolder, "Executive Delivery Review.preflight.json");
            string accessibilityPath = Path.Combine(outputFolder, "Executive Delivery Review.accessibility.json");
            string visualProofPath = Path.Combine(outputFolder, "Executive Delivery Review.visual-proof.json");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#0B7FAB", "e2e-proof", "executive delivery review")
                .WithIdentity("Delivery Review", eyebrow: "OFFICEIMO", footerLeft: "DELIVERY",
                    footerRight: "Generated proof");
            string brandImagePath = Path.Combine(AppContext.BaseDirectory, "Assets", "OfficeIMO.png");
            var brandImage = new PowerPointImageAsset(brandImagePath,
                "OfficeIMO brand mark used as a semantic visual asset") {
                Caption = "Semantic images carry visible context and accessibility metadata",
                Provenance = "OfficeIMO repository asset",
                Placement = PowerPointImagePlacement.Fit
            }.Annotate(new PowerPointImageAnnotation(0.5, 0.5, "Real asset",
                "The picture remains native and editable in PowerPoint."));
            var chartData = new OfficeChartData(
                new[] { "Q1", "Q2", "Q3", "Q4" },
                new[] {
                    new OfficeChartSeries("Adoption", new[] { 28D, 43D, 61D, 72D }, null, null, null,
                        showMarkers: false, renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries("Conversion", new[] { 6.2D, 7.8D, 9.6D, 11.4D }, null, null, null,
                        showMarkers: true, markerSize: 8, renderKind: OfficeChartKind.Line,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                });
            var chartStory = new PowerPointChartStoryContent(OfficeChartKind.ColumnClustered, chartData,
                new[] { "Adoption improved every quarter.", "Conversion increased with adoption." }) {
                Caption = "Quarterly adoption and conversion",
                Provenance = "Illustrative customer-success dataset",
                AlternativeText = "Adoption columns with conversion line on a secondary axis",
                DataSummary = "Adoption rose from 28 to 72 while conversion rose from 6.2 to 11.4."
            };
            var appendixData = new PowerPointTableData(new[] { "Workstream", "Status", "Complete" },
                Enumerable.Range(1, 17).Select(index => (IEnumerable<string>)new[] {
                    "Workstream " + index,
                    index % 3 == 0 ? "At risk" : "On track",
                    (50 + index * 2) + "%"
                })) {
                Caption = "Delivery evidence",
                Provenance = "Illustrative program dataset",
                Notes = new[] { "Rows continue deterministically.", "Column headers repeat on every page." }
            };
            var architecture = new PowerPointArchitectureContent(new[] {
                new PowerPointArchitectureNode("drawing", "Drawing core", "Shared semantics", "Core"),
                new PowerPointArchitectureNode("ppt", "PowerPoint", "Native authoring", "Surfaces"),
                new PowerPointArchitectureNode("markup", "Markup", "Thin adapter", "Surfaces"),
                new PowerPointArchitectureNode("output", "Proof bundle", "PPTX, PNG, SVG, PDF", "Outputs")
            }, new[] {
                new PowerPointArchitectureEdge("drawing", "ppt", "drives"),
                new PowerPointArchitectureEdge("drawing", "markup", "shares"),
                new PowerPointArchitectureEdge("ppt", "output", "publishes")
            });
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Executive delivery review", "Editable, measured, and ready for inspection")
                .AddExecutiveSummary("Executive summary", "A decision-ready opening",
                    new PowerPointExecutiveSummaryContent(
                        new[] { new PowerPointMetric("14", "proof slides"), new PowerPointMetric("0", "hidden rows") },
                        new[] {
                            new PowerPointCardContent("Decision", new[] { "Use semantic story families" }),
                            new PowerPointCardContent("Evidence", new[] { "Preflight every generated deck" })
                        }, "OfficeIMO now connects narrative intent to editable PowerPoint primitives."))
                .AddChartStory("Adoption story", "Native chart plus narrative and source context", chartStory)
                .AddComparison("Implementation choice", "Two options, one explicit recommendation",
                    new[] {
                        new PowerPointComparisonItem("Shared semantic core", "One owner for reusable behavior",
                            new[] { "Consistent", "Testable" }, new[] { "Requires deliberate contracts" }),
                        new PowerPointComparisonItem("Local slide helpers", "Per-project composition logic",
                            new[] { "Fast locally" }, new[] { "Drifts", "Duplicates behavior" })
                    })
                .AddArchitecture("End-to-end ownership", "Native shapes keep the system map editable", architecture)
                .AddScreenshotStory("Semantic image proof", "Crop, focal point, caption, provenance, and alt text",
                    brandImage, new[] { "Real image relationship", "Visible provenance", "Accessible description" })
                .AddSection("Delivery and evidence", "A contrast reset before implementation detail",
                    configure: options => options.SectionVariant = PowerPointSectionLayoutVariant.GeometricCover)
                .AddProcess("Delivery path", "Long content continues deterministically",
                    Enumerable.Range(1, 8).Select(index => new PowerPointProcessStep(
                        "Phase " + index, "Evidence and owner for delivery phase " + index + ".")))
                .AddCardGrid("Decision signals", "A semantic story rendered as native shapes",
                    Enumerable.Range(1, 8).Select(index => new PowerPointCardContent(
                        "Signal " + index, new[] { "Owner assigned", "Evidence available" })))
                .AddAppendixTable("Delivery detail", "Editable rows continue without truncation", appendixData)
                .AddClosing("Next decision", new PowerPointClosingContent(
                    "Beautiful automation is trustworthy automation.",
                    "Inspect the proof bundle and approve the next release."),
                    configure: options => options.Variant = PowerPointClosingLayoutVariant.Statement);

            var preflightOptions = new PowerPointDeckPreflightOptions {
                MinimumReadableFontSizePoints = 8,
                DetectShapeCollisions = false
            };
            PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(brief);
            composition.SelectBestAlternative = false;
            composition.AlternativeIndex = 0;
            composition.Preflight = preflightOptions;
            PowerPointCompositionResult result = presentation.Compose(plan, composition);
            PowerPointDeckRhythmReport rhythm = result.Plan.InspectRhythm(result.Design);
            PowerPointDeckPreflightReport report = result.Preflight;
            report.SaveJson(reportPath);
            PowerPointAccessibilityReport accessibility = presentation.InspectAccessibility();
            accessibility.EnsureCompliant().SaveJson(accessibilityPath);
            presentation.Save();

            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                string stem = "Slide " + (slideIndex + 1).ToString("00");
                presentation.Slides[slideIndex].SaveAsPng(Path.Combine(outputFolder, stem + ".png"));
                presentation.Slides[slideIndex].SaveAsSvg(Path.Combine(outputFolder, stem + ".svg"));
            }
            var pdfOptions = new PowerPointPdfSaveOptions().UseProfile(PdfExportProfile.Faithful);
            PdfDocumentConversionResult pdfResult = presentation.ToPdfDocumentResult(pdfOptions);
            pdfResult.Save(pdfPath);
            presentation.SaveAsPdf(handoutPath, new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.Handouts,
                HandoutSlidesPerPage = 3,
                IncludeSpeakerNotes = true
            });
            var htmlOptions = new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointVisualReview
            };
            PowerPointToHtmlResult htmlResult = presentation.ToHtmlResult(htmlOptions);
            File.WriteAllText(htmlPath, htmlResult.Value, Encoding.UTF8);

            PowerPointVisualProofReport visualProof = presentation.InspectVisuals()
                .RecordArtifact(Path.GetFileName(presentationPath),
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    File.ReadAllBytes(presentationPath))
                .RecordArtifact(Path.GetFileName(pdfPath), "application/pdf", File.ReadAllBytes(pdfPath),
                    pdfResult.Warnings.Count)
                .RecordArtifact(Path.GetFileName(htmlPath), "text/html", File.ReadAllBytes(htmlPath),
                    htmlResult.ImageDiagnostics.Count);
            visualProof.SaveJson(visualProofPath);

            List<ValidationErrorInfo> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                throw new InvalidOperationException("End-to-end proof deck has " + errors.Count +
                    " Open XML validation error(s): " +
                    string.Join("; ", errors.Take(3).Select(error => error.Description)));
            }

            Console.WriteLine("    Deck: " + presentationPath);
            Console.WriteLine("    Slides: " + presentation.Slides.Count);
            Console.WriteLine("    Preflight: " + report.ErrorCount + " errors, " + report.WarningCount + " warnings");
            Console.WriteLine("    Accessibility: " + accessibility.ErrorCount + " errors, " +
                              accessibility.WarningCount + " warnings");
            Console.WriteLine("    Rhythm: " + rhythm.Score + "/100, " + rhythm.Findings.Count + " finding(s)");
            Console.WriteLine("    Proof: PPTX, PNG, SVG, PDF, handout PDF, HTML, JSON, and Open XML validation");
            if (openPowerPoint) ExampleFileLauncher.Open(presentationPath);
        }
    }
}
