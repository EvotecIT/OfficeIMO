using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        /// <summary>Renders the small, self-contained HTML examples used by the website gallery.</summary>
        public static void Example_HtmlFeatureShowcase(string folderPath) {
            string sourceRoot = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "Showcase");
            string outputRoot = Path.Combine(folderPath, "HtmlFeatures");
            Directory.CreateDirectory(outputRoot);
            string[] names = Directory.GetFiles(sourceRoot, "*.html")
                .Select(path => Path.GetFileNameWithoutExtension(path)).OrderBy(name => name, StringComparer.Ordinal).ToArray();

            foreach (string name in names) {
                string html = File.ReadAllText(Path.Combine(sourceRoot, name + ".html"));
                var options = new HtmlToPdfOptions {
                    PageSize = OfficePageSizes.A4.Landscape(),
                    Margins = HtmlRenderMargins.All(24D),
                    BackgroundColor = OfficeColor.White,
                    ConicGradientQualitySegments = 72
                };
                HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
                HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
                HtmlDiagnostic[] warnings = rendered.Diagnostics
                    .Where(diagnostic => diagnostic.Severity != HtmlDiagnosticSeverity.Info).ToArray();
                if (warnings.Length > 0) {
                    throw new InvalidOperationException(name + ": " + string.Join("; ", warnings.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Detail)));
                }
                int expectedPages = name == "page-breaks" ? 2 : 1;
                if (rendered.Pages.Count != expectedPages) {
                    throw new InvalidOperationException($"{name}: expected {expectedPages} page(s), found {rendered.Pages.Count}.");
                }

                string stem = Path.Combine(outputRoot, name);
                File.WriteAllText(stem + ".html", html);
                document.SaveAsPdf(stem + ".pdf", options);
                document.ToImage(options).AsPng().OnFileConflict(OfficeImageExportFileConflictPolicy.Replace).Save(stem + ".png");
                document.ToImage(options).AsSvg().OnFileConflict(OfficeImageExportFileConflictPolicy.Replace).Save(stem + ".svg");
                var inspection = global::OfficeIMO.Pdf.PdfDocument.Load(File.ReadAllBytes(stem + ".pdf")).Inspect();
                if (name == "fillable-forms" && inspection.FormFieldCount < 4) {
                    throw new InvalidOperationException($"{name}: expected at least four interactive PDF fields, found {inspection.FormFieldCount}.");
                }
                Console.WriteLine($"{name}: {rendered.Pages.Count} page(s), {inspection.FormFieldCount} fields; HTML, PDF, PNG, SVG written to {outputRoot}");
            }
        }
    }
}
