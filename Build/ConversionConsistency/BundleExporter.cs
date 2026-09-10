using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.ConversionConsistency;

internal static class BundleExporter {
    internal static async Task ExportAsync(string repository, string suitePath, string output, string? caseId, CancellationToken cancellationToken) {
        ConsistencySuite suite = GateJson.Read<ConsistencySuite>(suitePath);
        if (suite.SchemaVersion != 1 || suite.Dpi < 36 || suite.Dpi > 600 || suite.Cases.Count == 0)
            throw new InvalidDataException("Unsupported or empty consistency suite.");
        if (Directory.Exists(output) && Directory.EnumerateFileSystemEntries(output).Any())
            throw new IOException("The output directory must be empty to prevent stale evidence: " + output);
        var selected = suite.Cases.Where(item => caseId == null || item.Id == caseId).ToList();
        if (selected.Count == 0 || selected.Select(item => item.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != selected.Count)
            throw new InvalidDataException("No matching cases or duplicate case identifiers.");
        byte[] fontBytes = File.ReadAllBytes(ArtifactPaths.Resolve(repository, suite.FontPath));
        var fonts = new OfficeFontFaceCollection().Add(suite.FontFamily, fontBytes).AddFallbackFamily(suite.FontFamily);
        var profile = new OfficeRenderingProfile("conversion-consistency", fonts, OfficeManagedTextShapingProvider.Instance);
        var provenance = await ArtifactPaths.ProvenanceAsync(repository);
        Directory.CreateDirectory(output);
        var cases = new List<CaseBundle>();
        foreach (ConsistencyCase contract in selected) {
            cancellationToken.ThrowIfCancellationRequested();
            if (contract.Id.Length == 0 || contract.Id.Any(character => !char.IsAsciiLetterOrDigit(character) && character != '-'))
                throw new InvalidDataException("Case identifiers must contain ASCII letters, digits, or hyphens.");
            if (contract.Pages.Count == 0 || contract.Pages.Any(page => page.Width <= 0 || page.Height <= 0 || page.Text.Count == 0))
                throw new InvalidDataException("Every case needs labelled page expectations.");
            string source = ArtifactPaths.Resolve(repository, contract.Source);
            string caseDirectory = ArtifactPaths.Resolve(output, contract.Id);
            Directory.CreateDirectory(caseDirectory);
            NativeExport exported = NativeAdapters.Export(contract, source, suite, profile, cancellationToken);
            string pdfPath = contract.Id + "/document.pdf";
            File.WriteAllBytes(ArtifactPaths.Resolve(output, pdfPath), exported.Pdf);
            var images = new List<ImageArtifact>();
            var diagnostics = new List<string>(exported.Diagnostics);
            var diagnosticDetails = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeImageExportFormat format in Enum.GetValues<OfficeImageExportFormat>()) {
                int page = 0;
                foreach (OfficeImageExportResult image in exported.Images(format)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string relative = contract.Id + "/page-" + (++page).ToString("D3") + image.FileExtension;
                    byte[] bytes = image.Bytes;
                    File.WriteAllBytes(ArtifactPaths.Resolve(output, relative), bytes);
                    images.Add(new ImageArtifact(page, format.ToString(), relative, ArtifactPaths.Hash(bytes), image.Width, image.Height));
                    diagnostics.AddRange(image.Diagnostics.Select(item => item.Code));
                    foreach (var diagnostic in image.Diagnostics)
                        diagnosticDetails.Add(diagnostic.Code + ": " + diagnostic.Message + " [" + diagnostic.Source + "]");
                }
            }
            cases.Add(new CaseBundle(contract, ArtifactPaths.HashFile(source), pdfPath, ArtifactPaths.Hash(exported.Pdf), exported.PdfRoute, images, diagnostics.Distinct().ToList()) {
                DiagnosticDetails = diagnosticDetails.OrderBy(value => value, StringComparer.Ordinal).ToList()
            });
        }
        GateJson.Write(Path.Combine(output, "bundle.json"), new EvidenceBundle(1, provenance.Commit, provenance.DiffHash,
            ArtifactPaths.Hash(fontBytes), suite.FontFamily, suite.Dpi, nameof(OfficeManagedTextShapingProvider), "#ffffff", cases, provenance.Untracked));
    }
}

internal sealed record NativeExport(byte[] Pdf, Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> Images, List<string> Diagnostics, string PdfRoute = "native");

internal static partial class NativeAdapters {
    internal static NativeExport Export(ConsistencyCase contract, string source, ConsistencySuite suite, OfficeRenderingProfile profile, CancellationToken cancellationToken) {
        if (contract.Format != "html") return ExportOffice(contract, source, suite, profile, cancellationToken);
        var document = HtmlConversionDocument.Parse(File.ReadAllText(source));
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, TargetDpi = suite.Dpi, BackgroundColor = OfficeColor.White,
            DefaultFontFamily = suite.FontFamily, Margins = HtmlRenderMargins.All(0),
            PageSize = new OfficePageSize(contract.Pages[0].Width / (double)suite.Dpi, contract.Pages[0].Height / (double)suite.Dpi)
        };
        options.UseRenderingProfile(profile);
        var pdf = document.ToPdfDocumentResult(new HtmlToPdfOptions(options), cancellationToken);
        return new NativeExport(pdf.ToBytes(cancellationToken), format => document.ExportImages(format, options),
            pdf.Report.Warnings.Select(item => item.Code).ToList());
    }
}
