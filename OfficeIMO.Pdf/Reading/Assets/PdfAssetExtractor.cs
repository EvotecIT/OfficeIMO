using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfAssetExtractor {
    internal static IReadOnlyList<string> Extract(PdfReadDocument document, Func<string, Stream> create,
        PdfAssetExtractionOptions options, CancellationToken token) {
        document.DemandContentExtraction("PDF asset");
        if (document.Pages.Count > options.MaximumPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaximumPages, document.Pages.Count);
        }

        var output = new PdfAssetOutput(create, options.MaximumOutputBytes, token);
        var report = new PdfAssetReport();
        var inventory = new PdfAssetResourceInventory(document, output, report, token);
        var resourceSets = new List<IReadOnlyList<PdfDictionary>>(document.Pages.Count);
        inventory.InspectOutputIntents();
        long placementWork = 0;
        long maxPlacementWork = document.ReadOptions.Limits.MaxContentOperands;
        for (int index = 0; index < document.Pages.Count; index++) {
            token.ThrowIfCancellationRequested();
            PdfReadPage page = document.Pages[index];
            int pageNumber = index + 1;
            int remainingPlacements = PdfAssetReport.Table.MaximumRows - report.Placements.Rows.Count;
            var placements = page.GetImagePlacements(pageNumber, remainingPlacements, units => {
                placementWork += units;
                if (placementWork > maxPlacementWork) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maxPlacementWork, placementWork);
                }
            }, token.ThrowIfCancellationRequested, token);
            AddBoxes(pageNumber, page.GetGeometry(), report);
            IReadOnlyList<PdfDictionary> pageResources = inventory.InspectPage(page, pageNumber, placements);
            resourceSets.Add(pageResources);
            ExtractImages(page, pageNumber, placements, pageResources, options.CompositeSoftMasks, output, report, token);
        }

        ExtractFonts(document, resourceSets, output, report, token);
        int svgPage = 0;
        PdfImageExportEngine.ExportEach(document, OfficeImageExportFormat.Svg,
            new PdfImageExportOptions { MaximumOutputCount = options.MaximumPages, MaximumTotalEncodedBytes = options.MaximumOutputBytes }, null, image => {
                int pageNumber = ++svgPage;
                string file = "pages/" + pageNumber.ToString("D4", CultureInfo.InvariantCulture) + ".svg";
                output.Write(file, image.Bytes);
                report.Pages.Add(pageNumber, file);
                foreach (OfficeImageExportDiagnostic diagnostic in image.Diagnostics) {
                    report.Diagnostics.Add(pageNumber, "SVG", diagnostic.Code, diagnostic.Message);
                }
            }, cancellationToken: token);

        output.WriteReport("report.json", writer => report.WriteJson(writer, token));
        output.WriteReport("report.html", writer => report.WriteHtml(writer, token));
        return output.Paths;
    }

    internal static void ExtractStandaloneFonts(PdfReadDocument document, PdfAssetOutput output, CancellationToken token) {
        document.DemandContentExtraction("font resource");
        var inventory = new PdfAssetResourceInventory(document, null, null, token);
        IReadOnlyList<IReadOnlyList<PdfDictionary>> resourceSets = document.Pages
            .Select((page, index) => inventory.InspectPage(page, index + 1)).ToArray();
        ExtractFonts(document, resourceSets, output, null, token);
    }

    private static void ExtractFonts(PdfReadDocument document, IReadOnlyList<IReadOnlyList<PdfDictionary>> resourceSets,
        PdfAssetOutput output, PdfAssetReport? report, CancellationToken token) {
        var options = new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true,
            InspectEmbeddedProgramMetadata = false,
            MaxEmbeddedProgramBytes = Math.Min(32 * 1024 * 1024, document.ReadOptions.Limits.MaxDecodedStreamBytes)
        };
        PdfFontInventory inventory = PdfFontInspector.Inspect(document, options, assetResources: resourceSets, cancellationToken: token);
        var exportedPrograms = new Dictionary<string, ExportedFont>(StringComparer.Ordinal);
        foreach (PdfFontInfo font in inventory.Fonts) {
            token.ThrowIfCancellationRequested();
            ExportedFont? exportedFont = null;
            string? sourceHash = null;
            if (font.DecodedProgram is { Length: > 0 } bytes) {
                // Hash the source program before name repair so aliases of the same subset share one file.
                sourceHash = PdfArtifactInfo.ComputeSha256Hex(bytes, token);
                if (!exportedPrograms.TryGetValue(sourceHash, out exportedFont)) {
                    string extension = PdfExtractedFontProgram.Extension(font.EmbeddedProgramSubtype, bytes);
                    bool repaired = false;
                    if (extension is "ttf" or "otf") {
                        bytes = PdfExtractedFontProgram.RepairName(bytes, font, out repaired, out bool readable);
                        if (!readable) report?.Diagnostics.Add(null, font.BaseFontName, "UnreadableFontDirectory",
                            "The embedded font table directory could not be repaired.");
                    }
                    string file = output.WriteDistinct("fonts", extension, bytes, sourceHash);
                    string fileHash = repaired ? PdfArtifactInfo.ComputeSha256Hex(bytes, token) : sourceHash;
                    exportedFont = new ExportedFont(file, fileHash, repaired);
                    exportedPrograms.Add(sourceHash, exportedFont);
                }
            }
            report?.Fonts.Add(font.BaseFontName, font.FamilyName, font.Subtype, font.IsEmbedded, font.IsSubset, font.SubsetTag,
                font.EmbeddedProgramSubtype, font.References.Select(static reference => reference.PageNumber).Distinct().ToArray(),
                exportedFont?.Path, exportedFont?.Repaired ?? false, sourceHash, exportedFont?.FileHash);
            if (report is not null) {
                foreach (PdfFontInspectionDiagnostic diagnostic in font.Diagnostics) AddFontDiagnostic(report, diagnostic, font.BaseFontName);
            }
        }
        if (report is not null) {
            foreach (PdfFontInspectionDiagnostic diagnostic in inventory.Diagnostics) AddFontDiagnostic(report, diagnostic, "Fonts");
        }
    }

    private sealed record ExportedFont(string Path, string FileHash, bool Repaired);

    private static void AddFontDiagnostic(PdfAssetReport report, PdfFontInspectionDiagnostic diagnostic, string source) =>
        report.Diagnostics.Add(diagnostic.PageNumber, source, diagnostic.Code.ToString(), diagnostic.Message);

    private static void ExtractImages(PdfReadPage page, int pageNumber, IReadOnlyList<PdfImagePlacement> placements,
        IReadOnlyList<PdfDictionary> resourceSets, bool compositeSoftMasks, PdfAssetOutput output, PdfAssetReport report,
        CancellationToken token) {
        page.VisitAssetImages(pageNumber, placements, resourceSets, compositeSoftMasks, image => {
            byte[] bytes = image.EncodedBytes;
            string hash = PdfArtifactInfo.ComputeSha256Hex(bytes, token);
            string path = output.WriteDistinct("images", image.FileExtension ?? "bin", bytes, hash);
            bool rgbaPng = image.FileExtension == "png" && ResourceResolver.IsRgbaPng(bytes);
            report.Images.Add(pageNumber, image.ResourceName, image.ObjectNumber, path, image.Width, image.Height,
                image.ColorSpace, image.Filter, image.RenderingIntent.ToString(), image.TransparencyMaskKind,
                image.TransparencyMaskResolved, rgbaPng, image.IsImageFile, hash);
            if (!image.IsImageFile) {
                report.Diagnostics.Add(pageNumber, image.ResourceName, "ImageFileUnavailable",
                    "The image stream was saved as raw bytes because a standalone image could not be produced.");
            }
            if (compositeSoftMasks && image.TransparencyMaskKind == "soft-mask" && (!rgbaPng || image.HasUnresolvedTransparencyMask)) {
                report.Diagnostics.Add(pageNumber, image.ResourceName, "ImageSoftMaskNotExported",
                    "An RGBA PNG with the PDF soft mask could not be produced.");
            }
            return path;
        }, (placement, path) => report.Placements.Add(pageNumber, placement.ResourceName, placement.ObjectNumber, path,
            placement.EffectiveDpiX, placement.EffectiveDpiY), token);
    }

    private static void AddBoxes(int pageNumber, PdfPageGeometry geometry, PdfAssetReport report) {
        double mmPerUnit = (geometry.UserUnit ?? 1D) * (25.4D / 72D);
        Add(geometry.MediaBox);
        Add(geometry.CropBox);
        Add(geometry.BleedBox);
        Add(geometry.TrimBox);
        Add(geometry.ArtBox);
        Add(geometry.EffectiveBox, "EffectiveBox");

        void Add(PdfPageBox? box, string? name = null) {
            if (box is null) return;
            report.Boxes.Add(pageNumber, name ?? box.Name, box.Left * mmPerUnit, box.Bottom * mmPerUnit,
                box.Right * mmPerUnit, box.Top * mmPerUnit, box.Width * mmPerUnit, box.Height * mmPerUnit);
        }
    }
}
