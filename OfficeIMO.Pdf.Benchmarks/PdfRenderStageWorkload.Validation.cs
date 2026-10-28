using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

public sealed partial class PdfRenderStageWorkload {
    /// <summary>
    /// Qualifies the measured result, completes downstream stages outside measurement, and retains compact scene and PNG evidence.
    /// Exact agreement with the public control proves boundary consistency for these bytes, not independent visual fidelity.
    /// </summary>
    public void Validate(string directory) {
        if (!_executed || _preparedOperation is null) throw new InvalidOperationException("No prepared rendering stage was executed.");
        if (SourceSha256 != Hash(_source)) throw new InvalidDataException("Immutable source bytes changed.");
        ResetValidationMetrics();

        PdfReadDocument reader = _reader ?? PdfReadDocument.Open(_source);
        ValidateSource(reader);
        OfficeDrawing[] drawings = _drawings ?? Project(reader);
        var sceneEvidence = new List<object>(PreviewPages);
        for (int index = 0; index < drawings.Length; index++) {
            OfficeDrawing drawing = drawings[index];
            OfficeDrawingText[] text = drawing.Elements.OfType<OfficeDrawingText>().ToArray();
            int images = drawing.Elements.OfType<OfficeDrawingImage>().Count();
            int shapes = drawing.Elements.OfType<OfficeDrawingShape>().Count();
            string logicalText = string.Join(' ', text.Select(element => element.Text));
            CheckText(logicalText, index + 1);
            if (drawing.Width != PdfRuntimeCorpus.PageWidth || drawing.Height != PdfRuntimeCorpus.PageHeight ||
                images != 1 || shapes == 0 || drawing.Fonts.Faces.Count == 0) {
                throw new InvalidDataException("A projected page lost its geometry, embedded font, image, or business shapes.");
            }
            ProjectionElements += drawing.Elements.Count;
            ProjectionTextElements += text.Length;
            ProjectionImageElements += images;
            ProjectionShapeElements += shapes;
            sceneEvidence.Add(new {
                PageNumber = index + 1, drawing.Width, drawing.Height, Elements = drawing.Elements.Count,
                TextElements = text.Length, ImageElements = images, ShapeElements = shapes,
                EmbeddedFontFaces = drawing.Fonts.Faces.Count,
                LogicalTextSha256 = Hash(Encoding.UTF8.GetBytes(Normalize(logicalText)))
            });
        }

        OfficeRasterImage[] rasters = _rasters ?? Rasterize(drawings);
        byte[][] pngs = _pngs ?? Encode(rasters);
        IReadOnlyList<PdfPageRenderResult> control = _control ?? RenderControl();
        // A measured complete control is compared against a second fresh public call as well as the prepared boundaries.
        IReadOnlyList<PdfPageRenderResult> reference = _control is null ? control : RenderControl();
        if (drawings.Length != PreviewPages || rasters.Length != PreviewPages || pngs.Length != PreviewPages ||
            control.Count != PreviewPages || reference.Count != PreviewPages) {
            throw new InvalidDataException("The four-page preview selection changed.");
        }

        var pageEvidence = new List<object>(PreviewPages);
        Directory.CreateDirectory(directory);
        for (int index = 0; index < PreviewPages; index++) {
            PdfPageRenderResult page = control[index];
            PdfPageRenderResult expected = reference[index];
            ValidateControlPage(page, index + 1);
            ValidateControlPage(expected, index + 1);
            CapabilityDiagnostics += page.CapabilityDiagnostics.Count;
            byte[] controlBytes = page.Bytes ?? throw new InvalidDataException("Public control returned no PNG bytes.");
            byte[] expectedBytes = expected.Bytes ?? throw new InvalidDataException("Fresh public reference returned no PNG bytes.");
            byte[] png = pngs[index];
            if (!png.AsSpan().SequenceEqual(controlBytes) || !controlBytes.AsSpan().SequenceEqual(expectedBytes)) {
                throw new InvalidDataException("Complete stage PNG bytes differ from the fresh public RenderPng control.");
            }
            OfficeRasterImage decoded = Decode(png);
            OfficeRasterImage controlImage = Decode(controlBytes);
            OfficeRasterImage referenceImage = Decode(expectedBytes);
            OfficeRasterImage raster = rasters[index];
            CheckRaster(raster);
            byte[] rgba = raster.GetPixels();
            byte[] decodedRgba = decoded.GetPixels();
            byte[] controlRgba = controlImage.GetPixels();
            byte[] referenceRgba = referenceImage.GetPixels();
            if (!rgba.AsSpan().SequenceEqual(decodedRgba) || !rgba.AsSpan().SequenceEqual(controlRgba) ||
                !controlRgba.AsSpan().SequenceEqual(referenceRgba)) {
                throw new InvalidDataException("Complete stage RGBA pixels differ from the decoded public RenderPng control.");
            }
            bool visibleContent = false;
            for (int pixel = 0; pixel < rgba.Length; pixel += 4) {
                if (rgba[pixel] < 240 || rgba[pixel + 1] < 240 || rgba[pixel + 2] < 240) {
                    visibleContent = true;
                    break;
                }
            }
            if (!visibleContent) throw new InvalidDataException("Preview raster contains no visible business content.");
            string name = "page-" + (index + 1).ToString("0000", CultureInfo.InvariantCulture) + ".png";
            File.WriteAllBytes(Path.Combine(directory, name), png);
            OutputBytes += png.LongLength;
            RasterBytes += rgba.LongLength;
            ComparedPixelBytes += rgba.LongLength;
            MatchedControlPages++;
            pageEvidence.Add(new {
                PageNumber = index + 1, decoded.Width, decoded.Height, EncodedBytes = png.LongLength,
                RgbaBytes = rgba.LongLength, PngSha256 = Hash(png), RgbaSha256 = Hash(rgba),
                CompletePngBytesMatch = true, CompleteRgbaPixelsMatch = true,
                CapabilityDiagnosticCount = page.CapabilityDiagnostics.Count, DiagnosticCount = page.Diagnostics.Count
            });
        }
        RasterDiagnostics = _rasterDiagnostics.Count;
        if (RasterDiagnostics != 0) throw new InvalidDataException("Prepared or result rasterization reported a fidelity diagnostic.");
        VerifiedPages = PreviewPages;
        File.WriteAllText(Path.Combine(directory, "render-stage-evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-render-stages-v1", Operation = _preparedOperation, CorpusProfile,
            SourceSha256, InputBytes, Pages, PreviewPages, Scale, PixelWidth, PixelHeight, HasIndependentSource,
            SourceVerifiedPages, VerifiedPages, OutputBytes, RasterBytes, ComparedPixelBytes,
            ProjectionElements, ProjectionTextElements, ProjectionImageElements, ProjectionShapeElements,
            CapabilityDiagnostics, RasterDiagnostics, MatchedControlPages,
            PreparedInputShape, MeasuredOutputShape,
            CachePolicy = "Fresh operation inputs per sample; process-wide JIT/font/codec state is retained. No process-cold claim.",
            LifetimePolicy = "Prepared inputs and measured outputs are retained through validation and released by the next Prepare or ReleaseResults.",
            MeasurementBoundary = "PowerForge measures Execute only; Prepare, independent source checks, downstream completion, readback and export are outside measurement.",
            Interpretation = "Separate prepared-stage medians are not additive. Full RenderPng includes public batch/diagnostic/result overhead absent from individual boundaries.",
            FidelityBoundary = "Complete PNG bytes and decoded pixels match the public control for this exact source. Independent source text/geometry is separate from independent visual fidelity.",
            PngEncoding = "Optimal compression; no pHYs resolution metadata, matching public PDF RenderPng.",
            Scenes = sceneEvidence, Previews = pageEvidence
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private void ValidateSource(PdfDocument reader) {
        if (reader.Pages.Count != Pages) throw new InvalidDataException("Source page count changed.");
        for (int index = 0; index < reader.Pages.Count; index++) {
            (double width, double height) = reader.Pages[index].GetPageSize();
            if (width != PdfRuntimeCorpus.PageWidth || height != PdfRuntimeCorpus.PageHeight) throw new InvalidDataException("Source page geometry changed.");
            CheckText(reader.Pages[index].ExtractText(), index + 1);
        }
        SourceVerifiedPages = Pages;
    }

    private void ValidateControlPage(PdfPageRenderResult page, int pageNumber) {
        if (!page.Succeeded || page.PageNumber != pageNumber || page.Format != global::OfficeIMO.Drawing.OfficeImageExportFormat.Png ||
            page.Width != PixelWidth || page.Height != PixelHeight || page.CapabilityDiagnostics.Count != 0 || page.Diagnostics.Count != 0) {
            throw new InvalidDataException("Public control changed page order, geometry, success, or diagnostics.");
        }
    }

    private OfficeRasterImage Decode(byte[] png) {
        if (!OfficePngReader.TryDecode(png, out OfficeRasterImage? image) || image is null) throw new InvalidDataException("Preview PNG cannot be decoded.");
        CheckRaster(image);
        return image;
    }

    private void CheckRaster(OfficeRasterImage image) {
        if (image.Width != PixelWidth || image.Height != PixelHeight) throw new InvalidDataException("Preview raster geometry changed.");
    }

    private static void CheckText(string text, int page) {
        string normalized = Normalize(text);
        foreach (string required in PdfRuntimeCorpus.ExpectedText(page, replaced: false, textOnly: false)) {
            if (!normalized.Contains(Normalize(required), StringComparison.Ordinal)) throw new InvalidDataException("Missing business page text: " + required);
        }
        if (normalized.Split(Normalize(PdfRuntimeCorpus.Target(page)), StringSplitOptions.None).Length != 2) {
            throw new InvalidDataException("Business page marker count changed.");
        }
    }

    private void ResetValidationMetrics() {
        SourceVerifiedPages = VerifiedPages = ProjectionElements = ProjectionTextElements = 0;
        ProjectionImageElements = ProjectionShapeElements = CapabilityDiagnostics = RasterDiagnostics = MatchedControlPages = 0;
        OutputBytes = RasterBytes = ComparedPixelBytes = 0L;
    }

    private static string Normalize(string text) => string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
}
