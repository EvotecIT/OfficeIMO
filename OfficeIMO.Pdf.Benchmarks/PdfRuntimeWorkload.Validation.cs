using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

public sealed partial class PdfRuntimeWorkload {
    /// <summary>Checks every observed page, render or saved PDF outside measurement.</summary>
    public void Validate() {
        if (_operation is null) throw new InvalidOperationException("No operation was executed.");
        if (SourceSha256 != Convert.ToHexString(SHA256.HashData(_source))) throw new InvalidDataException("Input changed.");
        if (_inspections is not null) {
            if (_inspections.Length != 4) throw new InvalidDataException("Inspection count changed.");
            foreach (PdfDocumentInfo info in _inspections) {
                if (info.PageCount != Pages || info.Metadata.Title != "OfficeIMO PDF runtime corpus" || info.HasForms || info.HasAnnotations ||
                    info.Pages.Any(page => page.Width != PdfRuntimeCorpus.PageWidth || page.Height != PdfRuntimeCorpus.PageHeight)) {
                    throw new InvalidDataException("Inspection did not preserve the complete page geometry.");
                }
            }
            VerifiedPages = Pages;
        } else if (_renders is not null) {
            ValidateRenders();
        } else if (_output is not null) {
            PdfReadDocument read = PdfReadDocument.Open(_output);
            if (read.Pages.Count != Pages) throw new InvalidDataException("Saved page count changed.");
            for (int index = 0; index < read.Pages.Count; index++) {
                CheckText(read.Pages[index].ExtractText(), index + 1, IsReplacementPage(index + 1));
            }
            if (_edit is not null && (_expectedAffectedSpans == 0 || _edit.AffectedCount != _expectedAffectedSpans)) {
                throw new InvalidDataException("The edit did not affect the prepared region's exact source span count.");
            }
            if (_serialization is not null) {
                bool forward = _operation is "GenerateForward" or "DestinationDirect" or "DestinationCopied";
                if (_serialization.IsForwardOnlyObjectSerialization != forward || forward && _serialization.PeakRetainedObjectBytes != 0L) {
                    throw new InvalidDataException("The requested object serialization contract was not observed.");
                }
            }
            OutputBytes = _output.LongLength;
            VerifiedPages = Pages;
        } else {
            throw new InvalidDataException("Operation produced no observable result.");
        }
    }

    /// <summary>Checks and retains Poppler source word geometry for later independent edit validation.</summary>
    public void ValidateIndependentSource(string xmlPath) {
        IndependentPage[] pages = ReadIndependentPages(xmlPath);
        CheckIndependentPages(pages, replacement: false);
        _independentSource = pages;
    }

    /// <summary>Checks exact saved output through independent per-page text and neighboring word geometry.</summary>
    public void ValidateIndependentResult(string xmlPath) {
        if (_output is null || _operation is null) throw new InvalidOperationException("No saved PDF result.");
        IndependentPage[] pages = ReadIndependentPages(xmlPath);
        CheckIndependentPages(pages, _operation == "ReplaceRegion");
        if (_edit is null) return;
        IndependentPage[] source = _independentSource ?? throw new InvalidOperationException("Independent source was not prepared.");
        for (int index = 0; index < pages.Length; index++) {
            bool targetPage = index + 1 == TargetPage;
            Word[] expectedNeighbors = source[index].Words.Where(word => !targetPage || !IsTargetWord(word.Text)).ToArray();
            Word[] actualNeighbors = pages[index].Words.Where(word => !targetPage || !IsResultTargetWord(word.Text)).ToArray();
            if (expectedNeighbors.Length != actualNeighbors.Length) throw new InvalidDataException("Independent neighboring word count changed.");
            for (int word = 0; word < expectedNeighbors.Length; word++) {
                CheckWord(actualNeighbors[word], expectedNeighbors[word], 0D);
            }
            if (targetPage && _operation == "MoveRegion") {
                Word[] original = source[index].Words.Where(word => IsTargetWord(word.Text)).ToArray();
                Word[] moved = pages[index].Words.Where(word => IsTargetWord(word.Text)).ToArray();
                if (original.Length != 2 || moved.Length != 2) throw new InvalidDataException("Independent moved text is incomplete.");
                for (int word = 0; word < original.Length; word++) CheckWord(moved[word], original[word], 8D);
            }
        }
    }

    /// <summary>Retains exact measured artifacts and fingerprints; called after validation, outside timing.</summary>
    public void ExportEvidence(string directory) {
        Directory.CreateDirectory(directory);
        var hashes = new Dictionary<string, string>(StringComparer.Ordinal);
        if (_output is not null) {
            const string name = "result.pdf";
            File.WriteAllBytes(Path.Combine(directory, name), _output);
            hashes[name] = Convert.ToHexString(SHA256.HashData(_output));
        }
        if (_renders is not null) {
            foreach (PdfPageRenderResult render in _renders) {
                byte[] bytes = render.Bytes ?? throw new InvalidDataException("Render bytes are missing.");
                string name = "page-" + render.PageNumber.ToString("0000", CultureInfo.InvariantCulture) + (_operation == "RenderSvg" ? ".svg" : ".png");
                File.WriteAllBytes(Path.Combine(directory, name), bytes);
                hashes[name] = Convert.ToHexString(SHA256.HashData(bytes));
            }
        }
        File.WriteAllText(Path.Combine(directory, "evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-runtime-v1", Operation = _operation, CorpusProfile, Pages, TargetPage, SourceSha256,
            InputBytes, OutputBytes, VerifiedPages, AffectedTextSpans, ExpectedAffectedTextSpans = _expectedAffectedSpans,
            InspectionCalls = _inspections?.Length ?? 0,
            InspectedPages = _inspections?.Select(info => info.PageCount).ToArray(),
            PeakRetainedPageContentBytes, PeakRetainedObjectBytes, LargestSerializedObjectBytes,
            OutputSha256 = hashes
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private void ValidateRenders() {
        int expected = Math.Min(Pages, _operation == "RenderSvg" ? 12 : 4);
        if (_renders!.Count != expected) throw new InvalidDataException("Rendered page count changed.");
        long outputBytes = 0;
        for (int index = 0; index < _renders.Count; index++) {
            PdfPageRenderResult render = _renders[index];
            byte[] bytes = render.Bytes ?? throw new InvalidDataException("Rendering returned no bytes.");
            if (!render.Succeeded || render.PageNumber != index + 1 || render.CapabilityDiagnostics.Count != 0) {
                throw new InvalidDataException("A render failed, reordered pages, or simplified the corpus.");
            }
            if (_operation == "RenderSvg") {
                using var stream = new MemoryStream(bytes, writable: false);
                XElement svg = XDocument.Load(stream).Root ?? throw new InvalidDataException("SVG has no root.");
                if (svg.Name.LocalName != "svg" || (string?)svg.Attribute("viewBox") != "0 0 600 800" || !svg.Descendants().Any()) {
                    throw new InvalidDataException("SVG has incorrect geometry or no content.");
                }
                // Shaped PDF glyphs can use private-use SVG characters with a matching embedded cmap.
                // Qualify logical content separately, then require the complete public page projection.
                OfficeDrawing drawing = PdfReadDocument.Open(_source).Pages[index].ToDrawing();
                CheckText(string.Join(' ', drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text)), index + 1, replaced: false);
                byte[] reference = OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D);
                if (!bytes.AsSpan().SequenceEqual(reference)) throw new InvalidDataException("Batch SVG differs from the complete public page drawing export.");
                // Non-interpolated images may be exported as exact nearest-neighbor vector rectangles.
                if (!drawing.Elements.OfType<OfficeDrawingImage>().Any()) throw new InvalidDataException("The SVG reference projection omitted the corpus image.");
            } else {
                if (!OfficePngReader.TryDecode(bytes, out OfficeRasterImage? image) || image is null || image.Width != 384 || image.Height != 512) {
                    throw new InvalidDataException("PNG did not decode at the requested geometry.");
                }
                byte[] pixels = image.GetPixels();
                if (!Enumerable.Range(0, pixels.Length / 4).Any(pixel => pixels[pixel * 4] < 240 || pixels[pixel * 4 + 1] < 240 || pixels[pixel * 4 + 2] < 240)) {
                    throw new InvalidDataException("PNG contains no visible corpus content.");
                }
            }
            outputBytes += bytes.LongLength;
        }
        OutputBytes = outputBytes;
        VerifiedPages = expected;
    }

    private bool IsReplacementPage(int page) => _operation == "ReplaceRegion" && page == TargetPage;
    private bool IsTargetWord(string text) => text is "Editable" || text == "marker" + PdfRuntimeCorpus.Number(TargetPage);
    private bool IsResultTargetWord(string text) => IsTargetWord(text) || _operation == "ReplaceRegion" && text == "Edited";

    private void CheckText(string text, int page, bool replaced) {
        string normalized = Normalize(text);
        foreach (string required in PdfRuntimeCorpus.ExpectedText(page, replaced, _textOnly)) {
            if (!normalized.Contains(Normalize(required), StringComparison.Ordinal)) throw new InvalidDataException("Missing required page content: " + required);
        }
        string target = Normalize(replaced ? PdfRuntimeCorpus.Replacement(page) : PdfRuntimeCorpus.Target(page));
        if (normalized.Split(target, StringSplitOptions.None).Length != 2) throw new InvalidDataException("Target text count changed.");
        if (replaced && normalized.Contains(Normalize(PdfRuntimeCorpus.Target(page)), StringComparison.Ordinal)) throw new InvalidDataException("Replaced text survives.");
    }

    private void CheckIndependentPages(IndependentPage[] pages, bool replacement) {
        if (pages.Length != Pages) throw new InvalidDataException("Independent page count changed.");
        for (int index = 0; index < pages.Length; index++) {
            IndependentPage page = pages[index];
            if (page.Width != PdfRuntimeCorpus.PageWidth || page.Height != PdfRuntimeCorpus.PageHeight) throw new InvalidDataException("Independent page geometry changed.");
            CheckText(string.Join(' ', page.Words.Select(word => word.Text)), index + 1, replacement && index + 1 == TargetPage);
            string documentCount = ((index + 1) * 37).ToString(CultureInfo.InvariantCulture);
            string amount = ((index + 1) * 101.25D).ToString("0.00", CultureInfo.InvariantCulture);
            if (!_textOnly && (page.Words.Count(word => word.Text == documentCount) != 1 || page.Words.Count(word => word.Text == amount) != 1)) {
                throw new InvalidDataException("Independent numeric table values changed.");
            }
        }
    }

    private static void CheckWord(Word actual, Word expected, double deltaX) {
        if (actual.Text != expected.Text || Math.Abs(actual.XMin - expected.XMin - deltaX) > 0.05D ||
            Math.Abs(actual.XMax - expected.XMax - deltaX) > 0.05D || Math.Abs(actual.YMin - expected.YMin) > 0.05D || Math.Abs(actual.YMax - expected.YMax) > 0.05D) {
            throw new InvalidDataException("Independent word text or geometry changed: " + expected.Text);
        }
    }

    private static IndependentPage[] ReadIndependentPages(string path) {
        using var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, XmlResolver = null });
        return XDocument.Load(reader).Descendants().Where(element => element.Name.LocalName == "page")
            .Select(page => new IndependentPage(Number(page, "width"), Number(page, "height"),
                page.Descendants().Where(element => element.Name.LocalName == "word")
                    .Select(word => new Word(word.Value, Number(word, "xMin"), Number(word, "yMin"), Number(word, "xMax"), Number(word, "yMax"))).ToArray())).ToArray();
    }

    private static double Number(XElement element, string attribute) => double.Parse(
        (string?)element.Attribute(attribute) ?? throw new InvalidDataException("Independent geometry is missing."), CultureInfo.InvariantCulture);
    private static string Normalize(string text) => string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
    private sealed record Word(string Text, double XMin, double YMin, double XMax, double YMax);
    private sealed record IndependentPage(double Width, double Height, Word[] Words);
}
