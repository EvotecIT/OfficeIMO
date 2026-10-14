using System.Security.Cryptography;
using System.IO.Compression;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument.Benchmarks;

/// <summary>A bounded Draw projection workload consumed by the shared PowerForge runner.</summary>
public sealed class OdgProjectionWorkload {
    private readonly OdgDocument _document;
    private readonly int _pageCount;
    private readonly bool _fields;
    private readonly string[] _sourceParts;
    private OdfConversionResult<IReadOnlyList<OfficeDrawing>>? _result;

    /// <summary>Creates pages sharing one master/layout; input preparation is outside measurement.</summary>
    public OdgProjectionWorkload(int pageCount, bool fields) {
        if (pageCount < 1 || pageCount > 1000) throw new ArgumentOutOfRangeException(nameof(pageCount));
        _pageCount = pageCount; _fields = fields;
        _document = OdgDocument.Create();
        var first = _document.AddPage("Page1");
        if (fields) AddFields(first);
        for (int index = 1; index < pageCount; index++) _document.ClonePage(0, "Page" + (index + 1));
        InputSha256 = Convert.ToHexString(SHA256.HashData(_document.ToBytes()));
        _sourceParts = SourceParts();
    }

    /// <summary>Hash of the prepared source package.</summary>
    public string InputSha256 { get; }
    /// <summary>Number of output pages verified after the operation.</summary>
    public int VerifiedPages { get; private set; }
    /// <summary>Checksum of all page ordinals and their fully checked visible text.</summary>
    public long ContentChecksum { get; private set; }

    /// <summary>Projects the entire document using the public bounded conversion contract.</summary>
    public void Execute() => _result = _document.ToDrawings(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);

    /// <summary>Checks every output page and leaves input generation/readback outside timing.</summary>
    public void Validate() {
        if (_result == null || _result.Value.Count != _pageCount) throw new InvalidOperationException("Projection did not return every source page.");
        long checksum = 0;
        for (int index = 0; index < _pageCount; index++) {
            var drawing = _result.Value[index];
            if (Math.Abs(drawing.Width - OdfLength.Centimeters(21).ToPoints()) > 0.000001 ||
                Math.Abs(drawing.Height - OdfLength.Centimeters(29.7).ToPoints()) > 0.000001)
                throw new InvalidOperationException("Projection changed page dimensions.");
            string[] text = drawing.Elements.OfType<OfficeDrawingRichText>().Select(element => element.PlainText).ToArray();
            string expected = _fields ? $"Page {index + 1} of {_pageCount}" : string.Empty;
            if (_fields ? text.Length != 1 || text[0] != expected : text.Length != 0)
                throw new InvalidOperationException("Projection returned stale or missing page fields.");
            checksum += (long)(index + 1) * (expected.Length + 1);
        }
        if (!SourceParts().SequenceEqual(_sourceParts)) throw new InvalidOperationException("Projection changed source XML or field caches.");
        VerifiedPages = _pageCount; ContentChecksum = checksum;
    }

    /// <summary>Releases prior scenes before another measurement.</summary>
    public void ReleaseResults() => _result = null;

    private static void AddFields(OdgPage page) {
        var paragraph = page.MasterShapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 14, 2), "Page ", "Header").Paragraphs[0];
        paragraph.FontFamily = "Arial"; paragraph.FontSize = OdfLength.Points(12); paragraph.Color = OdfColor.Parse("#000000");
        paragraph.AddField(OdfTextFieldKind.PageNumber, "stale").NumberFormat = "1";
        paragraph.AddText(" of "); paragraph.AddField(OdfTextFieldKind.PageCount, "99").NumberFormat = "1";
    }

    private string[] SourceParts() {
        using var package = new ZipArchive(_document.ToStream(), ZipArchiveMode.Read);
        return new[] { "content.xml", "styles.xml", "meta.xml" }.Select(part => {
            using var reader = new StreamReader(package.GetEntry(part)!.Open());
            return reader.ReadToEnd();
        }).ToArray();
    }
}
