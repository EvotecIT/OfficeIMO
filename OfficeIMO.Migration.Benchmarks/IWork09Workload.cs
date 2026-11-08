using System.Globalization;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Measures synthetic legacy XML with distinct inline uses of one shared raster, not native producer coverage.</summary>
internal sealed class IWork09Workload : MigrationWorkload {
    private static readonly byte[] Raster = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
    private const string Tail = " tail";
    private IWorkSourceDocument? _source;
    private IWorkPagesProjection? _projection;
    private PagesToWordResult? _converted;

    internal IWork09Workload(int units) : base(CreateInput(units), units) { }

    public override void Execute(string operation) {
        _source = IWorkSourceDocument.Load(Input);
        if (operation == "LoadInspect") _projection = _source.ReadPages();
        else if (operation == "ConvertSave") {
            _converted = _source.ToWordResult(new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true
            });
            _projection = _converted.Projection;
            using var output = new MemoryStream();
            _converted.Value.Save(output);
            Output = output.ToArray();
        } else throw new ArgumentOutOfRangeException(nameof(operation));
    }

    public override void Validate() {
        if (_source == null || _projection == null) throw new InvalidOperationException("The legacy XML operation did not produce a source projection.");
        Require(_source.Kind == IWorkDocumentKind.Pages && _source.Records.Count == 0, "source generation");
        Require(_source.Entries.Count == 2, "captured entry count");
        Require(_source.Entries.Single(entry => entry.Path == "shared.png").GetBytes().SequenceEqual(Raster), "captured raster");
        Require(_projection.Body.Paragraphs.Count == Units && _projection.Images.Count == Units
            && _projection.Drawables.Count == Units, "projected occurrence count");
        Require(_projection.Sections.Count == 0 && _projection.Tables.Count == 0
            && _projection.TextBoxes.Count == 0, "unexpected projected content");
        var identities = new HashSet<ulong>();
        var attachments = new HashSet<ulong>();
        int offset = 0;
        for (int index = 0; index < Units; index++) {
            IWorkTextParagraph paragraph = _projection.Body.Paragraphs[index];
            Require(paragraph.Text == Prefix(index) + Tail, "body text " + index);
            IWorkInlineObject anchor = paragraph.Runs.Single(run => run.InlineObject != null).InlineObject!;
            IWorkImageAsset image = _projection.Images[index];
            Require(image.PackagePath == "shared.png" && image.MediaType == "image/png"
                && image.PixelWidth == 1 && image.PixelHeight == 1 && image.GetBytes().SequenceEqual(Raster), "raster occurrence " + index);
            Require(image.Geometry is { } frame && frame.LeftPoints == 0 && frame.TopPoints == 0
                && frame.WidthPoints == Width(index) && frame.HeightPoints == Height(index)
                && frame.RotationDegrees == 0 && !image.HasMask, "source frame " + index);
            Require(image.SourceIdentity != null && identities.Add(image.SourceIdentity.RecordIdentifier)
                && anchor.Drawable.RecordIdentifier == image.SourceIdentity.RecordIdentifier
                && attachments.Add(anchor.Attachment.RecordIdentifier)
                && anchor.CharacterOffset == offset + Prefix(index).Length, "source anchor " + index);
            Require(_projection.Drawables[index].Kind == IWorkPagesDrawableKind.Image
                && _projection.Drawables[index].Image!.SourceIdentity!.RecordIdentifier == image.SourceIdentity!.RecordIdentifier,
                "drawable order " + index);
            offset += paragraph.Text.Length + 2; // One attachment marker and one paragraph boundary.
        }

        string[] expected = _converted == null
            ? new[] { "IWORK_LEGACY_XML_SOURCE_READ:None", "IWORK_LEGACY_PAGES_CONTENT_LIMITED:Omission" }
            : new[] { "IWORK_LEGACY_XML_SOURCE_READ:None", "IWORK_LEGACY_PAGES_CONTENT_LIMITED:Omission", "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION:Omission" };
        string[] actual = _converted == null
            ? _source.Diagnostics.Concat(_projection.Diagnostics).Select(item => item.Code + ":" + item.LossKind).ToArray()
            : _converted.Report.Diagnostics.Select(item => item.Code + ":" + item.LossKind).ToArray();
        Require(actual.OrderBy(item => item, StringComparer.Ordinal)
            .SequenceEqual(expected.OrderBy(item => item, StringComparer.Ordinal)), "diagnostic contract");
        DiagnosticDigest = string.Join("|", actual.OrderBy(item => item, StringComparer.Ordinal));
        if (_converted != null) {
            Require(!_converted.IsVisualFallback && _converted.Report.IsPartialEditableReconstruction, "explicit partial editable representation");
            Require(Output is { Length: > 0 }, "saved destination");
            using var bytes = new MemoryStream(Output!, writable: false);
            using WordDocument reopened = WordDocument.Load(bytes, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
            Require(!report.HasLoss && report.Diagnostics.Count == 0, "destination reopen report");
            InspectDestination(reopened.OpenXmlDocument);
        }
        VerifiedUnits = Units;
    }

    public override void ReleaseResults() {
        _converted?.Dispose();
        _converted = null;
        _projection = null;
        _source?.Dispose();
        _source = null;
        base.ReleaseResults();
    }

    private void InspectDestination(WordprocessingDocument document) {
        Require(!new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Any(), "Open XML validation");
        MainDocumentPart main = document.MainDocumentPart!;
        W.Paragraph[] paragraphs = main.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Require(paragraphs.Length == Units, "saved paragraph count");
        var referencedParts = new HashSet<Uri>();
        for (int index = 0; index < Units; index++) {
            W.Paragraph paragraph = paragraphs[index];
            Require(paragraph.InnerText == Prefix(index) + Tail, "saved body text " + index);
            W.Run[] runs = paragraph.Elements<W.Run>().Where(run => run.InnerText.Length > 0 || run.Descendants<W.Drawing>().Any()).ToArray();
            Require(runs.Length == 3 && runs[0].InnerText == Prefix(index) && runs[1].Descendants<W.Drawing>().Count() == 1
                && runs[2].InnerText == Tail, "saved image order " + index);
            WP.Inline inline = paragraph.Descendants<WP.Inline>().Single();
            WP.Extent extent = inline.Extent!;
            Require(extent.Cx!.Value == Width(index) * 12700L && extent.Cy!.Value == Height(index) * 12700L, "saved frame " + index);
            string relationship = inline.Descendants<A.Blip>().Single().Embed!.Value!;
            Require(main.GetPartById(relationship) is ImagePart, "saved image relationship " + index);
            var image = (ImagePart)main.GetPartById(relationship);
            referencedParts.Add(image.Uri);
            using var payload = new MemoryStream();
            using (Stream stream = image.GetStream()) stream.CopyTo(payload);
            Require(payload.ToArray().SequenceEqual(Raster), "saved raster bytes " + index);
        }
        Require(referencedParts.Count == main.ImageParts.Count(), "unreferenced destination raster");
    }

    private static string Prefix(int index) => "Picture " + index.ToString("D5", CultureInfo.InvariantCulture);
    private static int Width(int index) => 72 + index % 4 * 3;
    private static int Height(int index) => 36 + index % 3 * 3;

    private static void Require(bool condition, string boundary) {
        if (!condition) throw new InvalidDataException("The shared legacy XML raster workload changed " + boundary + ".");
    }

    private static byte[] CreateInput(int units) {
        if (units < 1) throw new ArgumentOutOfRangeException(nameof(units));
        var xml = new StringBuilder("<sl:document xmlns:sl=\"http://developer.apple.com/namespaces/sl\" xmlns:sf=\"http://developer.apple.com/namespaces/sf\" xmlns:sfa=\"http://developer.apple.com/namespaces/sfa\" sl:version=\"92008102400\"><sf:text-storage><sf:attachments>");
        for (int index = 0; index < units; index++) {
            xml.Append(CultureInfo.InvariantCulture, $"<sf:attachment sfa:ID=\"a{index}\"><sf:media sfa:ID=\"m{index}\"><sf:geometry><sf:size sfa:w=\"{Width(index)}\" sfa:h=\"{Height(index)}\"/><sf:position sfa:x=\"0\" sfa:y=\"0\"/></sf:geometry><sf:content><sf:image-media><sf:filtered-image><sf:unfiltered><sf:data sf:path=\"shared.png\"/></sf:unfiltered></sf:filtered-image></sf:image-media></sf:content></sf:media></sf:attachment>");
        }
        xml.Append("</sf:attachments><sf:text-body>");
        for (int index = 0; index < units; index++)
            xml.Append(CultureInfo.InvariantCulture, $"<sf:p>{Prefix(index)}<sf:attachment-ref sfa:IDREF=\"a{index}\"/>{Tail}</sf:p>");
        xml.Append("</sf:text-body></sf:text-storage></sl:document>");
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry indexEntry = zip.CreateEntry("index.xml", CompressionLevel.NoCompression);
            indexEntry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using (var writer = new StreamWriter(indexEntry.Open(), new UTF8Encoding(false))) writer.Write(xml);
            ZipArchiveEntry rasterEntry = zip.CreateEntry("shared.png", CompressionLevel.NoCompression);
            rasterEntry.LastWriteTime = indexEntry.LastWriteTime;
            using Stream image = rasterEntry.Open();
            image.Write(Raster);
        }
        return bytes.ToArray();
    }
}
