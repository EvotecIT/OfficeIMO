using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Threading;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {

    public sealed class WordSvgImageExportTests {
        private const string SvgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string RedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='40'><rect width='40' height='40' fill='#ff0000'/></svg>";

        [Theory]
        [InlineData("body", false)]
        [InlineData("body", true)]
        [InlineData("cell", true)]
        [InlineData("header", true)]
        [InlineData("footer", true)]
        public void SavedSvgImagesRenderInPngAndPdf(string story, bool extension) {
            using WordDocument source = CreateDocument(story, extension, RedSvg);
            using MemoryStream package = new System.IO.MemoryStream(source.ToBytes());
            using WordDocument document = WordDocument.Load(package);

            OfficeImageExportResult preview = Assert.Single(document.ExportImages(OfficeImageExportFormat.Png));
            AssertColor(preview.Bytes, red: true);
            PdfDocumentConversionResult conversion = document.ConvertToPdfWithEvidence();
            Assert.DoesNotContain(conversion.Warnings, warning => warning.LossKind == OfficeConversionLossKind.Omission);
            PdfDocument pdf = PdfDocument.Load(new System.IO.MemoryStream(conversion.Value.ToBytes()));
            AssertColor(Assert.Single(pdf.ExportImages(OfficeImageExportFormat.Png)).Bytes, red: true);
        }

        [Theory]
        [InlineData("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='40'><foreignObject width='40' height='40'/></svg>")]
        [InlineData("<svg this is broken")]
        [InlineData(null)]
        public void UnrenderableOrMissingSvgExtensionUsesBitmapFallback(string? svg) {
            using WordDocument source = CreateDocument("body", extension: true, svg);
            using MemoryStream package = new System.IO.MemoryStream(source.ToBytes());
            using WordDocument document = WordDocument.Load(package);
            AssertColor(Assert.Single(document.ExportImages(OfficeImageExportFormat.Png)).Bytes, red: false);
            PdfDocumentConversionResult conversion = document.ConvertToPdfWithEvidence();
            Assert.DoesNotContain(conversion.Warnings, warning => warning.LossKind == OfficeConversionLossKind.Omission);
            PdfDocument pdf = PdfDocument.Load(new System.IO.MemoryStream(conversion.Value.ToBytes()));
            AssertColor(Assert.Single(pdf.ExportImages(OfficeImageExportFormat.Png)).Bytes, red: false);
        }

        [Fact]
        public void UnsupportedSvgWithoutFallbackRetainsOmissionDiagnostic() {
            using WordDocument document = CreateDocument("body", extension: false,
                "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='40'><foreignObject width='40' height='40'/></svg>");
            Assert.Contains(document.ConvertToPdfWithEvidence().Warnings, warning =>
                warning.Code == "NativeBodyImageUnsupported" && warning.LossKind == OfficeConversionLossKind.Omission);
        }

        [Fact]
        public void SvgPreparationAndResourceExtractionHonorCancellation() {
            using CancellationTokenSource cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            byte[] svg = Encoding.UTF8.GetBytes(RedSvg);
            OperationCanceledException preparation = Assert.Throws<OperationCanceledException>(() =>
                PdfDocument.TryPrepareImageBytes(svg, cancellation.Token, out _, out _, out _, out _));
            Assert.Equal(cancellation.Token, preparation.CancellationToken);
            using WordDocument document = CreateDocument("body", extension: true, RedSvg);
            Assert.Throws<OperationCanceledException>(() =>
                document.Paragraphs.Last().Image!.GetRenderingImageCandidates(cancellation.Token).ToArray());
        }

        [Fact]
        public void VmlImageExtractionStopsBeforeNextReadWhenCancelled() {
            using CancellationTokenSource cancellation = new CancellationTokenSource();
            using CancelAfterReadStream stream = new CancelAfterReadStream(cancellation);
            Assert.Throws<OperationCanceledException>(() => WordPdfConverterExtensions.TryReadNativeVmlImageBytes(
                stream, WordPdfConverterExtensions.NativeVmlImageMaxBytes, out _, cancellation.Token));
            Assert.Equal(1, stream.ReadCalls);
        }

        [Fact]
        public void WatermarkImageExtractionHonorsCancellation() {
            string path = Path.Combine(Path.GetTempPath(), "svg-watermark-" + Guid.NewGuid().ToString("N") + ".png");
            try {
                File.WriteAllBytes(path, OfficePngWriter.Encode(new OfficeRasterImage(40, 40)));
                using WordDocument document = WordDocument.Create();
                WordHeader header = document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
                WordWatermark watermark = header.AddWatermark(WordWatermarkStyle.Image, path);
                Assert.Throws<OperationCanceledException>(() =>
                    watermark.TryGetImageBytes(out _, out _, new CancellationToken(true)));
            } finally {
                File.Delete(path);
            }
        }

        private sealed class CancelAfterReadStream : MemoryStream {
            private readonly CancellationTokenSource _cancellation;
            internal CancelAfterReadStream(CancellationTokenSource cancellation) : base(new byte[200_000]) => _cancellation = cancellation;
            internal int ReadCalls { get; private set; }
            public override int Read(byte[] buffer, int offset, int count) {
                ReadCalls++;
                int read = base.Read(buffer, offset, count);
                _cancellation.Cancel();
                return read;
            }
        }

        private static WordDocument CreateDocument(string story, bool extension, string? svg) {
            WordDocument document = WordDocument.Create();
            document.AddParagraph("SVG preview");
            WordParagraph paragraph = story == "cell" ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] :
                story == "header" ? document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph() :
                story == "footer" ? document.Sections[0].GetOrCreateFooter(WordHeaderFooterType.Default).AddParagraph() :
                document.AddParagraph();
            byte[] bytes = extension ? OfficePngWriter.Encode(new OfficeRasterImage(40, 40, OfficeColor.FromRgb(0, 0, 255))) :
                Encoding.UTF8.GetBytes(svg!);
            WordImage image = paragraph.AddImage(new MemoryStream(bytes), extension ? "fallback.png" : "image.svg", 40, 40).Image!;
            if (extension) {
                OpenXmlPart owner = WordPartOwnership.Resolve(document, paragraph._paragraph);
                string id = "missingSvg";
                if (svg != null) {
                    ImagePart part = owner.AddNewPart<ImagePart>("image/svg+xml");
                    part.FeedData(new MemoryStream(Encoding.UTF8.GetBytes(svg)));
                    id = owner.GetIdOfPart(part);
                }
                OpenXmlUnknownElement svgBlip = new OpenXmlUnknownElement("asvg", "svgBlip", SvgNamespace);
                svgBlip.SetAttribute(new OpenXmlAttribute("r", "embed", RelationshipNamespace, id));
                image._Image.Descendants<A.Blip>().Single().GetFirstChild<A.BlipExtensionList>()!.Append(
                    new A.BlipExtension(svgBlip) { Uri = "{96DAC541-7B7A-43D3-8B79-8F7C33B92B69}" });
            }
            return document;
        }

        private static void AssertColor(byte[] png, bool red) {
            Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? raster));
            byte[] pixels = raster!.GetPixels();
            int expected = 0, fallback = 0;
            for (int i = 0; i < pixels.Length; i += 4) {
                if (pixels[i] > 220 && pixels[i + 1] < 30 && pixels[i + 2] < 30) expected++;
                if (pixels[i] < 30 && pixels[i + 1] < 30 && pixels[i + 2] > 220) fallback++;
            }
            Assert.True(red ? expected > 300 : fallback > 300, $"Missing {(red ? "SVG red" : "fallback blue")} image: red={expected}, blue={fallback}");
            Assert.Equal(0, red ? fallback : expected);
        }
    }
}
