using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Features;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Threading;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public class WordImageOptimizationTests {
    [Fact]
    public void AnalysisUsesPerImageLimitWithoutRetainingTransactionCandidates() {
        using WordDocument word = WordDocument.Create();
        var first = Insert(word.AddParagraph(), 96, 48);
        var second = Insert(word.AddParagraph(), 96, 48);
        byte[] before = first.ToBytes();
        var options = new WordImageOptimizationOptions { KeepOriginalWhenNotSmaller = false };
        var baseline = word.AnalyzeImageOptimization(options);
        Assert.Equal(2, baseline.OptimizedCount);
        options.MaxStagedBytes = baseline.Images.Max(item => item.OriginalBytes + item.FinalBytes) + 1;
        Assert.True(options.MaxStagedBytes < baseline.Images.Sum(item => item.OriginalBytes + item.FinalBytes));
        var report = word.AnalyzeImageOptimization(options);
        Assert.Equal(baseline.BytesSaved, report.BytesSaved);
        Assert.False(report.Applied);
        Assert.Equal(before, first.ToBytes());
        Assert.Equal(before, second.ToBytes());
        Assert.Throws<InvalidDataException>(() => word.OptimizeImages(options));
        Assert.Equal(before, first.ToBytes());
        Assert.Equal(before, second.ToBytes());
        options.MaxImageBytes = before.Length - 1;
        Assert.Throws<InvalidDataException>(() => word.AnalyzeImageOptimization(options));
    }

    [Fact]
    public void SignedAndReadOnlyDocumentsAllowAnalysisButBlockMutation() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph(), 96, 48);
        byte[] before = image.ToBytes();
        word._wordprocessingDocument.AddNewPart<DigitalSignatureOriginPart>();
        Assert.Equal(1, word.AnalyzeImageOptimization().OptimizedCount);
        Assert.Throws<WordSignatureSavePolicyException>(() => word.OptimizeImages());
        Assert.Equal(before, image.ToBytes());
        using var bytes = new MemoryStream(word.ToBytes(options: new WordSaveOptions {
            SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation
        }));
        using WordDocument readOnly = WordDocument.Load(bytes, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        Assert.Equal(1, readOnly.AnalyzeImageOptimization().OptimizedCount);
        Assert.Throws<InvalidOperationException>(() => readOnly.OptimizeImages());
    }

    [Fact]
    public void NativeCropThatCollapsesAtBinaryPrecisionIsRejected() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph(), 96, 48);
        image.CropLeft = 49998; image.CropRight = 50001;
        Assert.Throws<NotSupportedException>(() => word.ToBytes(WordFileFormat.Doc));
    }

    [Fact]
    public void CancellationDuringCarrierReplacementRestoresOriginalRelationshipsAndBytes() {
        using WordDocument word = WordDocument.Create();
        var jpeg = Insert(word.AddParagraph(), 96, 48);
        byte[] originalJpeg = jpeg.ToBytes();
        byte[] originalBmp = CreateBmp();
        using var stream = new MemoryStream(originalBmp);
        var bmp = word.AddParagraph().InsertImage(stream, "source.bmp", 96, 48);
        string id = bmp.RelationshipId!;
        var originalPart = word._wordprocessingDocument.MainDocumentPart!.GetPartById(id);
        word._wordprocessingDocument.AddPartEventsFeature();
        using var cancellation = new CancellationTokenSource();
        word._wordprocessingDocument.Features.Get<IPartEventsFeature>()!.Change += args => {
            if (args.Type == EventType.Deleting && args.Argument == originalPart) cancellation.Cancel();
        };
        Assert.Throws<OperationCanceledException>(() => word.OptimizeImages(new() {
            KeepOriginalWhenNotSmaller = false, Mode = OfficeImageOptimizationMode.DownsampleAndRecompress
        }, cancellation.Token));
        Assert.Equal(originalJpeg, jpeg.ToBytes());
        Assert.Equal(originalBmp, bmp.ToBytes());
        Assert.Equal("image/bmp", bmp.ContentType);
        Assert.Same(originalPart, word._wordprocessingDocument.MainDocumentPart.GetPartById(id));
        Assert.Empty(word._wordprocessingDocument.MainDocumentPart.HeaderParts);
        using var saved = new MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(2, reopened.Images.Count);
    }
    [Theory]
    [InlineData("gif")]
    [InlineData("bmp")]
    public void FormatConversionRetainsRelationshipsAndLiveImageHandles(string format) {
        using WordDocument word = WordDocument.Create();
        byte[] bytes = format == "gif" ? Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=") : CreateBmp();
        using var stream = new MemoryStream(bytes);
        var image = word.AddParagraph("Converted media").InsertImage(stream, "source." + format, 96, 48);
        string id = image.RelationshipId!;
        word.AddHeadersAndFooters();
        var headerPart = word._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single();
        ImagePart originalPart = (ImagePart)word._wordprocessingDocument.MainDocumentPart.GetPartById(id);
        headerPart.AddPart(originalPart, "rIdShared");
        // A second story refers to the same image. It retains its own relationship ID after conversion.
        var drawing = (DocumentFormat.OpenXml.Wordprocessing.Drawing)image._Image.CloneNode(true);
        drawing.Descendants<A.Blip>().Single().Embed = "rIdShared";
        headerPart.Header.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(drawing)));
        string xml = word._wordprocessingDocument.MainDocumentPart.Document.OuterXml;
        var report = word.OptimizeImages(new() { KeepOriginalWhenNotSmaller = false });
        Assert.Equal(1, report.OptimizedCount);
        Assert.Equal(2, Assert.Single(report.Images).ReferenceCount);
        Assert.Equal(id, image.RelationshipId);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(OfficeImageFormat.Png, OfficeImageReader.Identify(image.ToBytes()).Format);
        Assert.Equal(xml, word._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Single(word._wordprocessingDocument.MainDocumentPart.HeaderParts);
        using var saved = new MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal("image/png", reopened.Images[0].ContentType);
        Assert.Single(reopened.AnalyzeImageOptimization().Images);
    }

    [Fact]
    public void NestedGroupUsesOuterDrawingScaleAndPreservesXml() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph(), 96, 48);
        var picture = image._Image.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Single();
        var data = picture.Parent!;
        picture.Remove();
        var group = new DocumentFormat.OpenXml.Office2010.Word.DrawingGroup.WordprocessingGroup(
            new DocumentFormat.OpenXml.Office2010.Word.DrawingGroup.GroupShapeProperties(new A.TransformGroup(
                new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 457200 },
                new A.ChildOffset { X = 0, Y = 0 }, new A.ChildExtents { Cx = 914400, Cy = 457200 })), picture);
        data.AppendChild(group);
        image._Image.Inline!.Extent!.Cx = 2 * 914400L;
        image._Image.Inline.Extent.Cy = 914400L;
        string before = image._Image.OuterXml;
        var item = Assert.Single(word.OptimizeImages(new() { TargetDpi = 96, KeepOriginalWhenNotSmaller = false }).Images);
        Assert.Equal(192, item.Final.Width);
        Assert.Equal(96, item.Final.Height);
        Assert.Equal(before, image._Image.OuterXml);
    }

    [Fact]
    public void UnknownPlacementPreservesDownsamplingButAllowsExplicitRecompression() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph(), 96, 48);
        image._Image.Descendants<A.FillRectangle>().Single().Left = 25000;
        byte[] before = image.ToBytes();
        Assert.Equal(WordImageOptimizationStatus.UnknownPlacement, Assert.Single(word.OptimizeImages().Images).Status);
        Assert.Equal(before, image.ToBytes());
        var item = Assert.Single(word.OptimizeImages(new() { Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 25 }).Images);
        Assert.Equal(WordImageOptimizationStatus.Optimized, item.Status);
        Assert.Equal(800, item.Final.Width);
    }

    [Fact]
    public void CompleteInventoryIncludesMultipleBodyPicturesAndHeaderWithoutChangingXml() {
        using WordDocument word = WordDocument.Create();
        word.AddHeadersAndFooters();
        var paragraph = word.AddParagraph("Before and after pictures");
        Insert(paragraph, 96, 48);
        Insert(paragraph, 192, 96);
        Insert(word.Header.Default!.AddParagraph("Header"), 96, 48);
        Insert(word.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], 96, 48);
        string before = word._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        var policy = new WordImageOptimizationOptions { TargetDpi = 72, KeepOriginalWhenNotSmaller = false };
        WordImageOptimizationReport analysis = word.AnalyzeImageOptimization(policy);
        Assert.False(analysis.Applied);
        Assert.Equal(4, analysis.ImageCount);
        Assert.Equal(4, analysis.OptimizedCount);
        Assert.Equal(800, OfficeImageReader.Identify(word.Images[0].ToBytes()).Width);
        var applied = word.OptimizeImages(policy);
        Assert.True(applied.Applied);
        Assert.Equal(analysis.BytesSaved, applied.BytesSaved);
        Assert.Equal(before, word._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        using var saved = new MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(4, reopened.AnalyzeImageOptimization(policy).ImageCount);
        Assert.All(applied.Images, image => Assert.Equal(WordImageOptimizationStatus.Optimized, image.Status));
    }

    [Fact]
    public void SharedImageUsesLargestPlacementAndCropDemand() {
        using WordDocument word = WordDocument.Create();
        var first = Insert(word.AddParagraph(), 96, 48);
        first.CropLeft = 25000;
        first.CropRight = 25000;
        var second = Insert(word.AddParagraph(), 192, 96);
        MainDocumentPart part = word._wordprocessingDocument.MainDocumentPart!;
        string obsoleteId = second.RelationshipId!;
        var secondBlip = second._Image.Descendants<A.Blip>().Single();
        secondBlip.Embed = first.RelationshipId;
        part.DeletePart(obsoleteId);
        string before = part.Document.OuterXml;
        var report = word.OptimizeImages(new WordImageOptimizationOptions {
            TargetDpi = 96, KeepOriginalWhenNotSmaller = false
        });
        var item = Assert.Single(report.Images);
        Assert.Equal(2, item.ReferenceCount);
        Assert.Equal(192, item.Final.Width);
        Assert.Equal(96, item.Final.Height);
        Assert.Equal(before, part.Document.OuterXml);
        Assert.Equal(192, OfficeImageReader.Identify(first.ToBytes()).Width);
    }

    [Fact]
    public void StretchedPlacementRetainsRequiredResolutionOnBothAxes() {
        using WordDocument word = WordDocument.Create();
        Insert(word.AddParagraph(), 96, 96);
        var item = Assert.Single(word.OptimizeImages(new WordImageOptimizationOptions {
            TargetDpi = 96, KeepOriginalWhenNotSmaller = false
        }).Images);
        Assert.Equal(192, item.Final.Width);
        Assert.Equal(96, item.Final.Height);
    }

    [Fact]
    public void CandidateFailureAndCancellationLeaveAllMediaUnchanged() {
        using WordDocument word = WordDocument.Create();
        var first = Insert(word.AddParagraph(), 96, 48);
        Insert(word.AddParagraph(), 96, 48);
        byte[] before = first.ToBytes();
        Assert.Throws<InvalidDataException>(() => word.OptimizeImages(new WordImageOptimizationOptions {
            MaxStagedBytes = before.Length + 1, KeepOriginalWhenNotSmaller = false
        }));
        Assert.Equal(before, first.ToBytes());
        Assert.Throws<OperationCanceledException>(() => word.OptimizeImages(cancellationToken: new CancellationToken(true)));
        Assert.Equal(before, first.ToBytes());
    }

    [Fact]
    public void SupportedLegacyInlinePictureCanBeOptimizedAndWrittenBack() {
        using WordDocument word = WordDocument.Create();
        Insert(word.AddParagraph("Picture"), 96, 48);
        using var source = new MemoryStream(word.ToBytes(WordFileFormat.Doc));
        using WordDocument legacy = WordDocument.Load(source);
        var report = legacy.OptimizeImages(new WordImageOptimizationOptions {
            Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 35
        });
        Assert.Equal(1, report.OptimizedCount);
        using var saved = new MemoryStream(legacy.ToBytes(WordFileFormat.Doc));
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(800, OfficeImageReader.Identify(reopened.Images[0].ToBytes()).Width);
        Assert.True(reopened.Images[0].ToBytes().Length < word.Images[0].ToBytes().Length);
    }

    [Fact]
    public void LegacyInlineCropSurvivesOptimizationAndNativeSave() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph("Cropped picture"), 96, 48);
        image.CropLeft = 25000; image.CropRight = 25000; image.CropTop = 12500;
        using var source = new MemoryStream(word.ToBytes(WordFileFormat.Doc));
        using WordDocument imported = WordDocument.Load(source);
        Assert.Equal(25000, imported.Images[0].CropLeft);
        Assert.Equal(12500, imported.Images[0].CropTop);
        var report = imported.OptimizeImages(new() { TargetDpi = 96, KeepOriginalWhenNotSmaller = false });
        Assert.Equal(192, Assert.Single(report.Images).Final.Width);
        using var saved = new MemoryStream(imported.ToBytes(WordFileFormat.Doc));
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(25000, reopened.Images[0].CropLeft);
        Assert.Equal(25000, reopened.Images[0].CropRight);
        Assert.Equal(12500, reopened.Images[0].CropTop);
        Assert.Equal(192, OfficeImageReader.Identify(reopened.Images[0].ToBytes()).Width);
    }

    [Fact]
    public void WordPdfExportUsesSameSizeRecompressionPolicy() {
        using WordDocument word = WordDocument.Create();
        var image = Insert(word.AddParagraph("Text stays text"), 192, 96);
        byte[] source = image.ToBytes();
        byte[] pdf = word.ToPdfBytes(new WordToPdfOptions {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                ImageOptimization = new OfficeIMO.Pdf.PdfImageOptimizationOptions {
                    Enabled = true, Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 35
                }
            }
        });
        var exported = Assert.Single(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal(800, exported.Width);
        Assert.True(exported.Bytes.Length < source.Length);
        Assert.Equal(source, image.ToBytes());
    }

    private static WordImage Insert(WordParagraph paragraph, double width, double height) {
        var raster = new OfficeRasterImage(800, 400);
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                raster.SetPixel(x, y, OfficeColor.FromRgb((byte)(x * 17), (byte)(y * 31), (byte)(x * y)));
        byte[] bytes = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 98, Subsampling = OfficeJpegSubsampling.Y444
        });
        using var stream = new MemoryStream(bytes);
        return paragraph.InsertImage(stream, "picture.jpg", width, height);
    }

    private static byte[] CreateBmp() {
        const int width = 8, height = 4, rowBytes = 24;
        byte[] bytes = new byte[54 + rowBytes * height];
        using var output = new MemoryStream(bytes);
        using var writer = new BinaryWriter(output);
        writer.Write((byte)'B'); writer.Write((byte)'M'); writer.Write(bytes.Length); writer.Write(0); writer.Write(54);
        writer.Write(40); writer.Write(width); writer.Write(height); writer.Write((short)1); writer.Write((short)24);
        writer.Write(0); writer.Write(rowBytes * height); writer.Write(0); writer.Write(0); writer.Write(0); writer.Write(0);
        for (int i = 54; i < bytes.Length; i++) bytes[i] = (byte)(i * 7);
        return bytes;
    }
}
