using DocumentFormat.OpenXml.Packaging;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableStreamExportContracts {
    [Fact]
    public void WordTableExport_WritesToNonSeekableDestination() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new NonSeekableWriteStream();

        logical.SaveAsWord(destination, PdfWordReadOptions.CreateTablesOnly());

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.MainDocumentPart);
    }

    [Fact]
    public void PowerPointTableExport_WritesToNonSeekableDestination() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new NonSeekableWriteStream();

        logical.SaveTablesAsPowerPoint(destination);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(destination.ToArray()), false);
        Assert.NotNull(package.PresentationPart);
    }

    [Fact]
    public async Task TableConversions_ProvideReportsAndAsyncCallerOwnedStreamWrites() {
        PdfLogicalDocument logical = CreateLogicalDocument();

        PdfWordConversionResult wordResult = logical.ToWordDocumentResult(PdfWordReadOptions.CreateTablesOnly());
        PdfExcelTableImportResult excelResult = logical.ImportTablesToExcelDocumentResult();
        PdfPowerPointTableImportResult powerPointResult = logical.ImportTablesToPowerPointPresentationResult();

        using var wordDocument = wordResult.RequireNoLoss();
        using var excelDocument = excelResult.RequireNoLoss();
        using var powerPointPresentation = powerPointResult.RequireNoLoss();
        Assert.NotEmpty(wordDocument.ToBytes());
        Assert.NotEmpty(excelDocument.ToBytes());
        Assert.NotEmpty(powerPointPresentation.ToBytes());
        Assert.False(wordResult.Report.HasLoss);
        Assert.False(excelResult.Report.HasLoss);
        Assert.False(powerPointResult.Report.HasLoss);
        Assert.True(excelResult.HasOmittedPageContent);
        Assert.True(powerPointResult.HasOmittedPageContent);
        Assert.Equal(1, excelResult.Report.SourceScope.NonTableTextBlockCount);
        Assert.Equal(1, powerPointResult.Report.SourceScope.NonTableTextBlockCount);
        Assert.Equal(0, excelResult.Report.SourceScope.DetectedTableCount);
        Assert.Equal(0, powerPointResult.Report.SourceScope.DetectedTableCount);

        using var wordStream = new MemoryStream();
        using var excelStream = new MemoryStream();
        using var powerPointStream = new MemoryStream();
        await logical.SaveAsWordAsync(wordStream, PdfWordReadOptions.CreateTablesOnly());
        await logical.SaveTablesAsExcelAsync(excelStream);
        await logical.SaveTablesAsPowerPointAsync(powerPointStream);

        wordStream.WriteByte(0);
        excelStream.WriteByte(0);
        powerPointStream.WriteByte(0);
        Assert.True(wordStream.Length > 1);
        Assert.True(excelStream.Length > 1);
        Assert.True(powerPointStream.Length > 1);
    }

    [Fact]
    public async Task TableConversionAsyncWrites_HonorPreCanceledTokens() {
        PdfLogicalDocument logical = CreateLogicalDocument();
        using var destination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            logical.SaveAsWordAsync(destination, PdfWordReadOptions.CreateTablesOnly(), cancellation.Token));
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public void TableConversions_ReportOmittedVectorGraphics() {
        byte[] source = PdfDocument.Create()
            .Rectangle(
                120,
                40,
                strokeColor: PdfColor.FromRgb(0, 64, 128),
                strokeWidth: 2,
                fillColor: PdfColor.FromRgb(204, 238, 255))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);

        PdfExcelTableImportResult excelResult = logical.ImportTablesToExcelDocumentResult();
        PdfPowerPointTableImportResult powerPointResult = logical.ImportTablesToPowerPointPresentationResult();

        Assert.Empty(logical.TextBlocks);
        Assert.Empty(logical.Images);
        Assert.True(logical.Pages[0].VectorPrimitiveCount > 0);
        Assert.True(excelResult.HasOmittedPageContent);
        Assert.True(powerPointResult.HasOmittedPageContent);
        Assert.True(excelResult.Report.SourceScope.VectorPrimitiveCount > 0);
        Assert.Equal(
            excelResult.Report.SourceScope.VectorPrimitiveCount,
            powerPointResult.Report.SourceScope.VectorPrimitiveCount);
    }

    [Fact]
    public void TableConversions_IgnoreInvisibleVectorGraphics() {
        byte[][] sources = {
            BuildSingleStreamPdf("1 0 0 rg\n300 300 40 40 re\nf"),
            BuildSingleStreamPdf(
                "/GS1 gs\n1 0 0 rg\n40 40 40 40 re\nf",
                "<< /ExtGState << /GS1 5 0 R >> >>",
                "5 0 obj\n<< /Type /ExtGState /ca 0 /CA 0 >>\nendobj"),
            BuildSingleStreamPdf("0 0 5 5 re W n\n1 0 0 rg\n40 40 40 40 re\nf"),
            BuildSingleStreamPdf("0 0 m\n100 0 l\n0 100 l\nh\nW n\n1 0 0 rg\n80 80 10 10 re\nf"),
            BuildSingleStreamPdf("0 0 m\n100 0 l\n0 100 l\nh\nW n\n1 0 0 RG\n1 w\n20 90 m\n90 20 l\nS"),
            BuildSingleStreamPdf("0 0 100 100 re\n30 30 40 40 re\nW* n\n1 0 0 rg\n40 40 10 10 re\nf"),
            BuildSingleStreamPdf("1 0 0 rg\n240 40 20 20 re\nf"),
            BuildSingleStreamPdf("40 40 20 20 re\n40 40 20 20 re\nf*"),
            BuildSingleStreamPdf("45 45 10 10 re W n\n1 0 0 RG\n1 w\n20 20 m\n20 80 l\nh\n80 20 l\nS"),
            BuildSingleStreamPdf(
                "/Pattern cs\n/P1 scn\n40 40 40 40 re\nf",
                "<< /Pattern << /P1 5 0 R >> >>",
                "5 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >> /Length 0 >>\nstream\n\nendstream\nendobj"),
            BuildSingleStreamPdf(
                "/Pattern cs\n/P1 scn\n40 40 40 40 re\nf",
                "<< /Pattern << /P1 5 0 R >> >>",
                "5 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /ExtGState << /GS1 6 0 R >> >> /Length 31 >>\nstream\n/GS1 gs\n1 0 0 rg\n0 0 10 10 re\nf\nendstream\nendobj",
                "6 0 obj\n<< /Type /ExtGState /ca 0 /CA 0 >>\nendobj"),
            BuildPatternWithClippedForm(hiddenByClip: true),
            BuildSingleStreamPdf("1 0 0 rg\n1e309 40 20 20 re\nf")
        };

        Assert.All(sources, source => {
            PdfLogicalDocument logical = PdfLogicalDocument.Load(source);
            PdfTableExtractionScopeReport scope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);

            Assert.Equal(0, logical.Pages[0].VectorPrimitiveCount);
            Assert.Equal(0, scope.VectorPrimitiveCount);
            Assert.False(scope.HasOmittedPageContent);
        });
    }

    [Fact]
    public void TableConversions_CountStrokesCrossingThePageBoundary() {
        byte[] source = BuildSingleStreamPdf("""
            1 0 0 RG
            4 w
            -2 40 2 20 re
            S
            -2 80 m
            0 90 l
            -2 100 l
            S
            """);
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);
        PdfTableExtractionScopeReport scope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);

        Assert.Equal(2, logical.Pages[0].VectorPrimitiveCount);
        Assert.Equal(2, scope.VectorPrimitiveCount);
        Assert.True(scope.HasOmittedPageContent);
    }

    [Fact]
    public void TableConversions_CountTilingPatternFormsVisibleWithinNestedClips() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(
            BuildPatternWithClippedForm(hiddenByClip: false));
        PdfTableExtractionScopeReport scope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);

        Assert.Equal(1, logical.Pages[0].VectorPrimitiveCount);
        Assert.Equal(1, scope.VectorPrimitiveCount);
        Assert.True(scope.HasOmittedPageContent);
    }

    [Fact]
    public void VectorVisibility_ComplexDisjointPathsStayBoundedAndConservative() {
        const int contourCount = 511;
        var content = new System.Text.StringBuilder(contourCount * 80);
        for (int i = 0; i < contourCount; i++) {
            content.Append("0 0 m\n80 0 l\n0 80 l\nh\n");
        }
        content.Append("W* n\n1 0 0 rg\n");
        for (int i = 0; i < contourCount; i++) {
            content.Append("100 100 m\n20 100 l\n100 20 l\nh\n");
        }
        content.Append('f');

        byte[] source = BuildSingleStreamPdf(content.ToString());
        var timer = System.Diagnostics.Stopwatch.StartNew();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);
        int vectorPrimitiveCount = logical.Pages[0].VectorPrimitiveCount;
        timer.Stop();

        Assert.Equal(1, vectorPrimitiveCount);
        Assert.True(
            timer.Elapsed < TimeSpan.FromSeconds(5),
            "Complex visibility analysis exceeded the bounded contract: " + timer.Elapsed + ".");
    }

    private static PdfLogicalDocument CreateLogicalDocument() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Non-seekable table export proof"))
            .ToBytes();
        return PdfLogicalDocument.Load(source);
    }

    private static byte[] BuildPatternWithClippedForm(bool hiddenByClip) {
        const string patternContent = "/F1 Do";
        string formContent = hiddenByClip
            ? "0 0 2 2 re W n\n1 0 0 rg\n5 5 2 2 re\nf"
            : "0 0 2 2 re W n\n1 0 0 rg\n0.5 0.5 1 1 re\nf";
        return BuildSingleStreamPdf(
            "/Pattern cs\n/P1 scn\n40 40 40 40 re\nf",
            "<< /Pattern << /P1 5 0 R >> >>",
            $"5 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 6 0 R >> >> /Length {System.Text.Encoding.ASCII.GetByteCount(patternContent)} >>\nstream\n{patternContent}\nendstream\nendobj",
            $"6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >> /Length {System.Text.Encoding.ASCII.GetByteCount(formContent)} >>\nstream\n{formContent}\nendstream\nendobj");
    }

    private static byte[] BuildSingleStreamPdf(
        string streamContent,
        string resources = "<< >>",
        params string[] extraObjects) {
        streamContent = streamContent.TrimEnd('\r', '\n');
        int streamLength = System.Text.Encoding.ASCII.GetByteCount(streamContent);
        string[] objects = {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Resources {resources} /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent,
            "endstream",
            "endobj"
        };
        string pdf = string.Join(
            "\n",
            objects.Concat(extraObjects).Concat(new[] {
                "trailer",
                "<< /Root 1 0 R >>",
                "%%EOF"
            })) + "\n";
        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private sealed class NonSeekableWriteStream : Stream {
        private readonly MemoryStream _inner = new MemoryStream();

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public byte[] ToArray() => _inner.ToArray();
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
