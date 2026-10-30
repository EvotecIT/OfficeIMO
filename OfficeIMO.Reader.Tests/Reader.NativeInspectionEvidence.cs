using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderNativeInspectionEvidenceTests {
    [Fact]
    public void WordReviewPreservesRepliesRevisionTextAndNativePartThroughJsonAndStructuredExtraction() {
        byte[] bytes = WordReviewBytes();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "review.docx", Options());
        Assert.Equal(2, result.Metadata.Count(item => item.Category == OfficeDocumentInspectionCategories.Comment));
        OfficeDocumentMetadataEntry revision = Assert.Single(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Revision);
        Assert.Equal("Tracked source", revision.Value);
        Assert.Equal("/word/document.xml", revision.Attributes["partUri"]);
        Assert.Equal("false", revision.Attributes["acceptedOrRejected"]);
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Attributes["isReply"] == "true" && item.Value == "Reply evidence");
        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(result.ToJson());
        Assert.Contains(restored.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Revision && item.Value == revision.Value && item.Attributes["partUri"] == revision.Attributes["partUri"]);
        Assert.Contains(reader.ExtractStructured(restored).Records, item => item.Value == "Tracked source" && item.Location?.SourceBlockKind == "revision");
        // Inspection must leave authored review markup intact for a subsequent read.
        using var input = new MemoryStream(bytes, writable: false);
        using WordDocument original = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        Assert.Single(original.InspectReview().Revisions);
    }

    [Fact]
    public void ExcelSeparatesExpressionsSavedCachesAndMissingCachesAndRetainsEmptyCellRepliesAndVisibility() {
        using var stream = new MemoryStream();
        using (ExcelDocument workbook = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = workbook.AddSheet("Hidden data");
            sheet.Cell(1, 1, 1); sheet.Cell(2, 1, 99);
            sheet.SetComment("A2", "Review cache", "Ada");
            var parent = sheet.AddThreadedComment("C4", "Blank-cell parent", "Grace");
            sheet.AddThreadedComment("C4", "Blank-cell reply", "Lin", parent.Id);
            var part = workbook.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single();
            Cell cached = part.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A2");
            cached.CellFormula = new CellFormula("A1+1");
            part.Worksheet!.GetFirstChild<SheetData>()!.Append(new Row(new Cell { CellReference = "A3", CellFormula = new CellFormula("A1+2") }) { RowIndex = 3U, Hidden = true });
            part.Worksheet!.InsertBefore(new Columns(new Column { Min = 2U, Max = 2U, Hidden = true }), part.Worksheet.GetFirstChild<SheetData>());
            workbook.OpenXmlDocument!.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single().State = SheetStateValues.VeryHidden;
            workbook.Save();
        }
        ReaderOptions options = Options(); options.Inspection!.IncludeFormulas = true; options.Inspection.IncludeVisibility = true;
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddExcelHandler().Build().ReadDocument(stream.ToArray(), "cache.xlsx", options);
        var formulas = result.Metadata.Where(item => item.Category == OfficeDocumentInspectionCategories.Formula).ToArray();
        Assert.Equal(2, formulas.Length);
        Assert.Contains(formulas, item => item.Location!.A1Range == "A2" && item.Value == "A1+1" && item.Attributes["storedValue"] == "99" && item.Attributes["calculated"] == "false");
        Assert.Contains(formulas, item => item.Location!.A1Range == "A3" && item.Attributes["hasStoredValue"] == "false");
        Assert.Equal(2, result.Metadata.Count(item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Location!.A1Range == "C4"));
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Worksheet" && item.Value!.Contains("veryHidden", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Row");
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Column range");
    }

    [Fact]
    public void PowerPointRetainsReviewAndHiddenSourceStateEvenWhenHiddenShapesAreExcludedFromText() {
        using var stream = new MemoryStream();
        byte[] bytes;
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide(); slide.Hidden = true;
            slide.AddTextBox("Hidden author text").Hidden = true;
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            presentation.AddClassicComment(slide, author, "Review evidence",
                x: OfficeLength.Points(1 / 8D), y: OfficeLength.Points(-25 / 8D));
            presentation.AddModernComment(slide, author, "Modern evidence",
                x: OfficeLength.Inches(1), y: OfficeLength.Inches(-1));
            bytes = presentation.ToBytes();
        }
        ReaderOptions options = Options(); options.Inspection!.IncludeVisibility = true;
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddPowerPointHandler(new ReaderPowerPointOptions { IncludeHiddenShapes = false }).Build().ReadDocument(bytes, "review.pptx", options);
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Value == "Review evidence" && item.Location!.Slide == 1);
        OfficeDocumentMetadataEntry classic = Assert.Single(result.Metadata,
            item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Value == "Review evidence");
        Assert.Equal("1", classic.Attributes["x"]);
        Assert.Equal("-25", classic.Attributes["y"]);
        OfficeDocumentMetadataEntry modern = Assert.Single(result.Metadata,
            item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Value == "Modern evidence");
        Assert.False(modern.Attributes.ContainsKey("x"));
        Assert.False(modern.Attributes.ContainsKey("y"));
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Slide");
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Shape");
    }

    [Fact]
    public void OpenDocumentRichResultRetainsMetadataLinksAndImageBytesWithoutFetchingLinkTargets() {
        OdtDocument document = OdtDocument.Create(); document.Metadata.Title = "Source title";
        OdtParagraph paragraph = document.AddParagraph("Body");
        paragraph.AddHyperlink("Target", "https://example.invalid/never-fetched");
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        paragraph.AddImage(png, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build().ReadDocument(document.ToBytes(), "rich.odt", Options());
        Assert.Equal("Source title", result.Source.Title);
        Assert.Equal("https://example.invalid/never-fetched", Assert.Single(result.Links).Uri);
        Assert.Equal(png, Assert.Single(result.Assets).PayloadBytes);
        Assert.NotNull(result.Assets[0].PayloadHash);
        Assert.Empty(result.Pages); // Flow content must not acquire inferred physical page numbers.
    }

    [Fact]
    public void RtfConcealmentAndUnicodeEvidenceUseTheNativeOwnerAndRemainExplicitlyHeuristic() {
        byte[] source = Encoding.ASCII.GetBytes(@"{\rtf1\ansi Visible {\v hidden instruction} text}");
        ReaderOptions options = Options(); options.Inspection!.IncludeContentSafety = true;
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddRtfHandler().Build().ReadDocument(source, "hidden.rtf", options);
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.ContentSafety && item.Value!.Contains("hidden instruction") && item.Attributes.ContainsKey("contentSha256"));
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Coverage && item.Name == "content-safety" && item.Value == "bounded-heuristics");
    }

    [Fact]
    public void NestedDocumentsRetainInspectionAndShareTerminalRecordLimits() {
        byte[] source = WordReviewBytes();
        using var zip = new MemoryStream();
        using (var archive = new ZipArchive(zip, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (string name in new[] { "first.docx", "second.docx" }) {
                using Stream output = archive.CreateEntry(name).Open(); output.Write(source, 0, source.Length);
            }
        }
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddZipHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(zip.ToArray(), "documents.zip", Options());
        Assert.Equal(2, result.NestedDocuments.Count);
        Assert.All(result.NestedDocuments, child => Assert.Contains(child.Document.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Comment));
        ReaderOptions bounded = Options(); bounded.Inspection!.MaxRecords = 7;
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(zip.ToArray(), "documents.zip", bounded));
        Assert.Equal("Inspection.MaxRecords", error.LimitName);
    }

    [Fact]
    public void InspectionCharacterLimitsAndCancellationAreTerminal() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
        ReaderOptions options = Options(); options.Inspection!.MaxCharacters = 40;
        Assert.Equal("Inspection.MaxCharacters", Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(WordReviewBytes(), "review.docx", options)).LimitName);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => reader.ReadDocument(WordReviewBytes(), "review.docx", Options(), cancellation.Token));
    }

    [Fact]
    public void PackageEvidenceUsesTheExactSourceAndNeverValidatesExternalTargetsOrSignatures() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Source package");
            document.OpenXmlDocument!.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.invalid/passive"), true);
            document.Save();
        }
        byte[] bytes = stream.ToArray();
        using var input = new MemoryStream(bytes, writable: false);
        using WordDocument owner = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true });
        Assert.NotNull(owner.SourcePackageSecurityReport);
        Assert.Equal(bytes.LongLength, owner.SourcePackageSecurityReport!.PackageBytes);
        Assert.Equal(1, owner.SourcePackageSecurityReport.ExternalRelationshipCount);
        ReaderOptions options = Options(); options.Inspection!.IncludePackageContent = true;
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddWordHandler().Build().ReadDocument(bytes, "source.docx", options);
        OfficeDocumentMetadataEntry package = Assert.Single(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.PackageSecurity && item.Name == "Package structure");
        Assert.Equal(bytes.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture), package.Attributes["packageBytes"]);
        Assert.Equal("source-container", package.Attributes["scope"]);
        Assert.Equal("1", package.Attributes["externalRelationshipCount"]); Assert.Equal("false", package.Attributes["executed"]);
        Assert.Contains(OfficeDocumentReadResultJson.Deserialize(result.ToJson()).Metadata, item => item.Id == package.Id && item.Attributes["externalRelationshipCount"] == "1");
    }

    [Fact]
    public void ChildInspectionLimitsRemainTerminalInsideAResourceOnlyParentOperation() {
        OfficeDocumentReader child = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
        ReaderOptions childOptions = Options(); childOptions.Inspection!.MaxRecords = 1;
        OfficeDocumentReader parent = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddProcessor(new DelegateOfficeDocumentProcessor("native-child", (document, _) => {
                child.ReadDocument(WordReviewBytes(), "child.docx", childOptions); return document;
            })).Build();
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => parent.ReadDocument(
            Encoding.UTF8.GetBytes("parent"), "parent.txt", new ReaderOptions { ResourceLimits = new ReaderResourceLimits { MaxChunks = 100 } }));
        Assert.Equal("Inspection.MaxRecords", error.LimitName);
    }

    private static ReaderOptions Options() => new ReaderOptions { Inspection = new ReaderInspectionOptions {
        IncludeContentSafety = false, IncludePackageContent = false, IncludeFormulas = false, IncludeVisibility = false
    }};

    private static byte[] WordReviewBytes() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Target").AddComment("Reviewer", "R", "Parent evidence");
            document.Comments.Single().AddReply("Reply author", "RA", "Reply evidence");
            document.OpenXmlDocument!.MainDocumentPart!.Document!.Body!.Append(new W.Paragraph(new W.InsertedRun(new W.Run(new W.Text("Tracked source"))) {
                Id = "7", Author = "Editor", Date = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc)
            }));
            document.Save();
        }
        return stream.ToArray();
    }
}
