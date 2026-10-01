using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkCellCommentCorpusTests {
    [Fact]
    public void Native_root_comments_preserve_selected_identity_text_author_timestamp_and_saved_xlsx_anchor() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "cell-comments");
        string path = Path.Combine(root, "native-roots.numbers");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "native-roots.json")));
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback, string.Join("\n", result.Report.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult read = reader.ReadDocument(path);
        string json = read.ToJson();
        OfficeDocumentReadResult transported = OfficeDocumentReadResultJson.Deserialize(json);
        Assert.Equal(json, transported.ToJson());
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string sourceSheet = expected.GetProperty("sheet").GetString()!, sourceTable = expected.GetProperty("table").GetString()!;
            IWorkTable table = result.Projection.Sheets.Single(s => s.Name == sourceSheet).Tables.Single(t => t.Name == sourceTable);
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            IWorkTableCell cell = table.GetCell(row, column)!;
            IWorkCellComment comment = Assert.IsType<IWorkCellComment>(cell.Comment);
            Assert.Equal(expected.GetProperty("text").GetString(), comment.Text);
            Assert.Equal(expected.GetProperty("author").GetString(), comment.Author);
            Assert.Equal(expected.GetProperty("commentId").GetUInt64(), comment.SourceIdentity.RecordIdentifier);
            Assert.Equal(expected.GetProperty("authorId").GetUInt64(), comment.SourceAuthorIdentity.RecordIdentifier);
            DateTime expectedDate = DateTime.Parse(expected.GetProperty("creationDateUtc").GetString()!,
                System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal);
            Assert.InRange(Math.Abs((comment.CreationDateUtc - expectedDate).Ticks), 0, 10);
            Assert.Equal(DateTimeKind.Utc, comment.CreationDateUtc.Kind);
            Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures & IWorkCellUnsupportedFeatures.Comment);
            string destinationName = result.WorksheetMappings.Single(m => m.SourceSheetName == sourceSheet && m.SourceTableName == sourceTable).DestinationName;
            var destination = Assert.Single(reopened.Sheets.Single(s => s.Name == destinationName).GetThreadedComments(),
                c => c.CellReference == A1.CellReference(row, column));
            Assert.Equal(comment.Text, destination.Text); Assert.Equal(comment.Author, destination.Author);
            Assert.Null(destination.ParentId);
            Assert.Equal(comment.CreationDateUtc, destination.Date);

            var metadata = Assert.Single(transported.Metadata, m => m.Category == "table.comment"
                && m.Location!.Sheet == sourceSheet
                && transported.Pages.Single(p => p.Location.Sheet == sourceSheet).Tables[m.Location.TableIndex!.Value].Title == sourceTable
                && m.Location.A1Range == A1.CellReference(row, column));
            Assert.Equal(comment.Text, metadata.Value);
            Assert.Equal(comment.Author, metadata.Attributes["author"]);
            Assert.Equal(comment.CreationDateUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
                metadata.Attributes["creationDateUtc"]);
            Assert.Equal(comment.SourceIdentity.EntryPath, metadata.Attributes["sourceEntryPath"]);
            Assert.Equal(comment.SourceIdentity.RecordIdentifier.ToString(), metadata.SourceObjectId);
            var block = Assert.Single(transported.Blocks, b => b.Id == metadata.Location!.BlockAnchor);
            Assert.Equal("comment", block.Kind);
            Assert.Equal(comment.Text, block.Text);
            var page = Assert.Single(transported.Pages, p => p.Location.Sheet == sourceSheet);
            Assert.Equal(sourceTable, page.Tables[metadata.Location!.TableIndex!.Value].Title);
            Assert.Contains(page.Blocks, b => b.Id == block.Id);
            Assert.Equal(comment.Text, string.Concat(transported.Chunks
                .Where(c => c.Location.BlockAnchor == block.Id).Select(c => c.Text)));
        }
        Assert.Equal(3, reopened.Sheets.Sum(s => s.GetThreadedComments().Count));
        Assert.Equal(3, transported.Blocks.Count(b => b.Kind == "comment"));
        saved.Position = 0;
        using var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(saved, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
    }
}
