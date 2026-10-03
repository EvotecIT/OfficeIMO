using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

Require(!System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported
    && !System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeCompiled,
    "The iWork qualification host must execute as NativeAOT, not a managed apphost.");

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
var fixtures = new[] {
    ("simple.pages", IWorkDocumentKind.Pages, "5AEE6D03277D2DB2104F593E64AFE081DEC539F0117B97124B6F99158124C93E"),
    ("simple.numbers", IWorkDocumentKind.Numbers, "D0B00D9CAE5985CCCAA3B2FB251FAE92EB0E38360FB4B5DF8B4350EB658F752B"),
    ("simple.key", IWorkDocumentKind.Keynote, "BA95755DF82CEB0CA834E1E03E2777C34FAD906320D8336B4F3FEFC6B48607EB"),
    ("formulas.numbers", IWorkDocumentKind.Numbers, "DD85BAD68898CE5B065F277C0B9BE1F3C32D696E3BAA6B09D3614BBD35A5249F"),
    ("comments.numbers", IWorkDocumentKind.Numbers, "81814EEC7D90108595F3A6E41457B980C1FD935E16A74D77BAD20AB11006DCAF")
};
foreach (var (name, kind, sha256) in fixtures) {
    string path = Path.Combine(AppContext.BaseDirectory, name);
    byte[] bytes = File.ReadAllBytes(path);
    Require(Convert.ToHexString(SHA256.HashData(bytes)) == sha256, name + " fixture bytes changed.");
    IWorkSourceDocument fromPath = IWorkSourceDocument.Open(path);
    using var sourceStream = new MemoryStream(bytes, writable: false);
    IWorkSourceDocument fromStream = IWorkSourceDocument.Open(sourceStream, kind);
    Require(sourceStream.CanRead && fromPath.Kind == kind && fromStream.Kind == kind,
        name + " source kind or caller-owned stream contract failed.");
    Require(fromPath.Records.Count > 0 && fromPath.Records.Count == fromStream.Records.Count
        && fromPath.BuildVersions.Count > 0, name + " lost source records or producer metadata.");
    Require(ProjectionSignature(fromPath) == ProjectionSignature(fromStream), name + " source path/stream mismatch.");
    VerifySource(fromPath, name);

    OfficeDocumentReadResult result = reader.ReadDocument(path);
    using var readerStream = new MemoryStream(bytes, writable: false);
    OfficeDocumentReadResult streamed = reader.ReadDocument(readerStream, name);
    Require(readerStream.CanRead && result.Kind == ReaderInputKind.IWork && result.Chunks.Count > 0,
        name + " Reader extraction or stream ownership failed.");
    Require(result.Markdown == streamed.Markdown && result.Tables.Count == streamed.Tables.Count
        && result.Pages.Count == streamed.Pages.Count, name + " Reader path/stream mismatch.");
    if (kind == IWorkDocumentKind.Pages) {
        Require(result.Chunks.Any(chunk => chunk.Text.Contains("hello pages", StringComparison.Ordinal)),
            "Reader lost Pages body text.");
    } else if (kind == IWorkDocumentKind.Keynote) {
        Require(result.Pages.Count == 2 && result.Chunks.Any(chunk => chunk.Text.Contains("note text here", StringComparison.Ordinal)),
            "Reader lost Keynote slides or presenter notes.");
    } else {
        Require(result.Tables.Count > 0, "Reader lost Numbers tables.");
    }
    string json = OfficeDocumentReadResultJson.Serialize(result);
    OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(json);
    Require(OfficeDocumentReadResultJson.Serialize(roundTrip) == json,
        name + " Reader transport changed during JSON round-trip.");
    if (name == "comments.numbers") VerifyComments(roundTrip);

    Reject<OperationCanceledException>(() => IWorkSourceDocument.Open(path, options: null,
        new CancellationToken(canceled: true)), name + " pre-cancellation");
    Reject<InvalidDataException>(() => IWorkSourceDocument.Open(path,
        new IWorkReadOptions { MaximumPackageBytes = 1 }), name + " package byte limit");
    Reject<OperationCanceledException>(() => reader.ReadDocument(path,
        cancellationToken: new CancellationToken(canceled: true)), name + " Reader pre-cancellation");
    Reject<IOException>(() => reader.ReadDocument(path,
        new ReaderOptions { MaxInputBytes = 1 }), name + " Reader input byte limit");
    Console.WriteLine($"PASS | {name} | {sha256} | source, Reader, JSON, cancellation, limits");
}
Console.WriteLine("PASS | bounded iWork NativeAOT source and Reader scenarios | "
    + System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);

static string ProjectionSignature(IWorkSourceDocument source) => source.Kind switch {
    IWorkDocumentKind.Pages => string.Join("\n", source.ReadPages().Paragraphs),
    IWorkDocumentKind.Keynote => string.Join("\n", source.ReadKeynote().Slides.Select(slide =>
        slide.Title + "\n" + string.Join("\n", slide.Body) + "\n" + slide.PresenterNotes)),
    IWorkDocumentKind.Numbers => string.Join("\n", source.ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables)
        .SelectMany(table => table.Cells).Select(cell => string.Join("|", cell.Row, cell.Column,
            cell.Kind, Convert.ToString(cell.Value, CultureInfo.InvariantCulture), cell.Formula, cell.FormulaIsComplete))),
    _ => throw new InvalidOperationException("Unexpected source kind.")
};

static void VerifySource(IWorkSourceDocument source, string name) {
    if (source.Kind == IWorkDocumentKind.Pages) {
        IWorkPagesProjection pages = source.ReadPages();
        Require(pages.Paragraphs[0] == "hello pages"
            && pages.Paragraphs.Any(text => text.Contains("second paragraph with some words", StringComparison.Ordinal)),
            "Pages source text changed.");
        Require(pages.HasRecoverableContent && !pages.HasEditableContent,
            "Pages partial-reconstruction classification changed.");
        Reject<InvalidOperationException>(() => pages.CreateConversionReport(IWorkProjectionKind.EditableReconstruction),
            "Pages complete-reconstruction report for partial source");
    } else if (source.Kind == IWorkDocumentKind.Keynote) {
        IWorkKeynoteProjection keynote = source.ReadKeynote();
        Require(keynote.Slides.Count == 2 && keynote.Slides[0].Title == "hello keynote"
            && keynote.Slides[0].Body.Any(text => text.Contains("first bullet", StringComparison.Ordinal))
            && keynote.Slides[1].Title == "second slide"
            && keynote.Slides[1].PresenterNotes.Contains("note text here", StringComparison.Ordinal),
            "Keynote source slides, text or notes changed.");
    } else {
        IWorkNumbersProjection numbers = source.ReadNumbers();
        IWorkTable table = numbers.Sheets[0].Tables[0];
        if (name == "simple.numbers") {
            Require(numbers.Sheets.Count == 1 && numbers.Sheets[0].Tables.Count == 1
                && table.RowCount == 3 && table.ColumnCount == 3
                && table.GetCell(1, 1)!.Value is string first && first == "a"
                && table.GetCell(2, 2)!.Value is double value && value == 2
                && table.GetCell(3, 3)!.Value is string last && last == "Z", "Numbers sparse typed values changed.");
        } else if (name == "formulas.numbers") {
            IWorkTableCell arithmetic = table.GetCell(2, 2)!;
            Require(arithmetic.Kind == IWorkCellKind.Formula && arithmetic.FormulaIsComplete
                && arithmetic.Formula == "=A1+A2" && arithmetic.Value is double value && value == 3
                && table.GetCell(6, 2)!.Formula == "=SUM(A1:A2)"
                && numbers.Sheets[1].Tables[0].GetCell(3, 2)!.Formula == "=LEFT(A3,1)",
                "Numbers formula expression or typed cache changed.");
        }
    }
}

static void VerifyComments(OfficeDocumentReadResult result) {
    using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "comments.json")));
    Require(result.Blocks.Count(block => block.Kind == "comment") == 3, "Reader lost native root comments.");
    foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
        string address = OfficeIMO.Spreadsheet.SpreadsheetRangeReference.FromCell(
            column: expected.GetProperty("column").GetInt32(), row: expected.GetProperty("row").GetInt32())
            .Format(OfficeIMO.Spreadsheet.SpreadsheetAddressDialect.UnboundedA1);
        OfficeDocumentMetadataEntry metadata = result.Metadata.Single(entry => entry.Category == "table.comment"
            && entry.Location!.Sheet == expected.GetProperty("sheet").GetString()
            && result.Pages.Single(page => page.Location.Sheet == entry.Location.Sheet)
                .Tables[entry.Location.TableIndex!.Value].Title == expected.GetProperty("table").GetString()
            && entry.Location.A1Range == address);
        Require(metadata.Value == expected.GetProperty("text").GetString()
            && metadata.Attributes["author"] == expected.GetProperty("author").GetString()
            && metadata.SourceObjectId == expected.GetProperty("commentId").GetUInt64().ToString(CultureInfo.InvariantCulture),
            "Reader lost native comment content, author or identity.");
        DateTime actualDate = DateTime.Parse(metadata.Attributes["creationDateUtc"], CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind);
        DateTime expectedDate = DateTime.Parse(expected.GetProperty("creationDateUtc").GetString()!,
            CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal);
        Require(actualDate.Kind == DateTimeKind.Utc && Math.Abs((actualDate - expectedDate).Ticks) <= 10,
            "Reader changed native comment creation time.");
        OfficeDocumentBlock block = result.Blocks.Single(block => block.Id == metadata.Location!.BlockAnchor);
        Require(block.Text == metadata.Value && block.Location.A1Range == address,
            "Reader comment block lost content or cell anchor.");
        Require(string.Concat(result.Chunks.Where(chunk => chunk.Location.BlockAnchor == block.Id).Select(chunk => chunk.Text)) == metadata.Value,
            "Reader comment chunks lost content.");
    }
}

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}
static void Reject<T>(Action action, string contract) where T : Exception {
    try { action(); } catch (T) { return; }
    throw new InvalidOperationException(contract + " did not reject the operation.");
}
