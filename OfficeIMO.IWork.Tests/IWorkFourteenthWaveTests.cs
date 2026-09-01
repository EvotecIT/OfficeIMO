using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_duration_functions_are_not_claimed_as_excel_compatible() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 1d))),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 212), VarintField(3, 1))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.False(result.IsComplete);
        Assert.Equal("=DURATION(1)", result.Text);
    }

    [Fact]
    public void Formula_renderer_parenthesizes_unary_exponent_bases() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 2d))),
            BytesField(1, Message(VarintField(1, 13))),
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 2d))),
            BytesField(1, Message(VarintField(1, 5))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.True(result.IsComplete);
        Assert.Equal("=(-2)^2", result.Text);
    }

    [Theory]
    [InlineData("\uFFFC")]
    [InlineData("\uFFFB")]
    public void Numbers_text_storage_reports_removed_inline_object_markers(string marker) {
        var options = new IWorkReadOptions();
        IWorkWireMessage storage = IWorkProtobuf.Parse(
            Message(StringField(3, "Before" + marker + "After")), options);

        string text = IWorkPagesReader.StorageText(storage,
            new IWorkProjectionBudget(options), out bool complete);

        Assert.False(complete);
        Assert.Equal("BeforeAfter", text);
    }

    [Fact]
    public void Reused_pages_headers_are_charged_by_destination_text_complexity() {
        using MemoryStream package = CreatePagesPackageWithRepeatedHeader(sectionCount: 2);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 5 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("text item count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Wrong_wire_repeated_object_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Table", 1, 1, 42d)
        }, includePreview: true, includeWrongWireDrawableReference: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_default_row_height_uses_the_constant_space_worksheet_default() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Tall table", 4097, 1, 42d, defaultRowHeight: 20d)
        });

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.Equal(20d, result.Document.Sheets[0].DefaultRowHeight);
    }

    private static MemoryStream CreatePagesPackageWithRepeatedHeader(int sectionCount) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong headerFooterId = 3;
        const ulong headerStorageId = 4;
        ulong[] sectionIds = Enumerable.Range(0, sectionCount)
            .Select(index => checked((ulong)(10 + index))).ToArray();
        byte[] sectionTable = Message(sectionIds.Select(sectionId =>
            BytesField(1, Message(ReferenceField(2, sectionId)))).ToArray());
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, string.Join("\u0004",
                        Enumerable.Repeat("Body", sectionCount))),
                    BytesField(17, sectionTable)), sectionIds),
            ArchiveRecord(headerFooterId, 10143, Message(ReferenceField(1, headerStorageId)),
                new[] { headerStorageId }),
            ArchiveRecord(headerStorageId, 2001, Message(StringField(3, "Header")))
        };
        records.AddRange(sectionIds.Select(sectionId => ArchiveRecord(sectionId, 10011,
            Message(ReferenceField(25, headerFooterId)), new[] { headerFooterId })));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
    }
}
