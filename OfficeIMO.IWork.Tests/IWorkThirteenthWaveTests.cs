using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Keynote_slide_names_are_charged_to_the_source_wide_text_budget() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Title", slideName: "Named");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 5 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_sheet_and_table_names_share_the_projection_text_budget() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("T", 1, 1, 42d)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 5 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Projected_text_metadata_shares_the_source_wide_character_budget() {
        using MemoryStream package = CreatePagesPackage(includeBody: false, textBox: "X",
            includePreview: false, textBoxDrawable: Message(StringField(4, "https://example.test")));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inline_breaks_are_charged_before_owner_node_materialization() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
            bodyText: "A\u2028B\u2028C");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 3 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("text item count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Malformed_Keynote_skipped_slide_flags_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            wrongWireSkippedFlag: true);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SKIPPED_SLIDE_UNSUPPORTED");
    }

    [Theory]
    [InlineData(true, "=A:A")]
    [InlineData(false, "=1:1")]
    public void Formula_renderer_emits_valid_whole_axis_references(bool columnOnly,
        string expected) {
        byte[] coordinate = Message(VarintField(1, 0));
        byte[] reference = Message(VarintField(1, 36),
            columnOnly ? BytesField(26, coordinate) : BytesField(27, coordinate));
        IWorkWireMessage formula = IWorkProtobuf.Parse(Message(BytesField(1,
            Message(BytesField(1, reference)))), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.True(result.IsComplete);
        Assert.Equal(expected, result.Text);
    }

    [Fact]
    public void Formula_renderer_rejects_wrong_wire_coordinate_messages() {
        byte[] reference = Message(VarintField(1, 36), VarintField(26, 0));
        IWorkWireMessage formula = IWorkProtobuf.Parse(Message(BytesField(1,
            Message(BytesField(1, reference)))), new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 32, 128);

        Assert.False(result.IsComplete);
        Assert.Equal("=#REF!", result.Text);
    }
}
