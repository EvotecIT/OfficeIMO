using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> InvalidSupportedFunctionArities() {
        (int Index, int Minimum, int Maximum)[] functions = {
            (1, 1, 1), (7, 1, 255), (15, 1, 255), (22, 0, 1),
            (30, 1, 255), (31, 1, 255), (32, 1, 1), (33, 2, 2),
            (39, 3, 3), (41, 1, 1), (52, 0, 0), (53, 2, 3),
            (60, 1, 1), (61, 1, 2), (62, 2, 3), (63, 2, 4),
            (76, 1, 2), (77, 1, 1), (84, 1, 255), (86, 1, 255),
            (87, 3, 3), (88, 1, 255), (89, 1, 1), (97, 0, 0),
            (101, 1, 255), (102, 0, 0), (112, 2, 2), (119, 1, 1),
            (124, 1, 2), (168, 1, 255), (169, 2, 3)
        };
        foreach ((int index, int minimum, int maximum) in functions) {
            yield return new object[] { index, minimum == 0 ? maximum + 1 : minimum - 1 };
        }
    }

    public static IEnumerable<object[]> ValidSupportedFunctionArities() {
        yield return new object[] { 1, 1 };
        yield return new object[] { 22, 0 };
        yield return new object[] { 22, 1 };
        yield return new object[] { 62, 3 };
        yield return new object[] { 168, 2 };
    }

    [Theory]
    [MemberData(nameof(InvalidSupportedFunctionArities))]
    public void Supported_formula_functions_reject_invalid_arity(int functionIndex,
        int argumentCount) {
        IWorkFormulaResult result = RenderFunction(functionIndex, argumentCount);

        Assert.False(result.IsComplete);
    }

    [Theory]
    [MemberData(nameof(ValidSupportedFunctionArities))]
    public void Supported_formula_functions_accept_valid_arity(int functionIndex,
        int argumentCount) {
        IWorkFormulaResult result = RenderFunction(functionIndex, argumentCount);

        Assert.True(result.IsComplete);
    }

    [Fact]
    public void Png_preview_budget_accounts_for_expanded_raster_bytes() {
        byte[] png = CreateSizedPreviewPng(8, 8, bitDepth: 1);

        (int? width, int? height) = IWorkImageInfo.Read(png, "image/png",
            maximumDecodedBytes: 100, out long decodedBytes);

        Assert.Null(width);
        Assert.Null(height);
        Assert.Equal(256, decodedBytes);
    }

    [Fact]
    public void Png_previews_reject_forbidden_transparency_chunks() {
        byte[] valid = ValidPreviewPng();
        byte[] malformed = Message(valid[..33],
            CreatePngChunk("tRNS", new byte[] { 0, 0 }), valid[33..]);

        (int? width, int? height) = IWorkImageInfo.Read(malformed, "image/png",
            maximumDecodedBytes: 1024);

        Assert.Null(width);
        Assert.Null(height);
    }

    [Fact]
    public void Nearly_opaque_text_remains_non_opaque_for_owner_preflight() {
        using MemoryStream package = CreatePagesPackageWithColor(0f, alpha: 0.999f,
            includePreview: true);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        IWorkTextRun run = Assert.Single(Assert.Single(result.Projection.Body.Paragraphs).Runs);
        Assert.Equal((byte)254, run.Style.Color!.Alpha);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Editable_text_geometry_requires_both_size_components(bool includeWidth,
        bool includeHeight) {
        byte[] size = Message(
            includeWidth ? FloatField(1, 100f) : Array.Empty<byte>(),
            includeHeight ? FloatField(2, 50f) : Array.Empty<byte>());
        IWorkWireMessage drawable = IWorkProtobuf.Parse(
            Message(BytesField(1, Message(BytesField(2, size)))), new IWorkReadOptions());

        IWorkGeometry? geometry = IWorkDrawingReader.ReadGeometry(drawable,
            out bool complete, requirePositiveSize: true);

        Assert.Null(geometry);
        Assert.False(complete);
    }

    private static IWorkFormulaResult RenderFunction(int functionIndex, int argumentCount) {
        var nodes = new List<byte[]>(argumentCount + 1);
        for (int index = 0; index < argumentCount; index++) {
            nodes.Add(BytesField(1, Message(VarintField(1, 17), DoubleField(4, index + 1d))));
        }
        nodes.Add(BytesField(1, Message(VarintField(1, 16),
            VarintField(2, checked((ulong)functionIndex)),
            VarintField(3, checked((ulong)argumentCount)))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, Message(nodes.ToArray()))), new IWorkReadOptions());
        return IWorkFormulaReader.Render(formula, 0, 0, 512, 4096);
    }
}
