using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Package_metadata_selectively_materializes_only_image_entries() {
        using MemoryStream package = CreatePagesImagePackage(
            duplicateMetadata: false, imageCount: 1,
            imageBytes: ValidPreviewPng(), metadataOuterFieldCount: 2);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedImages = 1 });
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);
        var budget = new IWorkProjectionBudget(source.Options);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(
            source, image, budget, out bool complete);

        Assert.NotNull(asset);
        Assert.True(complete);
    }

    [Fact]
    public void Formula_numeric_nodes_require_one_value_field() {
        byte[] node = Message(VarintField(1, 17),
            DoubleField(4, 1d), DoubleField(7, 2d));
        byte[] formulaBytes = Message(BytesField(1,
            Message(BytesField(1, node))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(formulaBytes,
            new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(
            formula, 0, 0, maximumNodes: 32, maximumCharacters: 128);

        Assert.False(result.IsComplete);
    }

    [Fact]
    public void Keynote_rotations_must_round_trip_through_owner_angle_units() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(
            1, rotation: 0.00001f);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Theory]
    [InlineData("/Root 2 0 R ")]
    [InlineData("/Size 4 ")]
    public void Classic_pdf_trailers_require_singular_structural_keys(string duplicateKey) {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            trailerDictionaryPrefix: duplicateKey);

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }
}
