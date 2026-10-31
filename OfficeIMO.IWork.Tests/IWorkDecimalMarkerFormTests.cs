using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(1UL, "(1)", "(7)", "(8)", "(3)", "(4)")]
    [InlineData(2UL, "1)", "7)", "8)", "3)", "4)")]
    public void Decimal_marker_forms_preserve_source_counters_and_saved_DOCX_markers(ulong numberType,
        string sample, params string[] markers) {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Pages, numberType: numberType);
        var source = IWorkSourceDocument.Open(package);
        var content = Assert.Single(source.ReadPages().Tables).Cells[0].RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, 8, 3, 4 }, content.Paragraphs.Select(p => p.ListNumber));
        Assert.All(content.Paragraphs, p => Assert.Equal(sample, p.ListLabel));
        Assert.Equal(new[] { "7. Value", "8. Next", "3. Restart", "4. Continue" },
            content.Paragraphs.Select(p => IWorkReadProjection.RichTextMarkdown(p)));
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true });
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes);
        Assert.Equal(markers, WordDocumentTraversal.BuildListMarkers(saved).Values.Select(value => value.Marker));
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData(1UL, PowerPointNumberingScheme.ArabicParenBoth)]
    [InlineData(2UL, PowerPointNumberingScheme.ArabicParenR)]
    public void Decimal_marker_forms_use_existing_PPTX_numbering_and_typed_starts(ulong numberType,
        PowerPointNumberingScheme scheme) {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Keynote, numberType: numberType);
        using var result = IWorkSourceDocument.Open(package).ToPowerPointDocumentResult(new() {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var content = Assert.Single(Assert.Single(result.Projection.Slides).Tables).Cells[0].RichText!;
        Assert.True(content.IsFormattingComplete);
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var saved = PowerPointDocument.Load(bytes);
        var paragraphs = Assert.Single(saved.Slides[0].Tables).GetCell(0, 0).Paragraphs;
        Assert.Equal(new int?[] { 7, null, 3, null }, paragraphs.Select(p => p.NumberingStartAt));
        Assert.All(paragraphs, p => Assert.Equal(scheme, p.NumberingScheme));
        Assert.Empty(saved.ValidateDocument());
    }
}
