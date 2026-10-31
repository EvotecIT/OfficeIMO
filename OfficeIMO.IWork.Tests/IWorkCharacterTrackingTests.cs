using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void PowerPoint_reports_character_spacing_quantization() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Keynote,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, Message(FloatField(3, 40), FloatField(27, 0.1234f)))));
        Assert.Contains(ConvertUnitReport(package, IWorkDocumentKind.Keynote, visual: false).FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_CHARACTER_SPACING_PRECISION"
                && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Character_tracking_uses_effective_font_size(IWorkDocumentKind kind) {
        using var package = SelectedRichTextPackage(kind,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, Message(FloatField(3, 40), FloatField(27, -0.25f)))));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTextRun run = Assert.Single(Assert.Single(Assert.Single(table.Cells).RichText!.Paragraphs).Runs);
        Assert.Equal(-0.25, run.Style.CharacterTrackingEm);
        Assert.Equal(-10, run.Style.CharacterSpacingPoints);
        Assert.True(Assert.Single(table.Cells).RichText!.IsFormattingComplete);
        package.Position = 0;
        Assert.DoesNotContain(ConvertUnitReport(package, kind).SourceDeclarationIssues, issue => issue.FieldPath == "11/27");
    }

    [Theory]
    [InlineData(0f, 0d)]
    [InlineData(0.25f, 10d)]
    public void Child_tracking_and_font_overrides_replace_inherited_tracking(float tracking, double spacing) {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Pages,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))), additionalRecords: Message(
                ArchiveRecord(19, 2021, Message(BytesField(1, ReferenceField(3, 20)), BytesField(11, Message(FloatField(3, 40), FloatField(27, tracking))))),
                ArchiveRecord(20, 2021, BytesField(11, Message(FloatField(3, 20), FloatField(27, -0.25f))))));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Pages);
        Assert.Equal(spacing, Assert.Single(Assert.Single(Assert.Single(table.Cells).RichText!.Paragraphs).Runs).Style.CharacterSpacingPoints);
    }
}

public sealed class IWorkCharacterTrackingCorpusTests {
    [Fact]
    public void Native_Keynote_tracking_matches_independent_Apple_PowerPoint_export() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        IWorkSourceDocument source = IWorkSourceDocument.Open(Path.Combine(root, "keynotekit", "imagedeck-v15.2.1.key"));
        IWorkTextStyle style = Assert.Single(source.ReadKeynote().Slides).TitleBox!.Content.Paragraphs[0].Runs[0].Style;
        Assert.Equal(-0.02, style.CharacterTrackingEm!.Value, 6);
        Assert.Equal(-2.32, style.CharacterSpacingPoints!.Value, 6);
        using var converted = source.ToPowerPointDocumentResult(new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream();
        converted.Value.Save(bytes);
        bytes.Position = 0;
        using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
        PowerPointTextRun title = Assert.Single(reopened.Slides[0].TextBoxes.First(box => box.Text == "Image Test").Paragraphs[0].Runs);
        Assert.Equal(-2.32, title.CharacterSpacing?.ToPoints());
        using PresentationDocument apple = PresentationDocument.Open(Path.Combine(root, "native-exports", "keynote-image-tracking-v15.4.pptx"), false);
        Assert.Contains(apple.PresentationPart!.SlideMasterParts.SelectMany(master => master.SlideLayoutParts)
            .SelectMany(layout => layout.SlideLayout!.Descendants<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>()),
            properties => properties.FontSize?.Value == 11600 && properties.Spacing?.Value == -232);
    }
}
