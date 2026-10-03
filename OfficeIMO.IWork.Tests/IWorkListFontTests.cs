using OfficeIMO.IWork;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Marker_font_survives_saved_editable_table_lists(IWorkDocumentKind kind) {
        using MemoryStream package = ListDeclarationPackage(kind,
            Message(VarintField(11, 2), StringField(16, "•"), StringField(23, "Courier New")),
            inherited: false, indent: 0);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        IWorkTextParagraph sourceParagraph = Assert.Single(Assert.Single(
            ReadSelectedRichTable(source, kind).Item1.Cells).RichText!.Paragraphs);
        Assert.Equal("Courier New", sourceParagraph.ListFontName);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            result.Value.Save(saved); saved.Position = 0;
            using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
            var paragraph = Assert.Single(document.MainDocumentPart!.Document!.Descendants<Paragraph>(),
                item => item.InnerText == "Value");
            int numberId = paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            Numbering numbering = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
            int abstractId = Assert.Single(numbering.Elements<NumberingInstance>(), item => item.NumberID!.Value == numberId)
                .AbstractNumId!.Val!.Value;
            Level level = Assert.Single(Assert.Single(numbering.Elements<AbstractNum>(),
                item => item.AbstractNumberId!.Value == abstractId).Elements<Level>());
            Assert.Equal("Courier New", level.NumberingSymbolRunProperties!.RunFonts!.Ascii!.Value);
            Assert.Equal("Courier New", level.NumberingSymbolRunProperties.RunFonts.HighAnsi!.Value);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document));
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var paragraph = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs[0];
            Assert.Equal("Value", paragraph.Text);
            Assert.Equal("Courier New", paragraph.BulletFontName);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData("inherit", "Symbol", true)]
    [InlineData("replace", "Courier New", true)]
    [InlineData("clear", null, true)]
    [InlineData("utf8", null, false)]
    [InlineData("duplicate", null, false)]
    [InlineData("conflict", null, false)]
    [InlineData("flag", null, false)]
    public void Marker_font_inheritance_and_invalid_child_declarations_are_explicit(
        string declaration, string? expected, bool complete) {
        byte[] child = declaration switch {
            "inherit" => Message(),
            "replace" => StringField(23, "Courier New"),
            "clear" => VarintField(22, 1),
            "utf8" => BytesField(23, new byte[] { 0xff }),
            "duplicate" => Message(StringField(23, "A"), StringField(23, "B")),
            "conflict" => Message(VarintField(22, 1), StringField(23, "A")),
            _ => VarintField(22, 2)
        };
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers, child,
            parentFields: StringField(23, "Symbol"));
        var source = IWorkSourceDocument.Open(package);
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(source, IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.Equal(expected, Assert.Single(text.Paragraphs).ListFontName);
        Assert.Equal(complete, text.IsFormattingComplete);
        if (!complete) {
            package.Position = 0;
            var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers,
                readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
            AssertListDeclaration(Assert.Single(report.SourceDeclarationIssues), declaration == "flag" ? "22" : "23",
                declaration == "duplicate" ? 2 : 1);
            Assert.Empty(report.PreservedRecords);
        }
    }

    [Fact]
    public void Reused_marker_fonts_remain_subject_to_projected_character_limits() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 2), StringField(16, "•"), StringField(23, new string('F', 100))),
            inherited: false, indent: 0, aliases: true);
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextCharacters = 250 });
        Assert.Contains("text character", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }
}
