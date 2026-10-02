using OfficeIMO.IWork;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Invalid_child_list_labels_preserve_parent_positions_and_physical_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = ListDeclarationPackage(kind,
            Message(BytesField(16, new byte[] { 0xff }), StringField(16, "9.")), aliases: true);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }), kind).Item1;
        Assert.Equal(2, table.Cells.Count);
        foreach (IWorkTableCell cell in table.Cells) {
            IWorkTextParagraph paragraph = Assert.Single(cell.RichText!.Paragraphs);
            Assert.Equal("Value", paragraph.Text);
            Assert.Equal(1, paragraph.ListLevel);
            Assert.Equal("c.", paragraph.ListLabel);
            Assert.False(cell.RichText.IsFormattingComplete);
        }
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: true,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false,
                MaximumSourceDeclarationIssues = 1 });
        AssertListDeclaration(Assert.Single(report.SourceDeclarationIssues), "16", 2);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData("label-wire", "16", 2)]
    [InlineData("type-wire", "11", 2)]
    [InlineData("type-packed", "11", 1)]
    [InlineData("type-enum", "11", 2)]
    [InlineData("indent-wire", "13", 2)]
    [InlineData("indent-finite", "13", 2)]
    public void Invalid_selected_list_vectors_do_not_overlay_qualified_parent_metadata(
        string defect, string path, int count) {
        byte[] fields = defect switch {
            "label-wire" => Message(StringField(16, "9."), VarintField(16, 7)),
            "type-wire" => Message(VarintField(11, 0), FloatField(11, 1)),
            "type-packed" => BytesField(11, new byte[] { 0x80 }),
            "type-enum" => Message(VarintField(11, 0), VarintField(11, 999)),
            "indent-wire" => Message(FloatField(13, 18), VarintField(13, 7)),
            _ => Message(FloatField(13, 18), FloatField(13, float.NaN))
        };
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers, fields);
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(text.IsFormattingComplete);
        IWorkTextParagraph paragraph = Assert.Single(text.Paragraphs);
        Assert.Equal(1, paragraph.ListLevel);
        Assert.Equal("c.", paragraph.ListLabel);
        package.Position = 0;
        AssertListDeclaration(Assert.Single(ConvertUnitReport(package,
            IWorkDocumentKind.Numbers).SourceDeclarationIssues), path, count);
    }

    [Fact]
    public void Rejected_list_labels_without_a_parent_do_not_shift_later_labels() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 1), VarintField(11, 1), FloatField(13, 0), FloatField(13, 18),
                BytesField(16, new byte[] { 0xff }), StringField(16, "9.")), inherited: false, indent: 0);
        IWorkTextParagraph paragraph = Assert.Single(Assert.Single(ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!.Paragraphs);
        Assert.Equal(0, paragraph.ListLevel);
        Assert.Null(paragraph.ListLabel);
    }

    [Fact]
    public void Valid_mixed_packed_and_unpacked_list_types_overlay_parent_vectors() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 0), BytesField(11, new byte[] { 2, 3 }),
                FloatField(13, 0), FloatField(13, 18), FloatField(13, 36),
                StringField(16, ""), StringField(16, "a."), StringField(16, "4.")), indent: 36);
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        IWorkTextParagraph paragraph = Assert.Single(text.Paragraphs);
        Assert.Equal(2, paragraph.ListLevel);
        Assert.Equal("4.", paragraph.ListLabel);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Parent_list_level_and_start_survive_saved_partial_editable_output(IWorkDocumentKind kind) {
        using MemoryStream package = InvalidListLabelPackage(kind);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.True(result.Report.IsPartialEditableReconstruction);
            AssertListDeclaration(Assert.Single(result.Report.SourceDeclarationIssues), "16", 2);
            result.Value.Save(saved); saved.Position = 0;
            using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
            MainDocumentPart main = document.MainDocumentPart!;
            Paragraph paragraph = Assert.Single(main.Document!.Body!.Descendants<Paragraph>(), item => item.InnerText == "Value");
            Assert.Equal(1, paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value);
            int numberId = paragraph.ParagraphProperties.NumberingProperties.NumberingId!.Val!.Value;
            Numbering numbering = main.NumberingDefinitionsPart!.Numbering!;
            int abstractId = Assert.Single(numbering.Elements<NumberingInstance>(), item => item.NumberID!.Value == numberId)
                .AbstractNumId!.Val!.Value;
            Level level = Assert.Single(Assert.Single(numbering.Elements<AbstractNum>(),
                item => item.AbstractNumberId!.Value == abstractId).Elements<Level>(), item => item.LevelIndex!.Value == 1);
            Assert.Equal(NumberFormatValues.LowerLetter, level.NumberingFormat!.Val!.Value);
            Assert.Equal(3, level.StartNumberingValue!.Val!.Value);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.True(result.Report.IsPartialEditableReconstruction);
            AssertListDeclaration(Assert.Single(result.Report.SourceDeclarationIssues), "16", 2);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var paragraph = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs[0];
            Assert.Equal("Value", paragraph.Text);
            Assert.Equal(1, paragraph.Level);
            Assert.Equal(3, paragraph.NumberingStartAt);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData("declaration", "declaration issues")]
    [InlineData("text", "text character")]
    public void Selected_list_resource_limits_remain_fatal(string limit, string message) {
        byte[] fields = limit switch {
            "declaration" => Message(BytesField(16, new byte[] { 0xff }), FloatField(13, float.NaN)),
            _ => Message(BytesField(16, new byte[] { 0xff }), StringField(16, new string('X', 100)))
        };
        var options = limit switch {
            "declaration" => new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 },
            _ => new IWorkReadOptions { MaximumProjectedTextCharacters = 10 }
        };
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers, fields);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, options);
        Assert.Contains(message, Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0, "1.")]
    [InlineData(1, "(1)")]
    [InlineData(2, "1)")]
    [InlineData(3, "I.")]
    [InlineData(4, "(I)")]
    [InlineData(5, "I)")]
    [InlineData(6, "i.")]
    [InlineData(7, "(i)")]
    [InlineData(8, "i)")]
    [InlineData(9, "A.")]
    [InlineData(10, "(A)")]
    [InlineData(11, "A)")]
    [InlineData(12, "a.")]
    [InlineData(13, "(a)")]
    [InlineData(14, "a)")]
    [InlineData(48, null)]
    public void Native_number_kinds_select_markers_without_claiming_counter_fidelity(int kind, string? marker) {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Pages,
            Message(VarintField(11, 3), VarintField(15, (ulong)kind), StringField(16, "•")),
            inherited: false, indent: 0);
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Pages).Item1.Cells).RichText!;
        Assert.Equal(marker, Assert.Single(text.Paragraphs).ListLabel);
        Assert.False(text.IsFormattingComplete);
    }

    [Theory]
    [InlineData("wire")]
    [InlineData("packed")]
    [InlineData("enum")]
    public void Invalid_native_number_vector_retains_physical_declaration_evidence(string defect) {
        byte[] invalid = defect switch {
            "wire" => FloatField(15, 0),
            "packed" => BytesField(15, new byte[] { 0x80 }),
            _ => VarintField(15, 65)
        };
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 3), invalid), inherited: false, indent: 0);
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        AssertListDeclaration(Assert.Single(report.SourceDeclarationIssues), "15", 1);
    }

    [Fact]
    public void Native_decimal_kind_survives_saved_keynote_table_text() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Keynote,
            Message(VarintField(11, 3), BytesField(15, new byte[] { 0 })),
            inherited: false, indent: 0);
        using var result = IWorkSourceDocument.Open(package).ToPowerPointPresentationResult(
            new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
        var paragraph = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs[0];
        Assert.Equal("Value", paragraph.Text);
        Assert.Equal(OfficeIMO.PowerPoint.PowerPointNumberingScheme.ArabicPeriod, paragraph.NumberingScheme);
        Assert.Equal(1, paragraph.NumberingStartAt);
        Assert.Empty(reopened.ValidateDocument());
    }

    private static MemoryStream InvalidListLabelPackage(IWorkDocumentKind kind) => ListDeclarationPackage(kind,
        Message(BytesField(16, new byte[] { 0xff }), StringField(16, "9.")));

    private static void AssertListDeclaration(IWorkSourceDeclarationIssue issue, string path, int count) {
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(count, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
    }

    private static MemoryStream ListDeclarationPackage(IWorkDocumentKind kind, byte[] fields,
        bool aliases = false, bool inherited = true, float indent = 18) => SelectedRichTextPackage(kind,
            Message(AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))),
                AttributeTable(5, AttributeEntry(0, ReferenceField(2, 21)))), aliases: aliases,
            additionalRecords: Message(
                ArchiveRecord(19, 2023, Message(inherited ? BytesField(1, ReferenceField(3, 20)) : Message(), fields)),
                ArchiveRecord(20, 2023, Message(VarintField(11, 1), VarintField(11, 1),
                    FloatField(13, 0), FloatField(13, 18), StringField(16, "1."), StringField(16, "c."))),
                ArchiveRecord(21, 2022, BytesField(12, FloatField(11, indent)))));
}
