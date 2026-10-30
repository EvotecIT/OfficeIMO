using System.Text.Json;
using System.Security.Cryptography;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_native_table_text_styles_match_source_projection() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("table-text-styles.json")));
        foreach (JsonElement expectedSource in manifest.RootElement.GetProperty("sources").EnumerateArray()) {
            string path = CorpusFixture(expectedSource.GetProperty("path").GetString()!);
            Assert.Equal(expectedSource.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            var source = IWorkSourceDocument.Open(path);
            IEnumerable<IWorkTable> tables = source.Kind switch {
                IWorkDocumentKind.Pages => source.ReadPages().Tables,
                IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets.SelectMany(sheet => sheet.Tables),
                _ => source.ReadKeynote().Slides.SelectMany(slide => slide.Tables)
            };
            foreach (JsonElement expectedTable in expectedSource.GetProperty("tables").EnumerateArray()) {
                IWorkTable table = Assert.Single(tables, table => table.ModelRecord!.Identifier == expectedTable.GetProperty("modelIdentifier").GetUInt64());
                foreach (JsonElement cell in expectedTable.GetProperty("selected").EnumerateArray()) {
                    IWorkTableCell value = Assert.IsType<IWorkTableCell>(table.GetCell(cell.GetProperty("row").GetInt32(), cell.GetProperty("column").GetInt32()));
                    AssertNativeTableStyle(cell.GetProperty("properties"), Assert.IsType<IWorkParagraphStyle>(value.ParagraphStyle));
                }
                JsonElement roles = expectedTable.GetProperty("roles");
                foreach (var (name, style) in new[] { ("body_text_style", table.TextStyles.Body),
                    ("header_row_text_style", table.TextStyles.HeaderRow), ("header_column_text_style", table.TextStyles.HeaderColumn),
                    ("footer_row_text_style", table.TextStyles.FooterRow) }) {
                    bool applicable = name switch {
                        "body_text_style" => table.RowCount > table.HeaderRowCount + table.FooterRowCount && table.ColumnCount > table.HeaderColumnCount,
                        "header_row_text_style" => table.HeaderRowCount > 0,
                        "header_column_text_style" => table.RowCount > table.HeaderRowCount && table.HeaderColumnCount > 0,
                        _ => table.FooterRowCount > 0 && table.ColumnCount > table.HeaderColumnCount
                    };
                    if (applicable && roles.TryGetProperty(name, out var role))
                        AssertNativeTableStyle(role.GetProperty("properties"), Assert.IsType<IWorkParagraphStyle>(style));
                }
            }
        }
    }

    [Fact]
    public void Native_Keynote_table_fonts_alignment_and_pagination_loss_survive_saved_PPTX() {
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(CorpusFixture("keynotekit/tabledeck-v15.2.1.key"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.PowerPoint.PowerPointDocument.Load(saved);
        var table = Assert.Single(reopened.Slides.SelectMany(slide => slide.Tables));
        for (int row = 0; row < 3; row++) for (int column = 0; column < 3; column++) {
            var paragraph = Assert.Single(table.GetCell(row, column).Paragraphs);
            Assert.Equal(OfficeIMO.PowerPoint.PowerPointTextAlignment.Center, paragraph.Alignment);
            Assert.All(paragraph.Runs, run => {
                Assert.Equal(OfficeLength.Points(32), run.FontSize);
                Assert.Equal("HelveticaNeue", run.FontName);
                Assert.Equal(row == 0, run.Bold);
            });
        }
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
        using var strict = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(CorpusFixture("keynotekit/tabledeck-v15.2.1.key"), conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(strict.IsVisualFallback);
    }

    [Fact]
    public void Native_Pages_selected_table_fonts_survive_saved_DOCX_including_empty_rich_paragraphs() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture("picodocs/sample-v14.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        for (int index = 0; index < result.Projection.Tables.Count; index++) {
            IWorkTable table = result.Projection.Tables[index];
            foreach (IWorkTableCell source in table.Cells) {
                IWorkTextStyle style = table.GetParagraphStyle(source.Row, source.Column)!.TextStyle;
                var paragraph = reopened.Tables[index].Rows[source.Row - 1].Cells[source.Column - 1].Paragraphs[0];
                Assert.Equal(Math.Round(style.FontSizePoints!.Value * 2, MidpointRounding.AwayFromZero) / 2, paragraph.FontSize?.ToPoints());
                Assert.Equal(style.FontName, paragraph.FontFamily);
            }
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Table_role_precedence_selected_override_and_unstored_empty_cells_survive_saved_destinations(IWorkDocumentKind kind) {
        using var package = TableTextStylePackage(kind);
        var source = IWorkSourceDocument.Open(package);
        IWorkTable table = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().Tables[0],
            IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets[0].Tables[0],
            _ => source.ReadKeynote().Slides[0].Tables[0]
        };
        Assert.Equal(2, table.Cells.Count);
        Assert.Equal(IWorkCellKind.Number, table.GetCell(1, 1)!.Kind);
        Assert.Equal(42d, table.GetCell(1, 1)!.Value);
        Assert.Equal(18, table.GetParagraphStyle(1, 1)!.TextStyle.FontSizePoints);
        Assert.False(table.GetParagraphStyle(1, 1)!.TextStyle.Bold);
        Assert.Equal(IWorkTextAlignment.Right, table.GetParagraphStyle(1, 1)!.Alignment);
        Assert.Equal(18, table.GetParagraphStyle(1, 2)!.TextStyle.FontSizePoints);
        Assert.Equal(14, table.GetParagraphStyle(1, 3)!.TextStyle.FontSizePoints);
        Assert.Equal(16, table.GetParagraphStyle(2, 1)!.TextStyle.FontSizePoints);
        Assert.Equal(12, table.GetParagraphStyle(2, 2)!.TextStyle.FontSizePoints);
        Assert.Equal(16, table.GetParagraphStyle(3, 1)!.TextStyle.FontSizePoints);
        Assert.Equal(20, table.GetParagraphStyle(3, 2)!.TextStyle.FontSizePoints);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetParagraphStyle(0, 1));
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            var numeric = reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0];
            Assert.Equal("42", numeric.Text); Assert.False(numeric.Bold); Assert.Equal(18, numeric.FontSize?.ToPoints());
            Assert.Equal(OfficeIMO.Word.WordParagraphAlignment.Right, numeric.ParagraphAlignment);
            Assert.Equal(20, reopened.Tables[0].Rows[2].Cells[1].Paragraphs[0].FontSize?.ToPoints());
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            var sheet = reopened.Sheets[0]; Assert.True(sheet.TryGetCellValueSnapshot(1, 1, out var value)); Assert.Equal(OfficeIMO.Excel.ExcelCellValueKind.Number, value!.Kind); Assert.Equal("42", value.RawValue);
            Assert.Equal(18, sheet.GetCellStyle(1, 1).FontSize?.ToPoints()); Assert.False(sheet.GetCellStyle(1, 1).Bold);
            Assert.Equal("right", sheet.GetCellStyle(1, 1).HorizontalAlignment);
            Assert.Equal(20, sheet.GetCellStyle(3, 2).FontSize?.ToPoints());
        } else {
            using var result = source.ToPowerPointDocumentResult(policy); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointDocument.Load(saved);
            var target = Assert.Single(reopened.Slides[0].Tables);
            Assert.Equal(18, target.GetCell(0, 0).Paragraphs[0].Runs[0].FontSize?.ToPoints());
            Assert.False(target.GetCell(0, 0).Paragraphs[0].Runs[0].Bold);
            Assert.Equal(OfficeIMO.PowerPoint.PowerPointTextAlignment.Right, target.GetCell(0, 0).Paragraphs[0].Alignment);
            Assert.Equal(20, target.GetCell(2, 1).Paragraphs[0].Runs[0].FontSize?.ToPoints());
        }
        package.Position = 0;
        Assert.Throws<InvalidDataException>(() => {
            var bounded = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumMaterializedCells = 1 });
            _ = kind switch { IWorkDocumentKind.Pages => (object)bounded.ReadPages(), IWorkDocumentKind.Numbers => bounded.ReadNumbers(), _ => bounded.ReadKeynote() };
        });
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("wrong-type")]
    [InlineData("duplicate-character")]
    [InlineData("duplicate-paragraph")]
    [InlineData("cycle")]
    [InlineData("depth")]
    public void Unresolved_selected_text_style_keeps_values_and_does_not_silently_use_role_defaults(string defect) {
        byte[] style = defect switch {
            "missing" => Array.Empty<byte>(),
            "wrong-type" => ArchiveRecord(30, 6004, Message()),
            "duplicate-character" => ArchiveRecord(30, 2022, Message(BytesField(11, FloatField(3, 18)), BytesField(11, FloatField(3, 20)))),
            "duplicate-paragraph" => ArchiveRecord(30, 2022, Message(BytesField(12, VarintField(1, 1)), BytesField(12, VarintField(1, 2)))),
            "cycle" => TableParagraphStyle(30, 18, false, 30),
            _ => TableParagraphStyle(30, 18, false, 31)
        };
        using var package = TableTextStylePackage(IWorkDocumentKind.Pages, style);
        var projection = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTextStyleInheritanceDepth = defect == "depth" ? 1 : 128 }).ReadPages();
        var table = Assert.Single(projection.Tables);
        Assert.Equal(42d, table.GetCell(1, 1)!.Value);
        Assert.Null(table.GetParagraphStyle(1, 1)); Assert.NotNull(table.GetParagraphStyle(1, 3));
        Assert.Contains(projection.Diagnostics, d => d.Code == "IWORK_TABLE_TEXT_STYLE_UNSUPPORTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
        Assert.DoesNotContain(projection.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FILL_UNSUPPORTED" || d.Code == "IWORK_TABLE_CELL_LAYOUT_UNSUPPORTED");
        if (defect == "wrong-type") {
            var issue = Assert.Single(projection.SourceReferenceIssues);
            Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
            Assert.Equal(13ul, issue.Owner.RecordIdentifier);
            Assert.Equal("3[1]/4", issue.FieldPath);
            Assert.Equal(30ul, issue.TargetIdentifier);
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 1_100_000_000f)]
    [InlineData(IWorkDocumentKind.Numbers, 410f)]
    [InlineData(IWorkDocumentKind.Keynote, 4001f)]
    public void Selected_table_font_outside_destination_range_falls_back_even_under_partial_policy(IWorkDocumentKind kind, float size) {
        using var package = TableTextStylePackage(kind, TableParagraphStyle(30, size, false));
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.True(result.IsVisualFallback); Assert.Equal(size, result.Projection.Tables[0].GetParagraphStyle(1, 1)!.TextStyle.FontSizePoints);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.True(result.IsVisualFallback); Assert.Equal(size, result.Projection.Sheets[0].Tables[0].GetParagraphStyle(1, 1)!.TextStyle.FontSizePoints);
        } else {
            using var result = source.ToPowerPointDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.True(result.IsVisualFallback); Assert.Equal(size, result.Projection.Slides[0].Tables[0].GetParagraphStyle(1, 1)!.TextStyle.FontSizePoints);
        }
    }

    [Fact]
    public void Numbers_table_oversized_font_name_falls_back_before_cell_style_expansion() {
        string name = new string('A', 256);
        using var package = TableTextStylePackage(IWorkDocumentKind.Numbers,
            TableParagraphStyle(30, 18, false, fontName: name));
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(
            IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }));

        Assert.True(result.IsVisualFallback);
        Assert.Equal(name, result.Projection.Sheets[0].Tables[0].GetParagraphStyle(1, 1)!.TextStyle.FontName);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("font name", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Numbers_rich_text_run_oversized_font_name_falls_back_before_cell_style_expansion() {
        string name = new string('B', 256);
        using var package = TableTextStylePackage(IWorkDocumentKind.Numbers, richRunFontName: name);
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(
            IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }));

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.RichText!.Paragraphs
            .SelectMany(paragraph => paragraph.Runs), run => run.Style.FontName == name);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("rich-text font name", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Mixed_selected_text_and_cell_styles_share_one_catalog_budget_and_preserve_both_properties() {
        using var selected = TableTextStylePackage(IWorkDocumentKind.Pages, cellStyle: true);
        var limited = IWorkSourceDocument.Open(selected, new IWorkReadOptions { MaximumTableCatalogEntries = 2 });
        Assert.All(limited.ReadPages().Tables[0].Cells, cell => {
            Assert.Equal("FF0000", cell.Fill!.Color!.RgbHex);
            Assert.Equal(18, cell.ParagraphStyle!.TextStyle.FontSizePoints);
        });
        selected.Position = 0;
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(selected,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 }).ReadPages());
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Explicit_empty_rich_paragraph_font_overrides_table_defaults_in_saved_documents(IWorkDocumentKind kind) {
        using var package = TableTextStylePackage(kind, emptyRichParagraph: true);
        var source = IWorkSourceDocument.Open(package);
        IWorkTable table = kind == IWorkDocumentKind.Pages ? source.ReadPages().Tables[0] : source.ReadKeynote().Slides[0].Tables[0];
        Assert.Equal(18, table.GetParagraphStyle(1, 1)!.TextStyle.FontSizePoints);
        Assert.All(table.GetCell(1, 1)!.RichText!.Paragraphs, paragraph => {
            Assert.Empty(paragraph.Runs);
            Assert.Equal(36, paragraph.Style.TextStyle.FontSizePoints);
            Assert.True(paragraph.Style.TextStyle.Bold);
        });
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            var paragraphs = reopened.Tables[0].Rows[0].Cells[0].Paragraphs;
            Assert.Equal(2, paragraphs.Count);
            Assert.All(paragraphs, paragraph => { Assert.Equal(36, paragraph.FontSize?.ToPoints()); Assert.True(paragraph.Bold); });
            Assert.Equal(18, reopened.Tables[0].Rows[0].Cells[1].Paragraphs[0].FontSize?.ToPoints());
        } else {
            using var result = source.ToPowerPointDocumentResult(policy); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointDocument.Load(saved);
            var target = Assert.Single(reopened.Slides[0].Tables);
            Assert.Equal(2, target.GetCell(0, 0).Paragraphs.Count);
            Assert.All(target.GetCell(0, 0).Paragraphs, paragraph => Assert.All(paragraph.Runs,
                run => { Assert.Equal(36, run.FontSize?.ToPoints()); Assert.True(run.Bold); }));
            Assert.Equal(18, target.GetCell(0, 1).Paragraphs[0].Runs[0].FontSize?.ToPoints());
        }
    }

    [Theory]
    [InlineData(null, null, 12d, 8d)]
    [InlineData(0f, 0f, 0d, 0d)]
    [InlineData(4f, 6f, 4d, 6d)]
    public void Pages_rich_table_paragraph_spacing_preserves_defaults_and_explicit_overrides(
        float? before, float? after, double expectedBefore, double expectedAfter) {
        using var package = TableTextStylePackage(IWorkDocumentKind.Pages,
            TableParagraphStyle(30, 18, false, spaceBefore: 12, spaceAfter: 8),
            richRunFontName: "HelveticaNeue", richSpaceBefore: before, richSpaceAfter: after);
        var source = IWorkSourceDocument.Open(package);
        IWorkTable table = source.ReadPages().Tables[0];
        Assert.Equal(12d, table.GetParagraphStyle(1, 1)!.SpaceBeforePoints);
        Assert.Equal(8d, table.GetParagraphStyle(1, 1)!.SpaceAfterPoints);
        Assert.All(table.GetCell(1, 1)!.RichText!.Paragraphs, paragraph => {
            Assert.Equal((double?)before, paragraph.Style.SpaceBeforePoints);
            Assert.Equal((double?)after, paragraph.Style.SpaceAfterPoints);
        });
        using var result = source.ToWordDocumentResult(new IWorkConversionOptions {
            AllowPartialEditableReconstruction = true
        });
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        var cells = reopened.Tables[0].Rows[0].Cells;
        Assert.Contains("Styled", cells[0].Paragraphs.Select(paragraph => paragraph.Text));
        Assert.All(cells[0].Paragraphs, paragraph => {
            Assert.Equal(expectedBefore, paragraph.LineSpacingBefore?.ToPoints());
            Assert.Equal(expectedAfter, paragraph.LineSpacingAfter?.ToPoints());
        });
        Assert.Equal(12d, cells[1].Paragraphs[0].LineSpacingBefore?.ToPoints());
        Assert.Equal(8d, cells[1].Paragraphs[0].LineSpacingAfter?.ToPoints());
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 1_100_000_000f)]
    [InlineData(IWorkDocumentKind.Keynote, 4001f)]
    public void Empty_rich_paragraph_font_outside_destination_range_falls_back(IWorkDocumentKind kind, float fontSize) {
        using var package = TableTextStylePackage(kind, emptyRichParagraph: true, richParagraphFontSize: fontSize);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.True(result.IsVisualFallback);
            Assert.Equal(fontSize, result.Projection.Tables[0].GetCell(1, 1)!.RichText!.Paragraphs[0].Style.TextStyle.FontSizePoints);
        } else {
            using var result = source.ToPowerPointDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.True(result.IsVisualFallback);
            Assert.Equal(fontSize, result.Projection.Slides[0].Tables[0].GetCell(1, 1)!.RichText!.Paragraphs[0].Style.TextStyle.FontSizePoints);
        }
    }

    private static void AssertNativeTableStyle(JsonElement expected, IWorkParagraphStyle actual) {
        IWorkTextStyle text = actual.TextStyle;
        if (expected.TryGetProperty("font_size", out var size)) Assert.Equal(size.GetDouble(), text.FontSizePoints);
        if (expected.TryGetProperty("font_name", out var font)) Assert.Equal(font.GetString(), text.FontName);
        if (expected.TryGetProperty("bold", out var bold)) Assert.Equal(bold.GetBoolean(), text.Bold);
        if (expected.TryGetProperty("italic", out var italic)) Assert.Equal(italic.GetBoolean(), text.Italic);
        if (expected.TryGetProperty("underline", out var underline)) Assert.Equal(underline.GetInt32() != 0, text.Underline);
        if (expected.TryGetProperty("strikethru", out var strike)) Assert.Equal(strike.GetInt32() != 0, text.Strikethrough);
        if (expected.TryGetProperty("alignment", out var alignment)) Assert.Equal(alignment.GetInt32() switch {
            0 => IWorkTextAlignment.Left, 1 => IWorkTextAlignment.Right, 2 => IWorkTextAlignment.Center,
            3 => IWorkTextAlignment.Justified, _ => IWorkTextAlignment.Natural }, actual.Alignment);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Header_roles_do_not_add_bold_to_plain_or_rich_source_cell_fonts(IWorkDocumentKind kind, bool richText) {
        using var package = TableTextStylePackage(kind, TableParagraphStyle(30, 18, null),
            richRunFontName: richText ? "Arial" : null, richParagraphBold: null);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.Null(result.Projection.Tables[0].GetParagraphStyle(1, 1)!.TextStyle.Bold);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            var paragraph = richText ? Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs, value => value.Text == "Styled")
                : reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0];
            Assert.Equal(richText ? "Styled" : "42", paragraph.Text);
            Assert.False(paragraph.Bold);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy);
            Assert.Null(result.Projection.Sheets[0].Tables[0].GetParagraphStyle(1, 1)!.TextStyle.Bold);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            Assert.False(reopened.Sheets[0].GetCellStyle(1, 1).Bold);
            Assert.True(reopened.Sheets[0].TryGetCellValueSnapshot(1, 1, out var value));
            Assert.Equal(richText ? "Styled\n" : "42", richText ? value!.Text : value!.RawValue);
        } else {
            using var result = source.ToPowerPointDocumentResult(policy);
            Assert.Null(result.Projection.Slides[0].Tables[0].GetParagraphStyle(1, 1)!.TextStyle.Bold);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointDocument.Load(saved);
            var table = reopened.Slides[0].Tables.Single();
            Assert.False(table.FirstRow); Assert.False(table.FirstColumn); Assert.False(table.LastRow); Assert.False(table.BandedRows);
            Assert.Equal(richText ? "Styled" : "42", table.GetCell(0, 0).Text);
            Assert.All(table.GetCell(0, 0).Paragraphs.SelectMany(paragraph => paragraph.Runs), run => Assert.False(run.Bold));
        }
    }

    private static byte[] TableParagraphStyle(ulong id, float fontSize, bool? bold, ulong? parent = null, ulong alignment = 2,
        string fontName = "HelveticaNeue", float? spaceBefore = null, float? spaceAfter = null) =>
        ArchiveRecord(id, 2022, Message(parent.HasValue ? BytesField(1, ReferenceField(3, parent.Value)) : Array.Empty<byte>(),
            BytesField(11, Message(FloatField(3, fontSize), bold.HasValue ? VarintField(1, bold.Value ? 1UL : 0) : Array.Empty<byte>(), StringField(5, fontName))),
            BytesField(12, Message(VarintField(1, alignment),
                spaceAfter.HasValue ? FloatField(20, spaceAfter.Value) : Array.Empty<byte>(),
                spaceBefore.HasValue ? FloatField(21, spaceBefore.Value) : Array.Empty<byte>()))));

    private static MemoryStream TableTextStylePackage(IWorkDocumentKind kind, byte[]? selectedStyle = null, bool cellStyle = false,
        bool emptyRichParagraph = false, float richParagraphFontSize = 36, string? richRunFontName = null,
        bool? richParagraphBold = true, float? richSpaceBefore = null, float? richSpaceAfter = null) {
        bool richText = emptyRichParagraph || richRunFontName != null;
        byte[] first = new byte[cellStyle ? 28 : 24]; first[0] = 5; first[1] = 2; first[8] = cellStyle ? (byte)0x62 : (byte)0x42;
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, first, 12, 8); first[20] = cellStyle ? (byte)2 : (byte)1;
        if (cellStyle) first[24] = 1;
        if (richText) {
            first = new byte[20]; first[0] = 5; first[1] = 9; first[8] = 0x50; first[12] = 1; first[16] = 1;
        }
        byte[] second = new byte[cellStyle ? 20 : 16]; second[0] = 5; second[8] = cellStyle ? (byte)0x60 : (byte)0x40; second[12] = cellStyle ? (byte)2 : (byte)1;
        if (cellStyle) second[16] = 1;
        byte[] tile = Message(VarintField(1, 1), VarintField(2, 0), VarintField(3, 3), VarintField(4, 1),
            BytesField(5, Message(VarintField(1, 0), VarintField(2, 2), BytesField(6, Message(first, second)),
                BytesField(7, new byte[] { 0, 0, (byte)first.Length, 0, 255, 255 }))), VarintField(6, 5), VarintField(7, 1));
        var records = new List<byte[]>();
        if (kind == IWorkDocumentKind.Pages) { records.Add(ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2, 10 })); records.Add(ArchiveRecord(2, 2001, StringField(3, "Body"))); }
        else if (kind == IWorkDocumentKind.Numbers) { records.Add(ArchiveRecord(1, 1, ReferenceField(1, 2))); records.Add(ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10)))); }
        else { records.Add(ArchiveRecord(1, 1, ReferenceField(2, 2))); records.Add(ArchiveRecord(2, 2, KeynoteShow(ReferenceField(2, 3)))); records.Add(ArchiveRecord(3, 4, ReferenceField(2, 4))); records.Add(ArchiveRecord(4, 5, ReferenceField(6, 10))); }
        records.Add(ArchiveRecord(10, 6000, Message(ReferenceField(2, 11), kind == IWorkDocumentKind.Keynote ? BytesField(1, GeometryDrawable(0, 0, 144, 72)) : Array.Empty<byte>())));
        records.Add(ArchiveRecord(11, 6001, Message(VarintField(6, 3), VarintField(7, 3), VarintField(9, 1), VarintField(10, 1), VarintField(11, 1),
            ReferenceField(24, 40), ReferenceField(25, 41), ReferenceField(26, 42), ReferenceField(27, 43),
            BytesField(4, Message(ReferenceField(5, 13), richText ? ReferenceField(17, 14) : Array.Empty<byte>(),
                BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12)))))))));
        records.Add(ArchiveRecord(12, 6002, tile));
        records.Add(ArchiveRecord(13, 6005, Message(VarintField(1, 4), BytesField(3, Message(VarintField(1, 1), ReferenceField(4, 30))),
            cellStyle ? BytesField(3, Message(VarintField(1, 2), ReferenceField(4, 60))) : Array.Empty<byte>())));
        if (cellStyle) records.Add(FillStyle(60, FillColor(1, 0, 0)));
        if (richText) {
            records.Add(ArchiveRecord(14, 6005, Message(VarintField(1, 8), BytesField(3, Message(VarintField(1, 1), ReferenceField(9, 15))))));
            records.Add(ArchiveRecord(15, 6218, ReferenceField(1, 16)));
            records.Add(ArchiveRecord(16, 2001, Message(StringField(3, richRunFontName is null ? "\n" : "Styled\n"),
                BytesField(5, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 17)))))));
            records.Add(TableParagraphStyle(17, richParagraphFontSize, richParagraphBold,
                fontName: richRunFontName ?? "HelveticaNeue", spaceBefore: richSpaceBefore, spaceAfter: richSpaceAfter));
        }
        records.Add(selectedStyle ?? TableParagraphStyle(30, 18, false, 31, alignment: 1));
        records.Add(TableParagraphStyle(31, 24, true));
        records.Add(TableParagraphStyle(40, 12, false)); records.Add(TableParagraphStyle(41, 14, true));
        records.Add(TableParagraphStyle(42, 16, true)); records.Add(TableParagraphStyle(43, 20, true));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))), ("preview.png", ValidPreviewPng()));
    }
}
