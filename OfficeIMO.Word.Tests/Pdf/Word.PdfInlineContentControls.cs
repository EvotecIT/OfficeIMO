using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfInlineContentControlTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlinePlaceholderDateRetainsCachedDisplayAndRunGeometryWithAFormWidget(bool inTable) {
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        AddText(paragraph, "LEAD ", 10);
        WordDatePicker date = paragraph.AddDatePicker(new DateTime(2026, 2, 21), tag: "InlineDate");
        date._sdtRun.SdtProperties!.PrependChild(new W.RunProperties(new W.FontSize { Val = "20" }));
        date._sdtRun.SdtProperties.InsertBefore(new W.ShowingPlaceholder(),
            date._sdtRun.SdtProperties.GetFirstChild<W.SdtContentDate>());
        date._sdtRun.SdtContentRun = new W.SdtContentRun(new W.Run(
            new W.RunProperties(new W.Color { Val = "808080" }),
            new W.Text("Choose a date") { Space = SpaceProcessingModeValues.Preserve }));
        AddText(paragraph, " TAIL", 10);
        AddText(source.AddParagraph(), "FOLLOW", 10);

        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        AddText(referenceParagraph, "LEAD ", 10);
        AddText(referenceParagraph, "Choose a date", null, "808080");
        AddText(referenceParagraph, " TAIL", 10);
        AddText(reference.AddParagraph(), "FOLLOW", 10);

        byte[] bytes = ConvertSavedDocument(source);
        using PdfPigDocument pdf = OpenVisiblePdf(bytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var referencePage = expected.GetPage(1);
        Assert.Contains("Choose a date", page.Text);
        Assert.DoesNotContain("2026-02-21", page.Text);
        foreach (string marker in new[] { "LEAD", "TAIL", "FOLLOW" }) {
            var word = Assert.Single(page.GetWords(), word => word.Text == marker);
            var expectedWord = Assert.Single(referencePage.GetWords(), word => word.Text == marker);
            Assert.InRange(Math.Abs(word.BoundingBox.Left - expectedWord.BoundingBox.Left), 0D, .05D);
            Assert.InRange(Math.Abs(word.BoundingBox.Bottom - expectedWord.BoundingBox.Bottom), 0D, .05D);
        }
        var cachedWord = Assert.Single(page.GetWords(), word => word.Text == "Choose");
        Assert.All(cachedWord.Letters, letter => Assert.Equal(12D, letter.FontSize, 2));

        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("InlineDate", field.Name);
        Assert.True(field.IsTextField);
        PdfCore.PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        var tail = Assert.Single(page.GetWords(), word => word.Text == "TAIL");
        Assert.InRange(widget.X1, lead.BoundingBox.Right, cachedWord.BoundingBox.Left + .1D);
        Assert.InRange(widget.X2, cachedWord.BoundingBox.Right, tail.BoundingBox.Left + .1D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void InlineCheckBoxesKeepAuthoredOrderBaselineAndFieldState(bool inTable, bool isChecked) {
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        AddText(paragraph, "LEAD ", 10);
        AddText(referenceParagraph, "LEAD ", 10);
        string[] labels = { "FIRST", "SECOND", "THIRD" };
        for (int index = 0; index < labels.Length; index++) {
            WordCheckBox checkBox = paragraph.AddCheckBox(isChecked, tag: "InlineCheck" + index);
            W.Run symbol = checkBox._sdtRun.SdtContentRun!.Elements<W.Run>().Single();
            symbol.RunProperties ??= new W.RunProperties();
            symbol.RunProperties.FontSize = new W.FontSize { Val = "20" };
            referenceParagraph._paragraph.Append(symbol.CloneNode(true));
            AddText(paragraph, " " + labels[index] + " ", 10);
            AddText(referenceParagraph, " " + labels[index] + " ", 10);
        }
        AddText(source.AddParagraph(), "FOLLOW", 10);
        AddText(reference.AddParagraph(), "FOLLOW", 10);

        byte[] bytes = ConvertSavedDocument(source);
        using PdfPigDocument pdf = OpenVisiblePdf(bytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        double baseline = lead.Letters[0].StartBaseLine.Y;
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(labels.Length, info.FormFields.Count);
        double previousRight = lead.BoundingBox.Right;
        for (int index = 0; index < labels.Length; index++) {
            var label = Assert.Single(page.GetWords(), word => word.Text == labels[index]);
            Assert.InRange(Math.Abs(label.Letters[0].StartBaseLine.Y - baseline), 0D, .05D);
            PdfCore.PdfFormField field = Assert.Single(info.FormFields, field => field.Name == "InlineCheck" + index);
            Assert.True(field.IsCheckBox);
            Assert.Equal(isChecked ? "Yes" : "Off", field.Value);
            PdfCore.PdfFormWidget widget = Assert.Single(field.Widgets);
            Assert.Equal(1, widget.PageNumber);
            Assert.InRange(widget.X1, previousRight, label.BoundingBox.Left);
            Assert.InRange(widget.X2, widget.X1 + .01D, label.BoundingBox.Left + .1D);
            Assert.InRange(baseline, widget.Y1 - .1D, widget.Y2 + .1D);
            previousRight = label.BoundingBox.Right;
        }
        var following = Assert.Single(page.GetWords(), word => word.Text == "FOLLOW");
        var expectedFollowing = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == "FOLLOW");
        Assert.InRange(Math.Abs(following.Letters[0].StartBaseLine.Y - expectedFollowing.Letters[0].StartBaseLine.Y), 0D, .1D);
    }

    [Theory]
    [InlineData(false, false, "H", "O")]
    [InlineData(false, true, "H", "O")]
    [InlineData(true, false, "H", "O")]
    [InlineData(true, true, "H", "O")]
    [InlineData(false, true, "I", "W")]
    [InlineData(true, true, "I", "W")]
    public void InlineCheckBoxUsesAuthoredStateGlyphsBeforeAndAfterFilling(bool inTable, bool isChecked,
        string checkedGlyph, string uncheckedGlyph) {
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        AddText(paragraph, "LEAD ", 10);
        WordCheckBox control = paragraph.AddCheckBox(isChecked, tag: "GlyphCheck");
        W14.SdtContentCheckBox properties = control._sdtRun.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
        properties.GetFirstChild<W14.CheckedState>()!.Val = char.ConvertToUtf32(checkedGlyph, 0)
            .ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
        properties.GetFirstChild<W14.CheckedState>()!.Font = "Arial";
        properties.GetFirstChild<W14.UncheckedState>()!.Val = char.ConvertToUtf32(uncheckedGlyph, 0)
            .ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
        properties.GetFirstChild<W14.UncheckedState>()!.Font = "Arial";
        W.Run cache = control._sdtRun.SdtContentRun!.GetFirstChild<W.Run>()!;
        cache.RunProperties = new W.RunProperties(new W.RunFonts {
            Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial"
        }, new W.FontSize { Val = "20" });
        string currentGlyph = isChecked ? checkedGlyph : uncheckedGlyph;
        string nextGlyph = isChecked ? uncheckedGlyph : checkedGlyph;
        cache.GetFirstChild<W.Text>()!.Text = currentGlyph;
        AddText(paragraph, " TAIL", 10);
        AddText(source.AddParagraph(), "FOLLOW", 10);

        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        AddText(referenceParagraph, "LEAD ", 10);
        referenceParagraph._paragraph.Append(cache.CloneNode(true));
        AddText(referenceParagraph, " TAIL", 10);
        AddText(reference.AddParagraph(), "FOLLOW", 10);

        byte[] bytes = ConvertSavedDocument(source);
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("GlyphCheck", field.Name);
        Assert.True(field.IsCheckBox);
        Assert.Equal(isChecked ? "Yes" : "Off", field.Value);
        PdfCore.PdfFormWidget widget = Assert.Single(field.Widgets);
        using PdfPigDocument initial = OpenVisiblePdf(bytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference));
        var initialPage = initial.GetPage(1);
        var initialGlyph = Assert.Single(initialPage.GetWords(), word => word.Text == currentGlyph);
        var expectedGlyph = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == currentGlyph);
        Assert.DoesNotContain(initialPage.GetWords(), word => word.Text == nextGlyph);
        Assert.Contains("Arial", initialGlyph.Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(10D, initialGlyph.Letters[0].FontSize, 2);
        Assert.InRange(Math.Abs(initialGlyph.BoundingBox.Width - expectedGlyph.BoundingBox.Width), 0D, .05D);
        Assert.InRange(Math.Abs(initialGlyph.BoundingBox.Height - expectedGlyph.BoundingBox.Height), 0D, .05D);
        Assert.InRange(Math.Abs(initialGlyph.Letters[0].StartBaseLine.X - expectedGlyph.Letters[0].StartBaseLine.X), 0D, .05D);
        Assert.InRange(Math.Abs(initialGlyph.Letters[0].StartBaseLine.Y - expectedGlyph.Letters[0].StartBaseLine.Y), 0D, .05D);
        Assert.InRange(initialGlyph.Letters[0].StartBaseLine.X, widget.X1 - .05D, widget.X2);
        var tail = Assert.Single(initialPage.GetWords(), word => word.Text == "TAIL");
        Assert.InRange(Math.Abs(initialGlyph.Letters[0].StartBaseLine.Y - tail.Letters[0].StartBaseLine.Y), 0D, .05D);

        byte[] filledBytes = PdfCore.PdfDocument.Load(bytes).Forms.Fill(new Dictionary<string, string> {
            ["GlyphCheck"] = isChecked ? "Off" : "Yes"
        }).ToBytes();
        PdfCore.PdfFormField filledField = Assert.Single(PdfCore.PdfInspector.Inspect(filledBytes).FormFields);
        Assert.Equal(isChecked ? "Off" : "Yes", filledField.Value);
        using PdfPigDocument filled = OpenVisiblePdf(filledBytes);
        Assert.Equal(initial.NumberOfPages, filled.NumberOfPages);
        var filledPage = filled.GetPage(1);
        var filledGlyph = Assert.Single(filledPage.GetWords(), word => word.Text == nextGlyph);
        Assert.DoesNotContain(filledPage.GetWords(), word => word.Text == currentGlyph);
        Assert.Contains("Arial", filledGlyph.Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(10D, filledGlyph.Letters[0].FontSize, 2);
        Assert.InRange(filledGlyph.Letters[0].GlyphRectangle.Left, widget.X1 - .05D, widget.X2);
        Assert.InRange(filledGlyph.Letters[0].GlyphRectangle.Right, widget.X1, widget.X2 + .05D);
        Assert.InRange(Math.Abs(filledGlyph.Letters[0].StartBaseLine.Y - initialGlyph.Letters[0].StartBaseLine.Y), 0D, .05D);
        foreach (string marker in new[] { "LEAD", "TAIL", "FOLLOW" }) {
            var initialWord = Assert.Single(initialPage.GetWords(), word => word.Text == marker);
            var filledWord = Assert.Single(filledPage.GetWords(), word => word.Text == marker);
            Assert.Equal(initialWord.Letters[0].StartBaseLine, filledWord.Letters[0].StartBaseLine);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyCachedCheckBoxRetainsItsEditableField(bool inTable) {
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        AddText(paragraph, "LABEL", 10);
        WordCheckBox control = paragraph.AddCheckBox(true, tag: "EmptyCachedCheck");
        control._sdtRun.SdtContentRun = new W.SdtContentRun(new W.Run());
        byte[] bytes = ConvertSavedDocument(source);
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("EmptyCachedCheck", field.Name);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        Assert.Equal(1, Assert.Single(field.Widgets).PageNumber);
        using PdfPigDocument visible = OpenVisiblePdf(bytes);
        Assert.Equal(1, visible.NumberOfPages);
        Assert.Contains("LABEL", visible.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false, "date")]
    [InlineData(true, "date")]
    [InlineData(false, "dropdown")]
    [InlineData(true, "dropdown")]
    [InlineData(false, "combo")]
    [InlineData(true, "combo")]
    public void SelectedInlineControlsPreserveCachedDisplayLogicalValuesAndOptions(bool inTable, string kind) {
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph);
        AddText(paragraph, "LEAD ", 10);
        const string choiceValue = "Blue option";
        string display = kind == "date" ? "29 May 2026" : choiceValue;
        W.SdtRun control;
        if (kind == "date") {
            control = paragraph.AddDatePicker(new DateTime(2026, 5, 29), tag: "SelectedControl")._sdtRun;
            control.SdtProperties!.GetFirstChild<W.SdtContentDate>()!.Append(new W.DateFormat { Val = "dd MMM yyyy" });
        } else if (kind == "dropdown") {
            WordDropDownList dropdown = paragraph.AddDropDownList(new[] { "Red option", choiceValue }, tag: "SelectedControl");
            dropdown.SelectedValue = choiceValue;
            control = dropdown._sdtRun;
        } else {
            control = paragraph.AddComboBox(new[] { "Red option", choiceValue }, tag: "SelectedControl", defaultValue: choiceValue)._sdtRun;
        }
        control.SdtContentRun = new W.SdtContentRun(new W.Run(new W.Text(display)));
        AddText(paragraph, " TAIL", 10);
        AddText(source.AddParagraph(), "FOLLOW", 10);

        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph);
        AddText(referenceParagraph, "LEAD ", 10);
        AddText(referenceParagraph, display, null);
        AddText(referenceParagraph, " TAIL", 10);
        AddText(reference.AddParagraph(), "FOLLOW", 10);

        byte[] bytes = ConvertSavedDocument(source);
        using PdfPigDocument pdf = OpenVisiblePdf(bytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        Assert.Contains(display, page.Text);
        foreach (string marker in new[] { "LEAD", "TAIL", "FOLLOW" }) {
            var word = Assert.Single(page.GetWords(), word => word.Text == marker);
            var expectedWord = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == marker);
            Assert.InRange(Math.Abs(word.BoundingBox.Left - expectedWord.BoundingBox.Left), 0D, .05D);
            Assert.InRange(Math.Abs(word.BoundingBox.Bottom - expectedWord.BoundingBox.Bottom), 0D, .05D);
        }
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("SelectedControl", field.Name);
        Assert.Equal(kind == "date" ? "2026-05-29" : choiceValue, field.Value);
        Assert.Single(field.Widgets);
        if (kind == "date") Assert.True(field.IsTextField);
        else {
            Assert.True(field.IsChoiceField);
            Assert.Equal(new[] { "Red option", choiceValue }, field.Options.Select(option => option.ExportValue));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineTextControlWrapsWithItsSourceFormattingAndKeepsOneEditableField(bool inTable) {
        const string first = "ALPHA BETA GAMMA DELTA ";
        const string second = "EPSILON ZETA ETA THETA";
        string selected = first + second;
        using WordDocument source = CreateDocument(inTable, out WordParagraph paragraph, tableWidth: 2800);
        AddText(paragraph, "LEAD ", 10);
        W.SdtRun control = paragraph.AddComboBox(new[] { selected, "Other option" },
            tag: "WrappedControl", defaultValue: selected)._sdtRun;
        var firstRun = new W.Run(new W.RunProperties(new W.FontSize { Val = "20" }),
            new W.Text(first) { Space = SpaceProcessingModeValues.Preserve });
        var secondRun = new W.Run(new W.RunProperties(new W.Bold(), new W.FontSize { Val = "28" }),
            new W.Text(second));
        control.SdtContentRun = new W.SdtContentRun(firstRun, secondRun);
        AddText(paragraph, " TAIL", 10);
        AddText(source.AddParagraph(), "FOLLOW", 10);

        using WordDocument reference = CreateDocument(inTable, out WordParagraph referenceParagraph, tableWidth: 2800);
        AddText(referenceParagraph, "LEAD ", 10);
        referenceParagraph._paragraph.Append(firstRun.CloneNode(true), secondRun.CloneNode(true));
        AddText(referenceParagraph, " TAIL", 10);
        AddText(reference.AddParagraph(), "FOLLOW", 10);
        var options = new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PdfCore.PageSize(220, 400),
            Margins = new PdfCore.PageMargins(24, 24, 24, 24)
        };
        byte[] bytes = ConvertSavedDocument(source, options);
        using PdfPigDocument pdf = OpenVisiblePdf(bytes);
        using PdfPigDocument expected = PdfPigDocument.Open(ConvertSavedDocument(reference, options));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        foreach (string marker in ("LEAD " + selected + " TAIL FOLLOW").Split(' ')) {
            var word = Assert.Single(page.GetWords(), word => word.Text == marker);
            var expectedWord = Assert.Single(expected.GetPage(1).GetWords(), word => word.Text == marker);
            Assert.InRange(Math.Abs(word.Letters[0].StartBaseLine.X - expectedWord.Letters[0].StartBaseLine.X), 0D, .05D);
            Assert.InRange(Math.Abs(word.Letters[0].StartBaseLine.Y - expectedWord.Letters[0].StartBaseLine.Y), 0D, .05D);
            Assert.Equal(expectedWord.Letters[0].FontSize, word.Letters[0].FontSize, 2);
        }
        var largeWord = Assert.Single(page.GetWords(), word => word.Text == "EPSILON");
        Assert.All(largeWord.Letters, letter => Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase));
        PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        Assert.Equal("WrappedControl", field.Name);
        Assert.Equal(selected, field.Value);
        Assert.True(field.IsChoiceField);
        Assert.True(field.WidgetCount >= 2);
        Assert.All(field.Widgets, widget => Assert.Equal(1, widget.PageNumber));
    }

    private static WordDocument CreateDocument(bool inTable, out WordParagraph paragraph, int tableWidth = 9000) {
        WordDocument document = WordDocument.Create();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" })));
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties();
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        });
        if (inTable) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.WidthType = WordTableWidthUnit.Dxa;
            table.Width = tableWidth;
            table.ColumnWidthType = WordTableWidthUnit.Dxa;
            table.ColumnWidth = new List<int> { tableWidth };
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        } else paragraph = document.AddParagraph();
        paragraph._paragraph.RemoveAllChildren();
        return document;
    }

    private static void AddText(WordParagraph paragraph, string text, double? size, string? color = null) {
        var properties = new W.RunProperties();
        if (size.HasValue) properties.FontSize = new W.FontSize { Val = (size.Value * 2).ToString(System.Globalization.CultureInfo.InvariantCulture) };
        if (color != null) properties.Color = new W.Color { Val = color };
        paragraph._paragraph.Append(new W.Run(properties, new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static byte[] ConvertSavedDocument(WordDocument source, WordToPdfOptions? options = null) {
        byte[] docx = source.ToBytes();
        using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(docx), false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(package));
        using WordDocument loaded = WordDocument.Load(new MemoryStream(docx));
        return loaded.ToPdfBytes(options ?? new WordToPdfOptions { IncludePageNumbers = false });
    }

    // PdfPig reads page content. Flatten a separate copy through the public PDF
    // operation to inspect the real widget appearances; retain the original PDF
    // bytes above for field values, states and widget placement assertions.
    private static PdfPigDocument OpenVisiblePdf(byte[] bytes) =>
        PdfPigDocument.Open(PdfCore.PdfDocument.Load(bytes).Forms.Flatten().ToBytes());
}
