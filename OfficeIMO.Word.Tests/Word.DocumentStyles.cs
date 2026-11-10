using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void DocumentStyles_CreatedEditedAndAppliedStylesSurviveReopen() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordList list = document.AddList(WordListStyle.Numbered);
            WordStyle baseStyle = document.Styles.AddParagraphStyle("ReportBody", "Report Body");
            baseStyle.FontFamily = "Georgia";
            baseStyle.LineSpacing = WordLineSpacing.Multiple(1.5);

            WordStyle step = document.Styles.AddParagraphStyle("ReportStep", "Report Step");
            step.BasedOnStyleId = baseStyle.StyleId;
            step.NextStyleId = baseStyle.StyleId;
            step.Underline = WordUnderlineStyle.Double;
            step.IndentationAfter = OfficeLength.Points(18);
            step.IndentationFirstLine = OfficeLength.Points(9);
            step.ShadingFillColor = OfficeColor.Parse("#FFF2CC");
            step.Borders.BottomStyle = WordBorderStyle.Single;
            step.Borders.BottomSize = OfficeLength.Points(1);
            step.SetList(list, level: 1);

            WordStyle term = document.Styles.AddCharacterStyle("ReportTerm", "Report Term");
            term.Bold = true;
            term.CapsStyle = WordCapsStyle.SmallCaps;
            term.LinkedStyleId = step.StyleId;
            step.LinkedStyleId = term.StyleId;

            WordStyle grid = document.Styles.AddTableStyle("ReportGrid", "Report Grid");
            grid.SetTableBorders(WordBorderStyle.Single, OfficeLength.Points(0.5), OfficeColor.Parse("#9FC5E8"));
            grid.GetTableRegion(WordTableStyleRegion.FirstRow).Bold = true;
            grid.GetTableRegion(WordTableStyleRegion.FirstRow).FillColor = OfficeColor.Parse("#0B5394");

            document.AddParagraph("Check the totals").SetStyle(step)
                .AddText(" carefully").SetCharacterStyle(term);
            document.AddTable(2, 2).SetStyle(grid);
            document.Save();
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        WordStyle saved = reopened.Styles.Get("Report Step")!;
        Assert.Equal("ReportStep", saved.StyleId);
        Assert.Equal(WordStyleType.Paragraph, saved.Type);
        Assert.True(saved.IsCustom);
        Assert.Equal("ReportBody", saved.BasedOnStyleId);
        Assert.Equal("ReportBody", saved.NextStyleId);
        Assert.Equal("ReportTerm", saved.LinkedStyleId);
        Assert.Equal(WordUnderlineStyle.Double, saved.Underline);
        Assert.Equal(OfficeLength.Points(18), saved.IndentationAfter);
        Assert.Equal(OfficeLength.Points(9), saved.IndentationFirstLine);
        Assert.Equal(OfficeColor.Parse("#FFF2CC"), saved.ShadingFillColor);
        Assert.Equal(WordBorderStyle.Single, saved.Borders.BottomStyle);
        Assert.Equal(1, saved.ListLevel);
        Assert.NotNull(saved.List);
        Assert.Equal(WordLineSpacing.Multiple(1.5), reopened.Styles["ReportBody"].LineSpacing);
        Assert.Equal("Georgia", reopened.Styles["ReportBody"].FontFamily);
        Assert.Equal(WordCapsStyle.SmallCaps, reopened.Styles["ReportTerm"].CapsStyle);
        Assert.True(reopened.Styles["ReportGrid"].GetTableRegion(WordTableStyleRegion.FirstRow).Bold);
        Assert.Equal(OfficeColor.Parse("#0B5394"), reopened.Styles["ReportGrid"].GetTableRegion(WordTableStyleRegion.FirstRow).FillColor);

        WordParagraph paragraph = reopened.Paragraphs.First(item => item.Text == "Check the totals");
        Assert.Equal("ReportStep", paragraph.StyleId);
        Assert.Equal("ReportTerm", reopened.Paragraphs.First(item => item.Text == " carefully").CharacterStyleId);
        Assert.Equal("ReportGrid", reopened.Tables[0]._tableProperties!.TableStyle!.Val!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
    }

    [Fact]
    public void DocumentStyles_BelongToTheirDocument() {
        using WordDocument first = WordDocument.Create();
        using WordDocument second = WordDocument.Create();

        WordStyle style = first.Styles.AddParagraphStyle("OnlyInFirst");

        Assert.False(second.Styles.Contains("OnlyInFirst"));
        Assert.Throws<ArgumentException>(() => second.AddParagraph("Elsewhere").SetStyle(style));
        Assert.Throws<ArgumentException>(() => first.Styles.AddParagraphStyle("OnlyInFirst"));
        Assert.Throws<ArgumentException>(() => first.Styles.AddCharacterStyle("Other", "onlyinfirst"));
        Assert.Throws<ArgumentException>(() => first.AddParagraph("Wrong type").SetCharacterStyle(style));
    }

    [Fact]
    public void DocumentStyles_BasingOnAnExistingCycleDoesNotHang() {
        using WordDocument document = WordDocument.Create();
        WordStyle first = document.Styles.AddParagraphStyle("CycleA");
        WordStyle second = document.Styles.AddParagraphStyle("CycleB");
        second.BasedOnStyleId = first.StyleId;
        // A loaded document can already contain a cycle the setter would refuse to create.
        first.Element.BasedOn = new BasedOn { Val = second.StyleId };
        WordStyle third = document.Styles.AddParagraphStyle("CycleC");

        third.BasedOnStyleId = first.StyleId;

        Assert.Equal("CycleA", third.BasedOnStyleId);
    }

    [Fact]
    public void DocumentStyles_ImportRenamesAStyleWhoseNameATargetStyleOfAnotherTypeUses() {
        using WordDocument source = WordDocument.Create();
        source.Styles.AddParagraphStyle("NoteParagraph", "Note").Bold = true;

        using WordDocument target = WordDocument.Create();
        target.Styles.AddCharacterStyle("NoteCharacter", "Note");

        WordStyleImportReport report = target.Styles.Import(source, WordStyleConflictMode.KeepTarget, new[] { "NoteParagraph" });

        Assert.Equal("NoteParagraph_1", report.Renamed["NoteParagraph"]);
        Assert.Equal(WordStyleType.Character, target.Styles.Get("Note")!.Type);
        Assert.Equal("Note 1", target.Styles["NoteParagraph_1"].Name);
        Assert.True(target.Styles["NoteParagraph_1"].Bold);
    }

    [Fact]
    public void DocumentStyles_RemovingAStyleRebasesDependentsAndClearsReferences() {
        using WordDocument document = WordDocument.Create();
        WordStyle parent = document.Styles.AddParagraphStyle("Parent");
        WordStyle middle = document.Styles.AddParagraphStyle("Middle");
        middle.BasedOnStyleId = parent.StyleId;
        WordStyle child = document.Styles.AddParagraphStyle("Child");
        child.BasedOnStyleId = middle.StyleId;
        child.NextStyleId = middle.StyleId;
        WordParagraph paragraph = document.AddParagraph("Uses middle").SetStyle(middle);

        Assert.Throws<ArgumentException>(() => parent.BasedOnStyleId = child.StyleId);
        Assert.True(document.Styles.Remove("Middle"));

        Assert.False(document.Styles.Contains("Middle"));
        Assert.Equal("Parent", child.BasedOnStyleId);
        Assert.Null(child.NextStyleId);
        Assert.Null(paragraph.StyleId);
        Assert.Throws<InvalidOperationException>(() => middle.Bold = true);
        Assert.Throws<InvalidOperationException>(() => document.Styles.Remove("Normal"));
        Assert.False(document.Styles.Remove("Middle"));
    }

    [Fact]
    public void DocumentStyles_ImportFromWordTemplateMatchesLocalizedBuiltInStylesByName() {
        using WordDocument template = WordDocument.Load(Path.Combine(_directoryDocuments, "ExampleTemplate.dotx"),
            new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        using WordDocument document = WordDocument.Create();

        WordStyleImportReport report = document.Styles.Import(template, styleIds: new[] { "Liste", "Titel" });

        // Word's German template names its built-in styles in English, so they map onto this document's own.
        Assert.Contains("Normal", report.Kept);
        Assert.Contains("Title", report.Kept);
        Assert.Contains("Liste", report.Added);
        Assert.Contains("Textbody", report.Added);
        Assert.Contains("Heading", report.Added);
        Assert.False(document.Styles.Contains("Standard"));
        Assert.Equal("Textbody", document.Styles["Liste"].BasedOnStyleId);
        Assert.Equal("Normal", document.Styles["Textbody"].BasedOnStyleId);
        Assert.Throws<KeyNotFoundException>(() => document.Styles.Import(template, styleIds: new[] { "Missing" }));
    }

    [Theory]
    [InlineData(WordStyleConflictMode.KeepTarget)]
    [InlineData(WordStyleConflictMode.Overwrite)]
    [InlineData(WordStyleConflictMode.Rename)]
    public void DocumentStyles_ImportResolvesConflictsAndCarriesDependenciesAndLists(WordStyleConflictMode mode) {
        using WordDocument source = WordDocument.Create();
        WordList sourceList = source.AddList(WordListStyle.Numbered);
        sourceList.AddItem("Source list item");
        WordStyle sourceBase = source.Styles.AddParagraphStyle("BrandBase", "Brand Base");
        sourceBase.FontFamily = "Georgia";
        WordStyle sourceBrand = source.Styles.AddParagraphStyle("Brand");
        sourceBrand.BasedOnStyleId = sourceBase.StyleId;
        sourceBrand.Bold = true;
        sourceBrand.SetList(sourceList);
        WordStyle sourceTable = source.Styles.AddParagraphStyle("Grid", "Source Grid");
        sourceTable.Italic = true;

        using var targetStream = new MemoryStream();
        int targetList;
        using (WordDocument target = WordDocument.Create(targetStream)) {
            WordList bullets = target.AddList(WordListStyle.Bulleted);
            bullets.AddItem("Target list item");
            targetList = bullets.NumberId;
            target.Styles.AddParagraphStyle("Brand").Italic = true;
            // Same identifier, different type: never kept or overwritten.
            target.Styles.AddCharacterStyle("Grid", "Target Grid");

            WordStyleImportReport report = target.Styles.Import(source, mode, new[] { "Brand", "Grid" });

            Assert.Equal(new[] { "BrandBase" }, report.Added);
            Assert.Contains("Normal", report.Kept);
            Assert.Equal("Grid_1", report.Renamed["Grid"]);
            Assert.Equal(WordStyleType.Character, target.Styles["Grid"].Type);
            Assert.True(target.Styles["Grid_1"].Italic);
            switch (mode) {
                case WordStyleConflictMode.KeepTarget:
                    Assert.Contains("Brand", report.Kept);
                    Assert.True(target.Styles["Brand"].Italic);
                    Assert.Null(target.Styles["Brand"].Bold);
                    break;
                case WordStyleConflictMode.Overwrite:
                    Assert.Contains("Brand", report.Replaced);
                    Assert.True(target.Styles["Brand"].Bold);
                    Assert.Null(target.Styles["Brand"].Italic);
                    Assert.Equal("BrandBase", target.Styles["Brand"].BasedOnStyleId);
                    break;
                case WordStyleConflictMode.Rename:
                    Assert.Equal("Brand_1", report.Renamed["Brand"]);
                    Assert.Equal("Brand 1", target.Styles["Brand_1"].Name);
                    Assert.True(target.Styles["Brand_1"].Bold);
                    Assert.True(target.Styles["Brand"].Italic);
                    break;
            }
            target.Save();
        }

        targetStream.Position = 0;
        using WordDocument reopened = WordDocument.Load(targetStream);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        Numbering numbering = reopened._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(numbering.Elements<NumberingInstance>().Count(), numbering.Elements<NumberingInstance>().Select(item => item.NumberID!.Value).Distinct().Count());
        Assert.Equal(numbering.Elements<AbstractNum>().Count(), numbering.Elements<AbstractNum>().Select(item => item.AbstractNumberId!.Value).Distinct().Count());
        string? importedBrand = mode switch {
            WordStyleConflictMode.Overwrite => "Brand",
            WordStyleConflictMode.Rename => "Brand_1",
            _ => null
        };
        if (importedBrand != null) {
            // The imported style numbers with its own copy of the source list, not with the target's list.
            int importedList = reopened.Styles[importedBrand].List!.NumberId;
            Assert.NotEqual(targetList, importedList);
            NumberingInstance instance = Assert.Single(numbering.Elements<NumberingInstance>(), item => item.NumberID!.Value == importedList);
            Assert.Contains(numbering.Elements<AbstractNum>(), item => item.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value);
        }
    }

    [Fact]
    public void DocumentStyles_GetOrAddReturnsTheDocumentsDefinitionOrAddsTheLibraryOne() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordStyle heading = document.Styles.GetOrAdd(WordParagraphStyles.Heading2);
            heading.ThemeColor = WordThemeColor.Accent2;
            Assert.Same(heading.Element, document.Styles.GetOrAdd(WordParagraphStyles.Heading2).Element);
            Assert.Equal(WordStyleType.Table, document.Styles.GetOrAdd(WordTableStyle.GridTable1Light).Type);
            document.AddParagraph("Styled").Style = WordParagraphStyles.Heading2;
            document.Save();
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        Assert.Equal(WordThemeColor.Accent2, reopened.Styles["Heading2"].ThemeColor);
        Assert.Equal(reopened.Theme.GetColor(WordThemeColor.Accent2), reopened.Styles["Heading2"].Color);
    }
}
