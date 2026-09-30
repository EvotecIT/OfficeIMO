using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfEffectiveNumberingRegressionTests {
    [Fact]
    public void Conversion_Reports_Expose_Unsupported_Markers_And_Markdown_Formatting_Loss() {
        RtfDocument document = CreateNestedList();
        document.ListDefinitions[0].Levels[0].NumberFormat = 10;
        document.AddParagraph("Numbered").SetList(3);
        Assert.False(new RtfListNumbering(document).Next(document.Paragraphs[0])!.IsNumberFormatSupported);
        Assert.Contains(document.ToHtmlResult().RtfReport.Diagnostics, item => item.Code == "RtfHtmlListNumberFormatFlattened" && item.Action == RtfConversionAction.Flattened);
        Assert.Contains(document.ToPdfDocumentResult().Warnings, item => item.Code == "ListNumberFormatFlattened" && item.Details["RtfAction"] == nameof(RtfConversionAction.Flattened));
        Assert.Contains(document.ToMarkdownResult().Report.Diagnostics, item => item.Code == "RTFMD017" && item.Action == RtfConversionAction.Flattened);
        document.ListDefinitions[0].Levels[0].NumberFormat = 1;
        Assert.Contains(document.ToMarkdownResult().Report.Diagnostics, item => item.Code == "RTFMD017" && item.Action == RtfConversionAction.Flattened);
        Assert.DoesNotContain(document.ToHtmlResult().RtfReport.Diagnostics, item => item.Code == "RtfHtmlListNumberFormatFlattened");
    }

    [Fact]
    public void Explicit_List_Removal_Overrides_Style_Numbering_In_Rtf_And_Html() {
        RtfDocument document = CreateNestedList();
        RtfStyle style = document.AddStyle(1, "Numbered");
        style.ListId = 3;
        style.ListLevel = 0;
        RtfParagraph paragraph = document.AddParagraph("Unnumbered");
        paragraph.StyleId = 1;
        paragraph.ListId = 0;
        Assert.Null(document.ResolveListFormatting(paragraph));
        Assert.Equal(RtfListKind.None, document.ResolveParagraphFormatting(paragraph).ListKind);
        RtfDocument native = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(0, native.Paragraphs[0].ListId);
        Assert.Null(native.ResolveListFormatting(native.Paragraphs[0]));
        RtfDocument html = HtmlConversionDocument.Parse(document.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true })).ToRtfDocument();
        Assert.Equal(0, html.Paragraphs[0].ListId);
        Assert.Null(html.ResolveListFormatting(html.Paragraphs[0]));
    }

    [Fact]
    public void Format_And_Start_Overrides_Write_Start_Inside_Replacement_Level() {
        RtfDocument document = CreateNestedList();
        RtfListLevelOverride item = document.ListOverrides[0].AddLevelOverride();
        item.OverrideFormat = true;
        item.OverrideStartAt = true;
        item.StartAt = 11;
        item.Formatting = new RtfListLevel(0) { NumberFormat = 2, Text = "%1)", StartAt = 99 };
        document.AddParagraph("Changed").SetList(3);
        string source = document.ToRtf();
        Assert.Contains(@"\listoverridestartat1{\listlevel", source, StringComparison.Ordinal);
        RtfDocument native = RtfDocument.Read(source).Document;
        Assert.Equal(11, native.ResolveListFormatting(native.Paragraphs[0])!.Level.StartAt);
        Assert.Equal("xi)", new RtfListNumbering(native).Next(native.Paragraphs[0])!.Text);
    }

    [Fact]
    public void Decimal_Zero_And_Legacy_Roman_Markers_Use_Their_Authored_Formats() {
        RtfDocument document = CreateNestedList();
        document.ListDefinitions[0].Levels[0].NumberFormat = 22;
        RtfParagraph paragraph = document.AddParagraph("Padded").SetList(3);
        Assert.Equal("04.", new RtfListNumbering(document).Next(paragraph)!.Text);
        Assert.Contains(">04.\t</span>", document.ToHtml(), StringComparison.Ordinal);
        var legacy = new RtfParagraph();
        legacy.LegacyNumbering.Enabled = true;
        legacy.LegacyNumbering.NumberStyle = RtfLegacyNumberingStyle.LowerRoman;
        legacy.LegacyNumbering.StartAt = 9;
        legacy.LegacyNumbering.TextBefore = "(";
        legacy.LegacyNumbering.TextAfter = ")";
        Assert.Equal("(ix)", new RtfListNumbering(document).Next(legacy)!.Text);
    }

    private static RtfDocument CreateNestedList() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100);
        RtfListLevel parent = definition.AddLevel();
        parent.NumberFormat = 1;
        parent.StartAt = 4;
        parent.Text = "%1.";
        RtfListLevel child = definition.AddLevel();
        child.NumberFormat = 4;
        child.Text = "%1.%2.";
        child.FollowCharacter = RtfListLevelFollowCharacter.Space;
        definition.AddLevel().StartAt = 7;
        document.AddListOverride(3, 100);
        return document;
    }

    [Fact]
    public void Sparse_Overrides_Resolve_By_Target_Level_And_Respect_Their_Start_Flag() {
        RtfDocument document = CreateNestedList();
        RtfListLevelOverride item = document.ListOverrides[0].AddLevelOverride();
        item.LevelIndex = 2;
        item.StartAt = 9;
        item.OverrideStartAt = true;
        RtfParagraph paragraph = document.AddParagraph("Nested").SetList(3, 2);
        Assert.Equal(9, document.ResolveListFormatting(paragraph)!.Level.StartAt);
        item.OverrideStartAt = false;
        Assert.Equal(7, document.ResolveListFormatting(paragraph)!.Level.StartAt);
        item.OverrideStartAt = true;
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(9, reopened.ListOverrides[0].LevelOverrides.Count);
        Assert.Equal(9, reopened.ResolveListFormatting(reopened.Paragraphs[0])!.Level.StartAt);
        Assert.Equal(4, reopened.ResolveListFormatting(new RtfParagraph { ListId = 3, ListLevel = 0 })!.Level.StartAt);
    }

    [Fact]
    public void Shared_Counters_Use_Ancestor_Formats_And_Restart_Children_Only_When_Required() {
        RtfDocument document = CreateNestedList();
        document.AddParagraph("Parent").SetList(3, 0);
        document.AddParagraph("Child").SetList(3, 1);
        document.AddParagraph("Child").SetList(3, 1);
        document.AddParagraph("Parent").SetList(3, 0);
        document.AddParagraph("Child").SetList(3, 1);
        var numbering = new RtfListNumbering(document);
        Assert.Equal(new[] { "IV.", "IV.a.", "IV.b.", "V.", "V.a." }, document.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text));
        document.ListDefinitions[0].Levels[1].NoRestart = true;
        numbering = new RtfListNumbering(document);
        Assert.Equal("V.c.", document.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text).ToArray().Last());
        document.ListDefinitions[0].Levels[1].LegalNumbering = true;
        numbering = new RtfListNumbering(document);
        Assert.Equal("4.a.", document.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text).ToArray()[1]);
    }

    [Fact]
    public void Native_Marker_Placeholders_And_Offset_Bytes_Survive_Normalized_Writing() {
        const string input = @"{\rtf1\ansi{\*\listtable{\list{\listlevel\levelnfc0{\leveltext\'06\'00.\'01.\'02.;}{\levelnumbers\'01\'03\'05;}}\listid100}}}";
        RtfDocument document = RtfDocument.Read(input).Document;
        RtfListLevel level = Assert.Single(Assert.Single(document.ListDefinitions).Levels);
        Assert.Equal("%1.%2.%3.", level.Text);
        Assert.Equal("\u0001\u0003\u0005", level.Numbers);
        string output = document.ToRtf();
        Assert.Contains(@"{\leveltext\'06\'00.\'01.\'02.;}", output, StringComparison.Ordinal);
        Assert.Contains(@"{\levelnumbers\'01\'03\'05;}", output, StringComparison.Ordinal);
        Assert.Equal(level.Text, RtfDocument.Read(output).Document.ListDefinitions[0].Levels[0].Text);
    }

    [Fact]
    public void Html_Renders_Proper_Nesting_And_Preserves_External_Start_Type_And_Value_Resets() {
        RtfDocument external = HtmlConversionDocument.Parse("<ol start='5' type='i'><li>First</li><li value='9'>Reset</li><li>Next</li></ol>").ToRtfDocument();
        var numbering = new RtfListNumbering(external);
        Assert.Equal(new[] { "v.", "ix.", "x." }, external.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text));
        RtfDocument document = CreateNestedList();
        document.AddParagraph("Parent").SetList(3, 0);
        document.AddParagraph("Child").SetList(3, 1);
        document.AddParagraph("Second").SetList(3, 0);
        document.AddListOverride(4, 100).AddLevelOverride().OverrideStartAt = true;
        document.AddParagraph("Other").SetList(4, 0);
        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, IncludeRoundTripMetadata = true });
        AngleSharp.Html.Dom.IHtmlDocument dom = new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html);
        Assert.Single(dom.QuerySelectorAll("ol > li > ol > li"));
        Assert.Equal(2, dom.QuerySelectorAll("body > ol").Length);
        RtfDocument reopened = HtmlConversionDocument.Parse(html).ToRtfDocument();
        Assert.Equal(new[] { "Parent", "Child", "Second", "Other" }, reopened.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        numbering = new RtfListNumbering(reopened);
        Assert.Equal(new[] { "IV.", "IV.a.", "V.", "IV." }, reopened.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text));
    }

    [Fact]
    public void Formatting_Overrides_Share_Counters_And_First_Start_Overrides_Reset_The_Shared_Sequence() {
        RtfDocument document = CreateNestedList();
        document.AddListOverride(4, 100);
        RtfListLevelOverride format = document.AddListOverride(5, 100).AddLevelOverride();
        format.OverrideFormat = true;
        format.Formatting = new RtfListLevel(0) { NumberFormat = 2, Text = "(%1)" };
        RtfListLevelOverride start = document.AddListOverride(6, 100).AddLevelOverride();
        start.OverrideStartAt = true;
        start.StartAt = 4;
        foreach (int listId in new[] { 3, 4, 5, 6, 3, 6 }) document.AddParagraph("Item").SetList(listId);
        foreach (RtfDocument value in new[] { document, RtfDocument.Read(document.ToRtf()).Document }) {
            var numbering = new RtfListNumbering(value);
            Assert.Equal(new[] { "IV.", "V.", "(vi)", "IV.", "V.", "VI." }, value.Paragraphs.Select(paragraph => numbering.Next(paragraph)!.Text));
        }
        using WordDocument word = document.ToWordDocument();
        RtfDocument imported = word.ToRtfDocument();
        var wordNumbering = new RtfListNumbering(imported);
        Assert.Equal(new[] { "IV.", "V.", "(vi)", "IV.", "V.", "VI." }, imported.Paragraphs.Select(paragraph => wordNumbering.Next(paragraph)!.Text));
    }

    [Fact]
    public void Pdf_And_Markdown_Continue_The_Same_Instance_Across_Intervening_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        document.AddListDefinition(100).AddLevel().StartAt = 3;
        document.AddListOverride(3, 100);
        document.AddListOverride(4, 100).AddLevelOverride().OverrideStartAt = true;
        document.ListOverrides[1].LevelOverrides[0].StartAt = 2;
        document.AddParagraph("First").SetList(3);
        document.AddParagraph("Between");
        document.AddParagraph("Continued").SetList(3);
        document.AddParagraph("Other").SetList(4);
        MarkdownDoc markdown = document.ToMarkdownDocument();
        Assert.Equal(new[] { 3, 4, 2 }, markdown.Blocks.OfType<OrderedListBlock>().Select(list => list.Start));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("3. First", text, StringComparison.Ordinal);
        Assert.Contains("4. Continued", text, StringComparison.Ordinal);
        Assert.Contains("2. Other", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_Formatting_Overrides_And_Style_Numbering_Are_Schema_Valid_And_Survive_Reopen() {
        RtfDocument document = CreateNestedList();
        RtfListLevelOverride item = document.ListOverrides[0].AddLevelOverride();
        item.LevelIndex = 2;
        item.OverrideFormat = true;
        item.Formatting = new RtfListLevel(2) { NumberFormat = 2, Text = "%3)", StartAt = 99, FollowCharacter = RtfListLevelFollowCharacter.Space };
        RtfStyle style = document.AddStyle(1, "Numbered");
        style.ListId = 3;
        style.ListLevel = 2;
        document.AddParagraph("Styled").StyleId = 1;
        using WordDocument word = document.ToWordDocument();
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word._wordprocessingDocument));
        using var bytes = new MemoryStream();
        word.Save(bytes);
        bytes.Position = 0;
        using WordDocument opened = WordDocument.Load(bytes);
        RtfDocument reopened = opened.ToRtfDocument();
        RtfListFormatting effective = reopened.ResolveListFormatting(reopened.Paragraphs[0])!;
        Assert.Equal(2, effective.Level.NumberFormat);
        Assert.Equal("%3)", effective.Level.Text);
        Assert.Equal(7, effective.Level.StartAt);
        RtfListNumbering numbering = new RtfListNumbering(reopened);
        Assert.Equal("vii)", numbering.Next(reopened.Paragraphs[0])!.Text);
        RtfDocument clone = document.Clone();
        clone.ListOverrides[0].LevelOverrides[0].Formatting!.Text = "Changed";
        Assert.Equal("%3)", item.Formatting.Text);
    }
}
