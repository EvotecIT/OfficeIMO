using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfMergeSectionRegressionTests {
    [Fact]
    public void Word_Section_Column_Settings_Clear_After_A_Multicolumn_Section_And_Rich_Styles_Are_Schema_Valid() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle style = document.AddStyle(1, "Rich");
        style.Bold = true;
        style.FontSize = 14;
        style.HighlightColorIndex = document.AddColor(255, 0, 0);
        style.ParagraphAlignment = RtfTextAlignment.Center;
        style.Priority = 3;
        style.QuickFormat = true;
        style.Locked = true;
        style.SemiHidden = true;
        style.UnhideWhenUsed = true;
        RtfSection first = document.AddSection();
        first.ColumnCount = 2;
        first.ColumnSpaceTwips = 720;
        first.AddParagraph("Columns").StyleId = 1;
        document.AddSection().AddParagraph("Default columns");
        using WordDocument word = document.ToWordDocument();
        Assert.Equal(2, word.Sections[0].ColumnCount);
        Assert.Null(word.Sections[1].ColumnCount);
        Assert.Null(word.Sections[1].ColumnsSpace);
        string[] errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word._wordprocessingDocument).Select(item => item.Description).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    [Fact]
    public void PreserveSections_Imports_Independent_Stories_Page_Setup_And_Source_Defaults() {
        RtfDocument destination = RtfDocument.Create();
        destination.PageSetup.Landscape = true;
        destination.PageSetup.MarginLeftTwips = 900;
        destination.AddHeader().AddParagraph("Destination header");
        destination.AddParagraph("Destination body");
        RtfDocument source = RtfDocument.Create();
        int red = source.AddColor(255, 0, 0);
        source.AddHeader().AddParagraph("Source header").Runs[0].ForegroundColorIndex = red;
        RtfSection first = source.AddSection();
        first.AddParagraph("Source first");
        RtfSection second = source.AddSection();
        second.PageSetup.PaperWidthTwips = 11000;
        second.ColumnCount = 2;
        second.AddParagraph("Source second");

        destination.AppendDocument(source, new RtfDocumentMergeOptions { PreserveSections = true }).Report.RequireNoLoss();
        Assert.Contains(destination.ToRtfResult().Report.Diagnostics, item => item.Code == "RtfNormalizationOrientationDefaultsMaterialized" && item.Action == RtfConversionAction.Flattened);
        Assert.Equal(3, destination.Sections.Count);
        RtfSection imported = destination.Sections[1];
        Assert.False(imported.PageSetup.Landscape);
        Assert.Equal(1800, imported.PageSetup.MarginLeftTwips);
        Assert.Equal(11000, destination.Sections[2].PageSetup.PaperWidthTwips);
        Assert.Equal(2, destination.Sections[2].ColumnCount);
        Assert.Equal("Source header", destination.GetEffectiveHeaderFooters(imported).Single(item => item.Kind == RtfHeaderFooterKind.Header).ToPlainText());
        foreach (RtfDocument value in new[] { destination, RtfDocument.Read(destination.ToRtf()).Document }) {
            Assert.Equal(new[] { "Destination body", "Source first", "Source second" }, value.Paragraphs.Select(item => item.ToPlainText()));
            Assert.False(value.GetEffectivePageSetup(value.Sections[1]).Landscape);
            Assert.Equal(1800, value.Sections[1].PageSetup.MarginLeftTwips);
        }
        using WordDocument word = destination.ToWordDocument();
        string[] errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word._wordprocessingDocument).Select(item => item.Description + " " + item.Node?.OuterXml).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
        Assert.Equal(OfficePageOrientation.Portrait, word.Sections[1].PageOrientation);
        Assert.Equal("Source header", string.Concat(word.Sections[1].Header.Default!.Paragraphs.Select(item => item.Text)));
        imported.AddParagraph("Attached after merge");
        Assert.Contains(destination.Paragraphs, item => item.ToPlainText() == "Attached after merge");
        Assert.DoesNotContain(source.Paragraphs, item => item.ToPlainText() == "Attached after merge");
    }

    [Fact]
    public void Explicit_Page_Setup_Resets_Override_Inheritance_And_Survive_Native_And_Html_Writing() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.Landscape = true;
        document.PageSetup.DifferentFirstPageHeaderFooter = true;
        document.PageSetup.RtlGutter = true;
        RtfSection inherited = document.AddSection();
        inherited.AddParagraph("Inherited");
        RtfSection explicitReset = document.AddSection();
        explicitReset.PageSetup.Landscape = false;
        explicitReset.PageSetup.DifferentFirstPageHeaderFooter = false;
        explicitReset.PageSetup.RtlGutter = false;
        explicitReset.AddParagraph("Reset");
        foreach (RtfDocument value in new[] { document, RtfDocument.Read(document.ToRtf()).Document,
            HtmlConversionDocument.Parse(document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile())).ToRtfDocument() }) {
            Assert.True(value.GetEffectivePageSetup(value.Sections[0]).Landscape, value.ToRtf());
            Assert.True(value.GetEffectivePageSetup(value.Sections[0]).DifferentFirstPageHeaderFooter);
            Assert.True(value.GetEffectivePageSetup(value.Sections[0]).RtlGutter);
            RtfPageSetup reset = value.GetEffectivePageSetup(value.Sections[1]);
            Assert.False(reset.Landscape);
            Assert.False(reset.DifferentFirstPageHeaderFooter);
            Assert.False(reset.RtlGutter);
        }
        using WordDocument word = document.ToWordDocument();
        Assert.Equal(OfficePageOrientation.Landscape, word.Sections[0].PageOrientation);
        Assert.Equal(OfficePageOrientation.Portrait, word.Sections[1].PageOrientation);
        Assert.False(word.Sections[1].DifferentFirstPage);
        Assert.False(word.Sections[1].RtlGutter);
    }
}
