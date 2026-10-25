using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfContinuousRunningStoryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ContinuousDefaultHeaderKeepsCurrentPageUntilRealBreak(bool replaceHeader, bool bodyPageBreak) {
        // Native Word DOCX and reopened RTF controls independently establish
        // one/one/two pages and headers A / A / A then B for these transitions.
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("S1INTRO");
        first.AddParagraph("S1END");
        RtfSection second = source.AddSection(RtfSectionBreakKind.Continuous);
        if (replaceHeader) second.AddHeader().AddParagraph("HEADERB");
        second.AddParagraph("S2INTRO");
        second.AddParagraph("S2BEFOREBREAK");
        second.AddParagraph("S2FOLLOWING").PageBreakBefore = bodyPageBreak;
        second.AddParagraph("S2END");

        Converted converted = Convert(source, $"rtf-continuous-{replaceHeader}-{bodyPageBreak}");

        Assert.Equal(bodyPageBreak ? 2 : 1, converted.Pages.Length);
        Assert.Equal(1, Occurrences(converted.Pages[0], "HEADERA"));
        Assert.DoesNotContain("HEADERB", converted.Pages[0], StringComparison.Ordinal);
        if (bodyPageBreak) {
            Assert.Equal(1, Occurrences(converted.Pages[1], "HEADERB"));
            Assert.DoesNotContain("HEADERA", converted.Pages[1], StringComparison.Ordinal);
        }
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        string[] body = { "S1INTRO", "S1END", "S2INTRO", "S2BEFOREBREAK", "S2FOLLOWING", "S2END" };
        for (int index = 0; index < body.Length; index++) {
            int expectedPage = bodyPageBreak && index >= 4 ? 1 : 0;
            Assert.Equal(1, Occurrences(converted.Pages[expectedPage], body[index]));
        }
        AssertBodyOnce(converted.Pages, body);
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void ContinuousHeaderRegistersItsSourceOnlyFontWithoutChangingBodyTypography() {
        RtfDocument source = CreateDocument(out RtfSection first);
        source.Settings.SetDefaultFont(source.AddFont("Arial"));
        int headerFont = source.AddFont("Courier New");
        first.AddParagraph("FIRSTBODY").Runs[0].FontSize = 10;
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        RtfRun header = changed.AddHeader().AddParagraph().AddText("COURIERHEADER");
        header.FontId = headerFont;
        header.FontSize = 14;
        changed.AddParagraph("BEFOREBREAK").Runs[0].FontSize = 10;
        RtfParagraph following = changed.AddParagraph("FOLLOWINGPAGE");
        following.PageBreakBefore = true;
        following.Runs[0].FontSize = 10;

        Converted converted = Convert(source, "rtf-continuous-source-only-font");

        Assert.Equal(2, converted.Pages.Length);
        Assert.DoesNotContain("COURIERHEADER", converted.Pages[0], StringComparison.Ordinal);
        Assert.Equal(1, Occurrences(converted.Pages[1], "COURIERHEADER"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(converted.Bytes);
        var paintedHeader = Assert.Single(independent.GetPage(2).GetWords(), word => word.Text == "COURIERHEADER");
        Assert.All(paintedHeader.Letters, letter => {
            Assert.Contains("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(14D, letter.PointSize, 2);
        });
        var body = Assert.Single(independent.GetPage(2).GetWords(), word => word.Text == "FOLLOWINGPAGE");
        Assert.All(body.Letters, letter => {
            Assert.DoesNotContain("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(10D, letter.PointSize, 2);
        });
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "BEFOREBREAK", "FOLLOWINGPAGE");
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void ContinuousHeaderReplacementWorksInsideALaterPageStartingSection() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("PREFIXBODY");
        RtfSection second = source.AddSection(RtfSectionBreakKind.NextPage);
        second.AddHeader().AddParagraph("HEADERC");
        second.AddParagraph("GROUPSTART");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERD");
        continuous.AddParagraph("CONTINUOUSBODY");
        continuous.AddParagraph("FOLLOWINGPAGE").PageBreakBefore = true;

        Converted converted = Convert(source, "rtf-continuous-nested-page-group");

        Assert.Equal(3, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERC", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERD", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("HEADERD", converted.Pages[2], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERC", converted.Pages[2], StringComparison.Ordinal);
        Assert.Contains("CONTINUOUSBODY", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("FOLLOWINGPAGE", converted.Pages[2], StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "PREFIXBODY", "GROUPSTART", "CONTINUOUSBODY", "FOLLOWINGPAGE");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void ContinuousHeaderReplacementSurvivesAnInheritedSuccessor() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("FIRSTBODY");
        RtfSection changed = source.AddSection(RtfSectionBreakKind.Continuous);
        changed.AddHeader().AddParagraph("HEADERB");
        changed.AddParagraph("CHANGEDBODY");
        RtfSection inherited = source.AddSection(RtfSectionBreakKind.Continuous);
        inherited.AddParagraph("INHERITEDBODY");
        inherited.AddParagraph("FOLLOWINGPAGE").PageBreakBefore = true;

        Converted converted = Convert(source, "rtf-continuous-inherited-successor");

        Assert.Equal(2, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERB", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERA", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("INHERITEDBODY", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("FOLLOWINGPAGE", converted.Pages[1], StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CHANGEDBODY", "INHERITEDBODY", "FOLLOWINGPAGE");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
        Assert.DoesNotContain(converted.Result.Warnings, warning => warning.Code == "ContinuousSectionPageSettingsFlattened");
    }

    [Fact]
    public void PageStartingSectionReplacesUnusedContinuousHeaderPresentation() {
        RtfDocument source = CreateDocument(out RtfSection first);
        first.AddParagraph("FIRSTBODY");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERB");
        continuous.AddParagraph("CONTINUOUSBODY");
        RtfSection explicitPage = source.AddSection(RtfSectionBreakKind.NextPage);
        explicitPage.AddHeader().AddParagraph("HEADERC");
        explicitPage.AddParagraph("EXPLICITPAGEBODY");

        Converted converted = Convert(source, "rtf-continuous-explicit-page-owner");

        Assert.Equal(2, converted.Pages.Length);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("CONTINUOUSBODY", converted.Pages[0], StringComparison.Ordinal);
        Assert.Contains("HEADERC", converted.Pages[1], StringComparison.Ordinal);
        Assert.Contains("EXPLICITPAGEBODY", converted.Pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", string.Concat(converted.Pages), StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CONTINUOUSBODY", "EXPLICITPAGEBODY");
        Assert.All(converted.Pages, page => Assert.Equal(1, Occurrences(page, "FOOTERFIXED")));
    }

    [Theory]
    [InlineData("footer", "Footers")]
    [InlineData("first", "StorySelection")]
    [InlineData("note", "Headers")]
    public void UnsupportedContinuousStoryReplacementRetainsItsWarning(string unsupported, string expectedProperty) {
        RtfDocument source = CreateDocument(out RtfSection first);
        RtfParagraph firstBody = first.AddParagraph("FIRSTBODY");
        RtfSection continuous = source.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddHeader().AddParagraph("HEADERB");
        continuous.AddParagraph("CONTINUOUSBODY");
        if (unsupported == "footer") continuous.AddFooter().AddParagraph("REPLACEMENTFOOTER");
        else if (unsupported == "first") {
            continuous.PageSetup.DifferentFirstPageHeaderFooter = true;
            continuous.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("FIRSTVARIANT");
        } else firstBody.AddFootnote("1", "NOTEBODY");

        Converted converted = Convert(source, "rtf-continuous-unsupported-" + unsupported);

        PdfCore.PdfConversionWarning warning = Assert.Single(converted.Result.Warnings,
            item => item.Code == "ContinuousSectionPageSettingsFlattened");
        Assert.Equal(nameof(RtfConversionAction.Flattened), warning.Details["RtfAction"]);
        Assert.Contains(expectedProperty, warning.Details["Properties"], StringComparison.Ordinal);
        Assert.Contains("HEADERA", converted.Pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("HEADERB", string.Concat(converted.Pages), StringComparison.Ordinal);
        Assert.DoesNotContain("REPLACEMENTFOOTER", string.Concat(converted.Pages), StringComparison.Ordinal);
        Assert.DoesNotContain("FIRSTVARIANT", string.Concat(converted.Pages), StringComparison.Ordinal);
        AssertBodyOnce(converted.Pages, "FIRSTBODY", "CONTINUOUSBODY");
        if (unsupported == "note") Assert.Equal(1, Occurrences(string.Concat(converted.Pages), "NOTEBODY"));
    }

    private static RtfDocument CreateDocument(out RtfSection first) {
        RtfDocument source = RtfDocument.Create();
        source.PageSetup.SetPaperSize(4800, 6400);
        source.PageSetup.SetMargins(leftTwips: 480, rightTwips: 480, topTwips: 640, bottomTwips: 640);
        source.PageSetup.HeaderDistanceTwips = 360;
        source.PageSetup.FooterDistanceTwips = 360;
        first = source.AddSection();
        first.AddHeader().AddParagraph("HEADERA");
        first.AddFooter().AddParagraph("FOOTERFIXED");
        return source;
    }

    private static Converted Convert(RtfDocument source, string artifactName) {
        string before = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        PdfCore.PdfDocumentConversionResult result = source.ToPdfDocumentResult(new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        byte[] bytes = result.ToBytes();
        Assert.Equal(before, source.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        string[] pages = PdfCore.PdfReadDocument.Open(bytes,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).Pages
            .Select(page => new string(page.ExtractText().Where(character => !char.IsWhiteSpace(character)).ToArray())).ToArray();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(independent.NumberOfPages, pages.Length);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } directory) {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, artifactName + ".pdf"), bytes);
        }
        return new Converted(result, pages, bytes);
    }

    private static void AssertBodyOnce(string[] pages, params string[] markers) {
        string text = string.Concat(pages);
        int previous = -1;
        foreach (string marker in markers) {
            Assert.Equal(1, Occurrences(text, marker));
            int position = text.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(position > previous, "Body markers must retain their source order.");
            previous = position;
        }
    }

    private static int Occurrences(string text, string marker) {
        int count = 0;
        for (int index = 0; (index = text.IndexOf(marker, index, StringComparison.Ordinal)) >= 0; index += marker.Length) count++;
        return count;
    }

    private sealed class Converted {
        internal Converted(PdfCore.PdfDocumentConversionResult result, string[] pages, byte[] bytes) {
            Result = result;
            Pages = pages;
            Bytes = bytes;
        }
        internal PdfCore.PdfDocumentConversionResult Result { get; }
        internal string[] Pages { get; }
        internal byte[] Bytes { get; }
    }
}
