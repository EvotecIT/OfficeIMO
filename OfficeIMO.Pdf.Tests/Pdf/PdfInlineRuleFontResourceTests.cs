using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineRuleFontResourceTests {
    private const string Family = "Inline rule metric font";

    [Fact]
    public void GlyphFreeRulesDoNotSelectTheirMetricOnlyNamedFonts() {
        var options = Options();
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => p.Runs(new[] {
            Rule(PdfTextRun.Normal(string.Empty, fontFamily: Family)),
            Rule(PdfTextRun.Bolded(string.Empty, fontFamily: Family))
        })).ToBytes();

        var snapshot = ReadPageFonts(bytes);
        Assert.Empty(snapshot.Uses);
        AssertAllSelectedFontsAreDeclared(snapshot);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Empty(pdf.GetPage(1).Letters);
    }

    [Theory]
    [InlineData(PdfTabLeaderStyle.None, false, false)]
    [InlineData(PdfTabLeaderStyle.Dots, false, false)]
    [InlineData(PdfTabLeaderStyle.None, true, false)]
    [InlineData(PdfTabLeaderStyle.Dots, true, false)]
    [InlineData(PdfTabLeaderStyle.None, false, true)]
    [InlineData(PdfTabLeaderStyle.None, true, true)]
    public void RuleSeparatorAndLeaderGlyphsDeclareTheirActualNamedFont(
        PdfTabLeaderStyle leader, bool artifactCanvas, bool trailingSeparator) {
        var options = Options();
        Assert.True(options.TryResolveNamedFontFace(Family, false, false, out PdfNamedFontFace metricFont));
        PdfTextRun sourceStyle = PdfTextRun.Normal(string.Empty, fontFamily: Family);
        PdfTextRun[] runs = trailingSeparator
            ? new[] { PdfTextRun.Normal("A"), Rule(sourceStyle, maximumWidth: null), PdfTextRun.Normal(" B") }
            : new[] { PdfTextRun.Normal("A"), new PdfTextRun("\t", fontFamily: Family, tabLeader: leader), Rule(sourceStyle), PdfTextRun.Normal("B") };
        var document = PdfDocument.Create(options);
        if (artifactCanvas) document.Canvas(canvas => canvas.Artifact(artifact => artifact.TextBox(runs, 20D, 20D, 100D, 100D)));
        else document.Paragraph(p => p.Runs(runs));
        byte[] bytes = document.ToBytes();

        var snapshot = ReadPageFonts(bytes);
        Assert.Contains(metricFont.ResourceName, snapshot.Uses);
        AssertAllSelectedFontsAreDeclared(snapshot);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("A", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("B", pdf.GetPage(1).Text, StringComparison.Ordinal);
        if (leader == PdfTabLeaderStyle.Dots)
            Assert.Contains(".", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRuleToRuleToTextReopensOnlyWithTheActualTextFont() {
        var options = Options();
        options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
        Assert.True(options.TryResolveNamedFontFace(Family, false, false, out PdfNamedFontFace regular));
        Assert.True(options.TryResolveNamedFontFace(Family, true, false, out PdfNamedFontFace bold));
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => p.Runs(new[] {
            Rule(PdfTextRun.Normal(string.Empty, fontFamily: Family)),
            Rule(PdfTextRun.Bolded(string.Empty, fontFamily: Family)), PdfTextRun.Normal("AB")
        })).ToBytes();

        var snapshot = ReadPageFonts(bytes);
        Assert.True(snapshot.Tagged);
        Assert.NotEmpty(snapshot.Uses);
        Assert.DoesNotContain(regular.ResourceName, snapshot.Uses);
        Assert.DoesNotContain(bold.ResourceName, snapshot.Uses);
        AssertAllSelectedFontsAreDeclared(snapshot);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal("AB", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
    }

    private static PdfTextRun Rule(PdfTextRun style, double? maximumWidth = 30D) =>
        PdfTextRun.Inline(new PdfInlineRule(.5D, PdfColor.Black, maximumWidth: maximumWidth), style);

    private static PdfOptions Options() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, ' ', '.', 'A', 'B');
        return new PdfOptions {
            PageWidth = 140D, PageHeight = 160D, MarginLeft = 20D, MarginRight = 20D,
            MarginTop = 20D, MarginBottom = 20D, DefaultFontSize = 10D, DefaultFont = PdfStandardFont.Helvetica
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(Family, font, bold: font));
    }

    private static void AssertAllSelectedFontsAreDeclared((string[] Uses, HashSet<string> Declared, bool Tagged) snapshot) =>
        Assert.All(snapshot.Uses, font => Assert.True(snapshot.Declared.Contains(font),
            "Page Tf references undeclared font /" + font));

    private static (string[] Uses, HashSet<string> Declared, bool Tagged) ReadPageFonts(byte[] bytes) {
        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        PdfObject Resolve(PdfObject value) => value is PdfReference reference ? objects[reference.ObjectNumber].Value : value;
        PdfDictionary Dictionary(PdfObject value) => Assert.IsType<PdfDictionary>(Resolve(value));
        var dictionaries = objects.Values.Select(value => value.Value).OfType<PdfDictionary>().ToArray();
        PdfDictionary page = Assert.Single(dictionaries, value => value.Get<PdfName>("Type")?.Name == "Page");
        PdfDictionary resources = Dictionary(page.Items["Resources"]);
        var declared = resources.Items.TryGetValue("Font", out PdfObject? fontObject)
            ? new HashSet<string>(Dictionary(fontObject!).Items.Keys, StringComparer.Ordinal)
            : new HashSet<string>(StringComparer.Ordinal);
        PdfStream content = Assert.IsType<PdfStream>(Resolve(page.Items["Contents"]));
        string text = Encoding.ASCII.GetString(OfficeIMO.Pdf.Filters.StreamDecoder.Decode(content.Dictionary, content.Data, objects));
        string[] uses = Regex.Matches(text, @"/(?<font>[^\s/<>\[\]()]+)\s+[-+]?(?:\d+\.?\d*|\.\d+)\s+Tf\b",
            RegexOptions.CultureInvariant).Cast<Match>().Select(match => match.Groups["font"].Value)
            .Distinct(StringComparer.Ordinal).ToArray();
        bool tagged = Assert.Single(dictionaries, value => value.Get<PdfName>("Type")?.Name == "Catalog").Items.ContainsKey("StructTreeRoot");
        return (uses, declared, tagged);
    }
}
