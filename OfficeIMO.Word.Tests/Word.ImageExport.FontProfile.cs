using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_BatchAndSnapshotsPreserveAllPagesWithConfiguredWideText(bool useShaper) {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        WordSection section = document.Sections[0];
        section.PageSettings.Width = (UInt32Value)5000U;
        section.PageSettings.Height = (UInt32Value)3000U;
        section.SetMargins(WordMargin.Narrow);
        WordParagraph paragraph = document.AddParagraph(string.Join(" ", Enumerable.Range(1, 40)
            .Select(index => "T" + index.ToString("00", CultureInfo.InvariantCulture))));
        paragraph.SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        paragraph.FontSizePoints = 11D;
        paragraph.AvoidWidowAndOrphan = false;
        var options = CreateWideTextOptions(useShaper);

        var images = document.ExportImages(OfficeImageExportFormat.Svg, options);
        var snapshots = document.CreateVisualSnapshots(options);

        Assert.True(images.Count > 1);
        Assert.Equal(images.Count, snapshots.Count);
        Assert.Equal(Enumerable.Range(0, snapshots.Count), snapshots.Select(snapshot => snapshot.PageIndex));
        AssertPaintedTokens(images.Select(image => Encoding.UTF8.GetString(image.Bytes)), "T", 40);
        AssertPaintedTokens(snapshots.Select(snapshot => OfficeDrawingSvgExporter.ToSvg(snapshot.Drawing)), "T", 40);
    }

    private static WordImageExportOptions CreateWideTextOptions(bool useShaper) {
        var options = new WordImageExportOptions { TextShapingLanguage = "en-GB" };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithAdvance(useShaper ? (ushort)500 : (ushort)1500,
                Enumerable.Range(32, 95).ToArray()));
        if (useShaper) options.TextShapingProvider = new WideAdvanceProvider();
        return options;
    }

    private sealed class WideAdvanceProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Assert.Equal("en-GB", request.Language);
            var glyphs = new List<OfficeShapedGlyph>();
            int textIndex = 0;
            foreach (string element in OfficeTextElements.Enumerate(request.Text)) {
                glyphs.Add(new OfficeShapedGlyph(1, element, textIndex, advanceWidth: 1500));
                textIndex += element.Length;
            }
            return new OfficeTextShapingResult(glyphs);
        }
    }
}
