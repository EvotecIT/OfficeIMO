using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

/// <summary>
/// The SDK-free slide text read must return, slide by slide in presentation order, what the Open XML SDK returns
/// for the InnerText of every DrawingML paragraph on the slide.
/// </summary>
public class OpenXmlSlideTextReaderTests {
    public static IEnumerable<object[]> CorpusPresentations() {
        string folder = Path.Combine(AppContext.BaseDirectory, "Documents");
        if (!Directory.Exists(folder)) yield break;
        foreach (string file in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
            yield return new object[] { file.Substring(AppContext.BaseDirectory.Length).TrimStart(Path.DirectorySeparatorChar) };
        }
    }

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void ReadsTheSameSlideTextAsTheSdkForCorpusPresentations(string relativePath) {
        byte[] package = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));

        Assert.Equal(ReadWithSdk(package), ReadWithPackageReader(package));
    }

    [Fact]
    public void ReadsTheBenchmarkDeckInPresentationOrderLikeTheSdk() {
        byte[] package = CreateDeck(12, moveLastSlideFirst: true);

        List<string> actual = ReadWithPackageReader(package);

        Assert.StartsWith("0:Quarterly review 12", actual[0], StringComparison.Ordinal);
        Assert.Equal(ReadWithSdk(package), actual);
    }

    private static List<string> ReadWithPackageReader(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        return OfficeOpenXmlTextReader.ReadSlideParagraphs(stream)
            .Select(paragraph => paragraph.SlideIndex.ToString(CultureInfo.InvariantCulture) + ":" + paragraph.Text)
            .ToList();
    }

    private static List<string> ReadWithSdk(byte[] package) {
        var settings = new OpenSettings {
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
        };
        using var stream = new MemoryStream(package, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false, settings);
        PresentationPart presentation = document.PresentationPart!;
        var paragraphs = new List<string>();
        int index = 0;
        foreach (SlideId slideId in presentation.Presentation!.SlideIdList?.Elements<SlideId>() ?? Enumerable.Empty<SlideId>()) {
            var slide = (SlidePart)presentation.GetPartById(slideId.RelationshipId!.Value!);
            foreach (A.Paragraph paragraph in slide.Slide!.Descendants<A.Paragraph>()) {
                paragraphs.Add(index.ToString(CultureInfo.InvariantCulture) + ":" + paragraph.InnerText);
            }
            index++;
        }
        return paragraphs;
    }

    // The slide shape of the PowerPoint read benchmarks: a title, a bulleted text box, a table, a picture and notes.
    internal static byte[] CreateDeck(int slides, bool moveLastSlideFirst = false) {
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            for (int number = 1; number <= slides; number++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTitle("Quarterly review " + number.ToString(CultureInfo.InvariantCulture));
                PowerPointTextBox body = slide.AddTextBox("Highlights for region " + number.ToString(CultureInfo.InvariantCulture) + " & <partners>");
                for (int item = 1; item <= 5; item++) {
                    body.AddParagraph("Point " + item.ToString(CultureInfo.InvariantCulture) + ": revenue, margin and pipeline commentary.");
                }
                PowerPointTable table = slide.AddTable(3, 3);
                for (int row = 0; row < 3; row++) {
                    for (int column = 0; column < 3; column++) {
                        table.GetRow(row).GetCell(column).Text = (number * 100 + row * 10 + column).ToString(CultureInfo.InvariantCulture);
                    }
                }
                using (var image = new MemoryStream(png, writable: false)) slide.AddImage(image, OfficeImageFormat.Png);
                slide.Notes.Text = "Speaker notes for slide " + number.ToString(CultureInfo.InvariantCulture);
            }
            if (moveLastSlideFirst) presentation.MoveSlide(slides - 1, 0);
            presentation.Save();
        }
        return stream.ToArray();
    }
}
