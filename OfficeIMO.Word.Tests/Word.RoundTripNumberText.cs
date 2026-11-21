using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordRoundTripNumberTextTests {
    // VML dimensions use the shortest round-trip digits; .NET Framework's "R" wrote this width with 17 digits.
    [Fact]
    public void VmlImageDimensionsUseTheShortestRoundTripDigits() {
        using var document = WordDocument.Create();
        using (FileStream image = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "EvotecLogo.png"))) {
            document.AddParagraph().InsertImageVml(image, "EvotecLogo.png", widthPoints: 0.4694715627858908D, heightPoints: 10D);
        }

        Assert.Contains("width:0.4694715627858908pt", document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
    }
}
