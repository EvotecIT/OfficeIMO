using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHtmlPositionedTextTests {
    [Fact]
    public void PositionedTextPreservesNativeBaselinesSizesAndColors() {
        var source = PdfDocument.Create().Paragraph(p => p.Text("Original text"));
        var stamped = PdfDocument.Load(source.ToBytes()).Stamp.TextWatermark("Review", new PdfTextStampOptions {
            X = 210, Y = 430, FontSize = 22, RotationDegrees = 0, Color = PdfColor.FromRgb(12, 34, 56)
        });
        var document = PdfDocument.Load(stamped.ToBytes()).Read();
        var result = document.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
        int start = result.Value.IndexOf("<svg class=\"pdf-native-text\"", StringComparison.Ordinal);
        int end = result.Value.IndexOf("</svg>", start, StringComparison.Ordinal) + 6;
        var svg = XElement.Parse(result.Value.Substring(start, end - start));
        XNamespace ns = "http://www.w3.org/2000/svg";
        var text = Assert.Single(svg.Descendants(ns + "text"), element => element.Value == "Review");
        Assert.Equal("210", (string?)text.Attribute("x"));
        Assert.Equal((document.Pages[0].Height - 430).ToString("0.###", CultureInfo.InvariantCulture), (string?)text.Attribute("y"));
        Assert.Equal("22", (string?)text.Attribute("font-size"));
        Assert.Equal("rgb(12,34,56)", (string?)text.Attribute("fill"));
        Assert.Equal(document.Pages[0].TextBlocks.SelectMany(b => b.Spans).Distinct().Count(s => s.IsVisible && s.Text.Length > 0),
            svg.Descendants(ns + "text").Count());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "PositionedFontSubstitution");
    }
}
