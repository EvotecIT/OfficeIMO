using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlFontOpticalSizeTests {
    [Theory]
    [InlineData(16)]
    [InlineData(24)]
    [InlineData(48)]
    public void HtmlPdfUsesAuthoredCssSizeForOpticalOutlines(int size) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf"));
        var source = HtmlConversionDocument.Parse("<style>@font-face{font-family:Probe;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(data) + "');font-weight:700}p{font-family:Probe;font-weight:700;font-size:"
            + size + "px}</style><p>Variable OfficeIMO</p>");
        var reference = new HtmlToPdfOptions();
        reference.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["opsz"] = size };
        Assert.Equal(source.ToPdfBytes(reference), source.ToPdfBytes(new HtmlToPdfOptions()));
    }
}
