using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("data:image/svg+xml;charset=bogus,%3Csvg%3E%3C/svg%3E")]
    [InlineData("data:image/svg+xml;charset=utf-8,%FF")]
    public void HtmlToWord_ImageSelection_ContinuesPastTextDataCandidateWithCharsetFailure(string candidate) {
        string path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
        string fallback = "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(path));
        string html = $"<img data-src=\"{candidate}\" src=\"{fallback}\" alt=\"Logo\">";

        HtmlToWordResult conversion = HtmlConversionDocument.Parse(html).ToWordDocumentResult(new HtmlToWordOptions());

        Assert.Single(conversion.Value.Images);
        Assert.Empty(conversion.Report.Diagnostics);
    }
}
