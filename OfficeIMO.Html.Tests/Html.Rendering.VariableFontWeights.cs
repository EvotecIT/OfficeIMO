using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlVariableFontWeightTests {
    [Theory]
    [InlineData(400, 400)]
    [InlineData(600, 600)]
    [InlineData(700, 700)]
    [InlineData(1, 100)]
    [InlineData(1000, 1000)]
    public void CssWeightDescriptorSelectsVariableFontWeight(int weight, int selectedWeight) {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(CreateDocument(weight));
        OfficeFontFace face = Assert.Single(rendered.Fonts.Faces);
        Assert.Contains("wght=" + selectedWeight, face.Program.Fingerprint, StringComparison.Ordinal);
        Assert.Equal(weight, face.Descriptor.Weight);
    }

    [Fact]
    public void ExplicitCallerWeightOverridesCssWeightDescriptor() {
        var options = new HtmlRenderOptions();
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 725F };
        OfficeFontFace face = Assert.Single(HtmlRenderEngine.Render(CreateDocument(700), options).Fonts.Faces);
        Assert.Contains("wght=725", face.Program.Fingerprint, StringComparison.Ordinal);
    }

    [Fact]
    public void CssWeightPreservesOtherCallerAxes() {
        var options = new HtmlRenderOptions();
        options.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wdth"] = 112F };
        OfficeFontFace face = Assert.Single(HtmlRenderEngine.Render(CreateDocument(600), options).Fonts.Faces);
        Assert.Contains("wght=600", face.Program.Fingerprint, StringComparison.Ordinal);
        Assert.Contains("wdth=112", face.Program.Fingerprint, StringComparison.Ordinal);
    }

    [Fact]
    public void AutomaticCssWeightProducesTheExplicitWeightPdf() {
        HtmlConversionDocument source = CreateDocument(700);
        var explicitOptions = new HtmlToPdfOptions();
        explicitOptions.Fonts.FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F };
        Assert.Equal(source.ToPdfBytes(new HtmlToPdfOptions()), source.ToPdfBytes(explicitOptions));
    }

    private static HtmlConversionDocument CreateDocument(int weight) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf"));
        return HtmlConversionDocument.Parse("<style>@font-face{font-family:Variable;src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(data) + "');font-weight:" + weight + "}p{font-family:Variable;font-weight:"
            + weight + "}</style><p>Variable weight measurement</p>");
    }
}
