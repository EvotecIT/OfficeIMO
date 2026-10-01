using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfNormalizationReportTests {
    [Fact]
    public void Source_Omissions_Remain_Observable_After_Appending_The_Read_Model() {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi{\*\vendorpayload Hidden}Visible\par}").Document;
        RtfDocument destination = RtfDocument.Create();
        RtfDocumentMergeResult merged = destination.AppendDocument(source);
        Assert.Contains(merged.Report.Diagnostics, item => item.Feature == "vendorpayload" && item.Action == RtfConversionAction.Omitted);
        Assert.Contains(destination.ToRtfResult().Report.Diagnostics, item => item.Feature == "vendorpayload" && item.Action == RtfConversionAction.Omitted);
        Assert.Equal("Visible", destination.Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void Normalization_Reports_Read_Policy_Blocks_And_Invalidated_Alternate_Html() {
        RtfReadResult blocked = RtfDocument.Read(@"{\rtf1\ansi{\*\filetbl{\file\fid0 path}}Visible\par}", RtfReadOptions.CreateUntrustedProfile());
        Assert.Contains(blocked.Document.ToRtfResult().Report.Diagnostics, item => item.Code == "RTF106" && item.Action == RtfConversionAction.Blocked);
        RtfDocument html = RtfDocument.Read(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>Original HTML</p>}\htmlrtf1 Original fallback\par}").Document;
        Assert.DoesNotContain(html.ToRtfResult().Report.Diagnostics, item => item.Code == "RtfNormalizationHtmlOmitted");
        html.Paragraphs[0].Runs[0].Text = "Edited";
        Assert.Contains(html.ToRtfResult().Report.Diagnostics, item => item.Code == "RtfNormalizationHtmlOmitted" && item.Action == RtfConversionAction.Omitted);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Normalization_Reports_Unbound_Source_Content_Even_When_Read_Warnings_Are_Disabled(bool warn) {
        const string source = @"{\rtf1\ansi{\*\vendorpayload Secret metadata}\pard Visible\par}";
        RtfReadResult read = RtfDocument.Read(source, new RtfReadOptions { WarnOnUnsupportedDestinations = warn });
        Assert.Equal(source, read.ToRtfLossless());
        RtfConversionResult<string> output = read.Document.ToRtfResult();
        Assert.Contains(output.Report.Diagnostics, diagnostic => diagnostic.Feature == "vendorpayload" && diagnostic.Action == RtfConversionAction.Omitted);
        Assert.Throws<RtfConversionLossException>(() => output.RequireNoLoss());
        Assert.Equal("Visible", RtfDocument.Read(output.Value).Document.Paragraphs[0].ToPlainText());
        Assert.DoesNotContain("vendorpayload", output.Value, StringComparison.Ordinal);
        RtfConversionResult<string> cloned = read.Document.Clone().ToRtfResult();
        Assert.Equal(output.Report.Diagnostics.Select(item => item.Code), cloned.Report.Diagnostics.Select(item => item.Code));
    }

    [Fact]
    public void Fresh_Semantic_Documents_Have_Independent_Loss_Free_Normalization_Reports() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("First");
        var first = document.ToRtfResult();
        first.RequireNoLoss();
        first.Report.Add(RtfConversionSeverity.Warning, "Caller", "Caller loss", RtfConversionAction.Omitted);
        var second = document.ToRtfResult();
        second.RequireNoLoss();
        Assert.False(second.HasLoss);
        Assert.Equal("First", RtfDocument.Read(second.Value).Document.Paragraphs[0].ToPlainText());
    }
}
