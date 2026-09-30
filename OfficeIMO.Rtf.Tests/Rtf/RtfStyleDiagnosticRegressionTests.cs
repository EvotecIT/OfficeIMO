using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfStyleDiagnosticRegressionTests {
    [Theory]
    [InlineData(@"\emdash\endash\emspace\enspace\qmspace\bullet\lquote\rquote\ldblquote\rdblquote\ltrmark\rtlmark\zwj\zwnj", "—–\u2003\u2002\u2005•‘’“”\u200E\u200F\u200D\u200C")]
    [InlineData(@"\tab\line\par", null)]
    public void NameCharacterControlsArePreservedWithoutFalseOmission(string controls, string? expected) {
        expected ??= "\t" + Environment.NewLine + Environment.NewLine;
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi{\stylesheet{\s1 Named" + controls + @" Style;}}Plain\par}").Document;
        Assert.Equal("Named" + expected + "Style", document.Styles[0].Name);
        var result = document.ToRtfResult();
        result.RequireNoLoss();
        Assert.Equal(document.Styles[0].Name, RtfDocument.Read(result.Value).Document.Styles[0].Name);
        document.ToHtmlResult().RtfReport.RequireNoLoss();
    }

    [Theory]
    [InlineData(@"\strike", "RtfNormalizationStyleControlOmitted", "strike")]
    [InlineData(@"{\*\vendorstyle Private}", "RtfNormalizationStyleDestinationOmitted", "vendorstyle")]
    public void UnsupportedStyleSyntaxCannotPassStrictSemanticConversion(string syntax, string code, string feature) {
        string input = @"{\rtf1\ansi{\stylesheet{\s1" + syntax + @" Style;}}\pard\s1 Visible\par}";
        RtfReadResult read = RtfDocument.Read(input, new RtfReadOptions { WarnOnUnsupportedDestinations = false });
        Assert.Equal(input, read.ToRtfLossless());
        var result = read.Document.ToRtfResult(new RtfWriteOptions { MaterializeStyleFormatting = false });
        Assert.Contains(result.Report.Diagnostics, item => item.Code == code && item.Feature == feature);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        Assert.Contains(read.Document.Clone().ToRtfResult().Report.Diagnostics, item => item.Code == code);
        Assert.Contains(read.Document.ToHtmlResult().RtfReport.Diagnostics, item => item.Code == code);
        Assert.Contains(read.Document.ToMarkdownResult().Report.Diagnostics, item => item.Code == code);
        var word = read.Document.ToWordDocumentResult();
        using (word.Value) Assert.Contains(word.Report.Diagnostics, item => item.Code == code);
        Assert.Contains(read.Document.ToPdfDocumentResult().Report.Warnings, item => item.Code == code);
    }

    [Fact]
    public void CyclesAndMissingReferencesAreReportedOnceAcrossSemanticAdapters() {
        RtfDocument document = RtfDocument.Create();
        document.AddStyle(1, "First").BasedOnStyleId = 2;
        document.AddStyle(2, "Second").BasedOnStyleId = 1;
        document.AddParagraph("Cycle").StyleId = 1;
        RtfParagraph missing = document.AddNote(RtfNoteKind.Footnote).AddParagraph("Missing");
        missing.StyleId = 9;
        missing.AddText("Missing again").StyleId = 7;
        document.AddParagraph("Repeated missing").StyleId = 9;

        AssertDiagnostics(document.ToRtfResult().Report);
        AssertDiagnostics(document.ToHtmlResult().RtfReport);
        AssertDiagnostics(document.ToMarkdownResult().Report);
        var word = document.ToWordDocumentResult();
        using (word.Value) AssertDiagnostics(word.Report);
        var pdf = document.ToPdfDocumentResult();
        Assert.Contains(pdf.Report.Warnings, item => item.Code == "RtfStyleInheritanceCycle");
        Assert.Contains(pdf.Report.Warnings, item => item.Code == "RtfStyleReferenceMissing");

        static void AssertDiagnostics(RtfConversionReport report) {
            Assert.Single(report.Diagnostics, item => item.Code == "RtfStyleInheritanceCycle");
            Assert.Equal(2, report.Diagnostics.Count(item => item.Code == "RtfStyleReferenceMissing"));
            Assert.Throws<RtfConversionLossException>(() => report.RequireNoLoss());
        }
    }

    [Fact]
    public void OrdinaryImplicitStyleZeroAndEncodedStyleNamesAreLossFree() {
        const string input = @"{\rtf1\ansi{\stylesheet{\s1\uc1 Named\u233?;}}Plain\par}";
        RtfDocument document = RtfDocument.Read(input).Document;
        document.ToRtfResult().RequireNoLoss();
        document.ToHtmlResult().RtfReport.RequireNoLoss();
        Assert.Equal("Namedé", document.Styles[0].Name);
    }
}
