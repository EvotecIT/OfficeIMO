using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOptionalBreakEncodingDiagnosticsTests {
    // Generated running content is encoded after document preflight. Exercise that encoder directly.
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void GeneratedEncoder_ReportsOriginalUtf16Offset(bool cff, bool named) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = File.ReadAllBytes(path!);
        var options = new PdfOptions();
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("OffsetFont", font));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "OffsetFont");
        Type writer = typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.PdfWriter", throwOnError: true)!;
        MethodInfo encoder = writer.GetMethods(BindingFlags.Static | BindingFlags.NonPublic).Single(method =>
            method.Name == "EncodeTextShowCommand" && method.GetParameters().Any(parameter => parameter.Name == "namedFont") == named);
        object?[] arguments = encoder.GetParameters().Select(parameter => parameter.HasDefaultValue ? parameter.DefaultValue : null).ToArray();
        arguments[0] = "AV\uFFFF";
        arguments[1] = PdfStandardFont.Helvetica;
        if (named) {
            Type face = typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.PdfNamedFontFace", true)!;
            arguments[2] = Activator.CreateInstance(face, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "OFFSETFONT", "OffsetFont", false, false }, null);
        }
        arguments[named ? 3 : 2] = options;
        var baselineInvocation = Assert.Throws<TargetInvocationException>(() => encoder.Invoke(null, arguments));
        var baseline = Assert.IsType<ArgumentException>(baselineInvocation.InnerException);
        arguments[0] = "AV\u200B\uFFFF";
        var invocation = Assert.Throws<TargetInvocationException>(() => encoder.Invoke(null, arguments));
        var error = Assert.IsType<ArgumentException>(invocation.InnerException);
        Assert.Equal(3, error.Data["index"]);
        Assert.Equal("U+FFFF", error.Data["codePoint"]);
        Assert.Equal(baseline.Data["code"], error.Data["code"]);
        Assert.Equal(baseline.Message.Replace("at index 2", "at index 3"), error.Message);
    }

    [Fact]
    public void EmbeddedFontDiagnosticAnnotations_PreserveCustomCodeAndMessage() {
        string? path = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(path);
        var original = Assert.Single(PdfTextDiagnostics.AnalyzeEmbeddedFontText("AV\uFFFF", File.ReadAllBytes(path!), "Identifier"));
        Assert.Equal("missing-embedded-font-glyph", original.Code);

        var annotated = original.WithPageNumber(2).WithTableCell(3, 4).WithFieldName("MailHost");

        Assert.Equal(original.Code, annotated.Code);
        Assert.Equal(original.Message, annotated.Message);
        Assert.Equal(original.Source, annotated.Source);
        Assert.Equal(original.Index, annotated.Index);
        Assert.Equal(original.CodePoint, annotated.CodePoint);
        Assert.Equal(original.Text, annotated.Text);
        Assert.Equal(2, annotated.PageNumber);
        Assert.Equal(3, annotated.TableRowIndex);
        Assert.Equal(4, annotated.TableColumnIndex);
        Assert.Equal("MailHost", annotated.FieldName);
        var warning = annotated.ToConversionWarning();
        Assert.Equal(original.Code, warning.Code);
        Assert.Equal(original.Message, warning.Message);
    }

    [Fact]
    public void RunningHeaderDiagnostic_PreservesFallbackDetailsAndPageLocation() {
        var options = new PdfOptions { ShowHeader = true, HeaderFormat = "AV\uFFFF" };
        options.RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Identifier fallback",
                OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'V'))
        }));

        var diagnostic = Assert.Single(PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("AV")).AnalyzeTextEncoding());

        Assert.Equal("missing-embedded-font-fallback-glyph", diagnostic.Code);
        Assert.Contains("Identifier fallback", diagnostic.Message);
        Assert.Equal(2, diagnostic.Index);
        Assert.Equal(1, diagnostic.PageNumber);
        Assert.Null(diagnostic.RunIndex);
        Assert.False(string.IsNullOrWhiteSpace(diagnostic.Location));
    }
}
