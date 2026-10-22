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
            method.Name == "EncodeTextShowCommand" && method.GetParameters().Length == (named ? 9 : 8));
        object?[] arguments = encoder.GetParameters().Select(parameter => parameter.HasDefaultValue ? parameter.DefaultValue : null).ToArray();
        arguments[0] = "AV\u200B\uFFFF";
        arguments[1] = PdfStandardFont.Helvetica;
        if (named) {
            Type face = typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.PdfNamedFontFace", true)!;
            arguments[2] = Activator.CreateInstance(face, BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "OFFSETFONT", "OffsetFont", false, false }, null);
        }
        arguments[named ? 3 : 2] = options;
        var invocation = Assert.Throws<TargetInvocationException>(() => encoder.Invoke(null, arguments));
        var error = Assert.IsType<ArgumentException>(invocation.InnerException);
        Assert.Equal(3, error.Data["index"]);
        Assert.Equal("U+FFFF", error.Data["codePoint"]);
        Assert.Contains("3", error.Message);
    }
}
